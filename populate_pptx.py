#!/usr/bin/env python3
"""
populate_pptx.py — PDI report export (.xls) -> Voter Profile PPTX  (v4)

v4 adds partial-data tolerance and geographic pies.

Partial data: a slot whose universe isn't in the file is skipped, not
fatal. Skipped content is BLANKED — charts get their series cleared so
they render empty, text placeholders are emptied — never left holding
template values. Slides where every chart ended up blank are deleted.
Summary-table rows whose cells all ended up empty are deleted. Every run
ends with a RUN SUMMARY naming what was found, blanked, and removed.

Still fatal (no deck written): a file with no parsable table; none of
the configured universes present; a required standard segment row
missing; non-numeric or negative counts in a resolved column; a present
universe whose segment sums fail validation. A present universe whose
Total is 0 is demoted to "absent" with a warning (it behaves like
missing-by-design).

Geographic pies: rows beginning with "CA -" are classified as city/town,
unincorporated, council district, or supervisorial district; any other
CA-prefixed row is logged by name and never guessed. One pie per
category per slot: <Slot>_CityPie, <Slot>_CouncilPie, <Slot>_SupPie.
Slice values are count / that slot's Total row, half-up one decimal.

Usage:
    python3 populate_pptx.py "District Name.xls" [more ...] \
        [--template template_placeholders.pptx] \
        [--config universes.json] \
        [--mapping placeholder_mapping.json] \
        [--outdir ./output]
"""

import argparse
import json
import logging
import re
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from html.parser import HTMLParser
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData

log = logging.getLogger("populate_pptx")

# ----------------------------------------------------------------------------
# Label normalization
# ----------------------------------------------------------------------------

def norm(label):
    return re.sub(r"[^A-Z0-9]+", " ", label.upper()).strip()


SEG = {
    "TOTAL": norm("Total"),
    "DEM": norm("Dem"),
    "REP": norm("Rep"),
    "IND": norm("NPP/DTS (Only)"),
    "AFAM": norm("African-American"),
    "ASIAN": norm("Asian"),
    "LATINO": norm("Latino"),
    "EARLY": norm("Any Early VBM"),
    "LATE": norm("Always Late VBM"),
    "HOME": norm("Homeowner"),
}
AGE_BUCKETS = [
    (norm("Age 18-24"), "18-24"), (norm("Age 25-34"), "25-34"),
    (norm("Age 35-44"), "35-44"), (norm("Age 45-54"), "45-54"),
    (norm("Age 55-64"), "55-64"), (norm("Age 65"), "65+"),
]
REQUIRED = list(SEG.values()) + [a for a, _ in AGE_BUCKETS]

SLOT_IDS = ("total", "near", "far")
DEMO_KINDS = ("PartyPie", "AgeChart", "EthChart", "HomePie")
GEO_KINDS = ("CityPie", "CouncilPie", "SupPie")

TEXT_KEYS = ("label", "voters", "houses", "evbm", "alvbm", "vitm")


# ----------------------------------------------------------------------------
# Settings file (universes.json) — unchanged from v3
# ----------------------------------------------------------------------------

def load_config(path):
    try:
        raw = json.loads(path.read_text())
    except FileNotFoundError:
        raise ValueError(f"settings file not found: {path} — it should sit "
                         f"next to populate_pptx.py")
    except json.JSONDecodeError as e:
        raise ValueError(f"{path.name} is not valid JSON (line {e.lineno}, "
                         f"column {e.colno}): {e.msg}. Undo the last edit "
                         f"or restore from backup.")
    slots = raw.get("slots")
    if not isinstance(slots, list):
        raise ValueError(f"{path.name}: expected a top-level 'slots' list")
    by_id = {}
    for s in slots:
        sid = s.get("id")
        if sid not in SLOT_IDS:
            raise ValueError(f"{path.name}: slot id {sid!r} is not one of "
                             f"{SLOT_IDS}")
        if sid in by_id:
            raise ValueError(f"{path.name}: slot id {sid!r} appears twice")
        universe, label = s.get("universe"), s.get("label")
        if not universe or not isinstance(universe, str):
            raise ValueError(f"{path.name}: slot '{sid}' needs a 'universe' "
                             f"string")
        if not label or not isinstance(label, str):
            raise ValueError(f"{path.name}: slot '{sid}' needs a 'label' "
                             f"string")
        alts = s.get("alternates", [])
        if not isinstance(alts, list) or not all(isinstance(a, str) for a in alts):
            raise ValueError(f"{path.name}: slot '{sid}': 'alternates' must "
                             f"be a list of strings")
        by_id[sid] = {"candidates": [universe] + alts, "label": label}
    missing = [sid for sid in SLOT_IDS if sid not in by_id]
    if missing:
        raise ValueError(f"{path.name}: missing slot(s): {missing}")
    return by_id


# ----------------------------------------------------------------------------
# Geographic row classification — council BEFORE city, all case-insensitive
# ----------------------------------------------------------------------------

RE_GEO_CANDIDATE = re.compile(r"^\s*CA\s*-", re.I)
RE_COUNCIL = re.compile(
    r"^\s*CA\s*-\s*(?:City|Town)\s+of\s+(?P<city>.+?)\s+"
    r"Council\s+District\s+(?P<n>\d+)\s*$", re.I)
RE_SUP = re.compile(r"(?P<n>\d+)(?:st|nd|rd|th)\s+Supervisorial\s+District",
                    re.I)
RE_UNINC = re.compile(r"^\s*CA\s*-\s*Uninc\b", re.I)
RE_CITY = re.compile(r"^\s*CA\s*-\s*(?:City|Town)\s+of\s+(?P<name>.+?)\s*$",
                     re.I)


def classify_geo(raw):
    """-> (kind, sort_key, slice_label, parent_city); kind '?' means
    unrecognized. Council is tested before City so 'City of X Council
    District N' isn't swallowed by the city pattern."""
    m = RE_COUNCIL.match(raw)
    if m:
        n = int(m.group("n"))
        return ("council", n, f"D{n}", m.group("city").strip())
    m = RE_SUP.search(raw)
    if m:
        n = int(m.group("n"))
        return ("sup", n, f"SD {n}", None)
    if RE_UNINC.match(raw):
        return ("city", None, "Unincorporated", None)   # None sorts last
    m = RE_CITY.match(raw)
    if m:
        name = m.group("name").strip()
        return ("city", name.upper(), name, None)
    return ("?", None, raw, None)


# ----------------------------------------------------------------------------
# HTML table parsing (stdlib only)
# ----------------------------------------------------------------------------

class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell = [], None, None
        self._in_td = False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag == "td" and self._row is not None:
            self._in_td, self._cell = True, ""

    def handle_endtag(self, tag):
        if tag == "td" and self._in_td:
            self._row.append(self._cell.strip())
            self._in_td = False
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._in_td:
            self._cell += data


def district_from_filename(path):
    stem = re.sub(r"[_\s]+", " ", path.stem).strip()
    if re.match(r"(?i)^report\b", stem) or not stem:
        raise ValueError(
            f"'{path.name}' looks like a raw download. Rename the file to "
            f"the district name (e.g. 'Example Unified School District.xls') "
            f"— the filename becomes the client name on the slides.")
    return stem


def parse_report(path, config):
    """Returns (district, data, geo_rows, resolved, file_universes).
    data[norm_label][sid] = (voters, houses) for RESOLVED slots only.
    geo_rows: list of (raw_label, {sid: (voters, houses)}).
    resolved[sid] = universe code used, or None when absent (allowed)."""
    district = district_from_filename(path)

    parser = _TableParser()
    parser.feed(path.read_text(errors="replace"))
    rows = [r for r in parser.rows if r]
    if not rows:
        raise ValueError(f"{path.name}: no HTML table rows found — is this "
                         f"the PDI tabular export?")

    header, header_i = None, None
    for i, r in enumerate(rows):
        cells = [norm(c) for c in r]
        if "DESCRIPTION" in cells:
            header, header_i = cells, i
            break
    if header is None:
        raise ValueError(f"{path.name}: no header row containing "
                         f"'Description' found — the export format may have "
                         f"changed")
    file_universes = sorted({c[:-2].strip() for c in header
                             if c.endswith(" H")})

    cols, resolved = {}, {}
    for sid, slot in config.items():
        resolved[sid] = None
        for code in slot["candidates"]:
            v_col, h_col = norm(code), norm(code + " H")
            if v_col in header and h_col in header:
                resolved[sid] = code
                cols[sid] = (header.index(v_col), header.index(h_col))
                break
        if resolved[sid] is None:
            log.info("%s: slot '%s' has no data — tried %s; file contains: "
                     "%s", path.name, sid, slot["candidates"], file_universes)
        elif resolved[sid] != slot["candidates"][0]:
            log.warning("%s: slot '%s' using ALTERNATE universe '%s' "
                        "(primary '%s' not in file)", path.name, sid,
                        resolved[sid], slot["candidates"][0])
    if not cols:
        raise ValueError(
            f"{path.name}: none of the configured universes are in this "
            f"file. Expected any of "
            f"{[c for s in config.values() for c in s['candidates']]}; the "
            f"file contains: {file_universes}. Update universes.json or "
            f"check the export.")

    data, geo_rows, seen_geo = {}, [], set()
    for r in rows[header_i + 1:]:
        if len(r) != len(header):
            continue
        raw = r[0].strip()
        if not raw:
            continue
        entry = {}
        for sid, (vi, hi) in cols.items():
            try:
                v, h = int(r[vi].replace(",", "")), int(r[hi].replace(",", ""))
            except ValueError:
                raise ValueError(
                    f"{path.name}: non-numeric count in row '{raw}' (slot "
                    f"{sid}): voters={r[vi]!r} houses={r[hi]!r}")
            if v < 0 or h < 0:
                raise ValueError(f"{path.name}: negative count in row "
                                 f"'{raw}' (slot {sid})")
            entry[sid] = (v, h)
        if RE_GEO_CANDIDATE.match(raw):
            if raw not in seen_geo:
                seen_geo.add(raw)
                geo_rows.append((raw, entry))
        else:
            data.setdefault(norm(raw), entry)

    log.info("%s: district='%s', %d segment rows, %d geographic rows, "
             "universes: %s", path.name, district, len(data), len(geo_rows),
             {sid: resolved[sid] for sid in SLOT_IDS})
    return district, data, geo_rows, resolved, file_universes


# ----------------------------------------------------------------------------
# Percentages
# ----------------------------------------------------------------------------

def pct(count, total):
    return (Decimal(count) * 100 / Decimal(total)).quantize(
        Decimal("0.1"), rounding=ROUND_HALF_UP)


def fmt(d):
    s = f"{d}"
    return s[:-2] if s.endswith(".0") else s


# ----------------------------------------------------------------------------
# Validation (present slots only)
# ----------------------------------------------------------------------------

def validate(district, data, resolved, summary):
    missing = [lbl for lbl in REQUIRED if lbl not in data]
    if missing:
        raise ValueError(f"{district}: required segment row(s) not found in "
                         f"the file: {missing}")
    problems = []
    for sid in [s for s in SLOT_IDS if resolved[s]]:
        tag = f"[{sid}={resolved[sid]}]"
        tv = data[SEG["TOTAL"]][sid][0]
        if tv <= 0:
            log.warning("%s %s TOTAL voters is 0 — treating this slot as "
                        "having no data", district, tag)
            summary["zero_slots"].append(sid)
            resolved[sid] = None
            continue
        party = sum(data[SEG[k]][sid][0] for k in ("DEM", "REP", "IND"))
        if party > tv:
            problems.append(f"{tag} Dem+Rep+Independent = {party:,} exceeds "
                            f"TOTAL {tv:,}")
        ages = sum(data[lbl][sid][0] for lbl, _ in AGE_BUCKETS)
        if abs(ages - tv) > tv * 0.015:
            problems.append(f"{tag} age buckets sum to {ages:,}, expected "
                            f"~TOTAL {tv:,}")
        eth = sum(data[SEG[k]][sid][0] for k in ("AFAM", "ASIAN", "LATINO"))
        if eth > tv:
            problems.append(f"{tag} AfAm+Asian+Latino = {eth:,} exceeds "
                            f"TOTAL {tv:,}")
        vbm = data[SEG["EARLY"]][sid][0] + data[SEG["LATE"]][sid][0]
        if vbm > tv:
            problems.append(f"{tag} early + always-late VBM = {vbm:,} "
                            f"exceeds TOTAL {tv:,}")
        if data[SEG["HOME"]][sid][0] > tv:
            problems.append(f"{tag} Homeowner exceeds TOTAL")
    if problems:
        for p in problems:
            log.error("VALIDATION %s %s", district, p)
        raise ValueError(f"{district}: {len(problems)} validation failure(s) "
                         f"— no deck written.")
    if not any(resolved[s] for s in SLOT_IDS):
        raise ValueError(f"{district}: every present universe has a zero "
                         f"Total — nothing to build a deck from.")


# ----------------------------------------------------------------------------
# Geographic pies
# ----------------------------------------------------------------------------

def geo_chart_values(district, geo_rows, resolved, data, summary):
    """Returns (chart_values, blank_names) for the geographic pies."""
    city, council, sup = [], [], []
    for raw, entry in geo_rows:
        kind, key, label, parent = classify_geo(raw)
        if kind == "?":
            summary["unrecognized_geo"].append(raw)
            log.warning("%s: unrecognized geographic row (not charted): %r",
                        district, raw)
        elif kind == "city":
            city.append((key, label, entry))
        elif kind == "council":
            council.append((key, label, parent, entry))
        elif kind == "sup":
            sup.append((key, label, entry))

    # Ordering: cities alphabetical with Unincorporated last; districts
    # numeric. Council districts from multiple parent cities get
    # '<City> D<n>' labels (plain D<n> would collide), city-first order.
    city.sort(key=lambda t: (t[0] is None, t[0] or ""))
    parents = {p for _, _, p, _ in council}
    if len(parents) > 1:
        summary["council_multi_city"] = True
        council = [((p.upper(), n), f"{p} D{n}", p, e)
                   for n, _, p, e in council]
    council.sort(key=lambda t: t[0] if isinstance(t[0], tuple) else (0, t[0]))
    sup.sort(key=lambda t: t[0])

    groups = {"CityPie": [(lbl, e) for _, lbl, e in city],
              "CouncilPie": [(lbl, e) for _, lbl, _, e in council],
              "SupPie": [(lbl, e) for _, lbl, e in sup]}

    chart_values, blanks = {}, []
    present = [s for s in SLOT_IDS if resolved[s]]
    for kind, items in groups.items():
        summary["geo_counts"][kind] = len(items)
        for sid in SLOT_IDS:
            name = f"{sid.capitalize()}_{kind}"
            if sid not in present or not items:
                blanks.append(name)
                continue
            tv = data[SEG["TOTAL"]][sid][0]
            chart_values[name] = (
                [lbl for lbl, _ in items],
                [float(pct(e[sid][0], tv)) for _, e in items])
        # Cities + unincorporated should cover ~100% of each universe
        if kind == "CityPie" and items:
            for sid in present:
                tv = data[SEG["TOTAL"]][sid][0]
                cover = pct(sum(e[sid][0] for _, e in items), tv)
                if not (Decimal("98") <= cover <= Decimal("102")):
                    msg = (f"cities+unincorporated cover {fmt(cover)}% of "
                           f"the '{sid}' universe (expected ~100%) — a row "
                           f"may be unrecognized or the export incomplete")
                    summary["warnings"].append(msg)
                    log.warning("%s: %s", district, msg)
    return chart_values, blanks


# ----------------------------------------------------------------------------
# Demographics values
# ----------------------------------------------------------------------------

def compute_values(district, data, resolved, config, summary):
    text_values = {"district_name": district}
    chart_values, blanks = {}, []
    present = [s for s in SLOT_IDS if resolved[s]]
    total_present = "total" in present
    all_total = data[SEG["TOTAL"]]["total"][0] if total_present else None

    for sid in SLOT_IDS:
        prefix = sid.capitalize()
        if sid not in present:
            for key in TEXT_KEYS:
                text_values[f"{key}_{sid}"] = ""
            summary["empty_placeholders"] += [f"{k}_{sid}" for k in TEXT_KEYS]
            if sid != "total":
                text_values[f"share_{sid}"] = ""
                summary["empty_placeholders"].append(f"share_{sid}")
            blanks += [f"{prefix}_{k}" for k in DEMO_KINDS]
            continue

        tv, th = data[SEG["TOTAL"]][sid]
        get = lambda key: data[SEG[key]][sid][0]
        evbm, alvbm = pct(get("EARLY"), tv), pct(get("LATE"), tv)

        text_values[f"label_{sid}"] = config[sid]["label"]
        text_values[f"voters_{sid}"] = f"{tv:,}"
        text_values[f"houses_{sid}"] = f"{th:,}"
        text_values[f"evbm_{sid}"] = fmt(evbm)
        text_values[f"alvbm_{sid}"] = fmt(alvbm)
        text_values[f"vitm_{sid}"] = fmt(Decimal(100) - evbm - alvbm)
        if sid != "total":
            if total_present:
                text_values[f"share_{sid}"] = fmt(pct(tv, all_total))
            else:
                text_values[f"share_{sid}"] = ""
                summary["empty_placeholders"].append(f"share_{sid}")
                log.warning("%s: share_%s left empty — the total slot has "
                            "no data to divide by", district, sid)

        dem, rep, ind = get("DEM"), get("REP"), get("IND")
        chart_values[f"{prefix}_PartyPie"] = (
            ["DEM", "REP", "NPP", "Other"],
            [float(pct(c, tv)) for c in (dem, rep, ind, tv - dem - rep - ind)])
        chart_values[f"{prefix}_AgeChart"] = (
            [cat for _, cat in AGE_BUCKETS],
            [float(pct(data[lbl][sid][0], tv)) for lbl, _ in AGE_BUCKETS])
        afam, asian, latino = get("AFAM"), get("ASIAN"), get("LATINO")
        chart_values[f"{prefix}_EthChart"] = (
            ["African-American", "Asian", "Latino", "Caucasian or Other"],
            [float(pct(c, tv))
             for c in (afam, asian, latino, tv - afam - asian - latino)])
        home = get("HOME")
        chart_values[f"{prefix}_HomePie"] = (
            ["Renter", "Homeowner"],
            [float(pct(tv - home, tv)), float(pct(home, tv))])

    return text_values, chart_values, blanks


# ----------------------------------------------------------------------------
# PPTX population
# ----------------------------------------------------------------------------

PLACEHOLDER_RE = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")


def _replace_in_paragraph(para, values):
    replaced = 0
    while True:
        runs = para.runs
        full = "".join(r.text for r in runs)
        matches = [m for m in PLACEHOLDER_RE.finditer(full)
                   if m.group(1) in values]
        if not matches:
            return replaced
        m = matches[0]
        start, end = m.span()
        pos = 0
        first = last = None
        first_off = last_off = 0
        for r in runs:
            r_start, r_end = pos, pos + len(r.text)
            if r_start <= start < r_end:
                first, first_off = r, start - r_start
            if r_start < end <= r_end:
                last, last_off = r, end - r_start
            pos = r_end
        if first is last:
            first.text = (first.text[:first_off] + str(values[m.group(1)])
                          + first.text[last_off:])
        else:
            first.text = first.text[:first_off] + str(values[m.group(1)])
            seen = False
            for r in runs:
                if r is first:
                    seen = True
                    continue
                if not seen:
                    continue
                if r is last:
                    r.text = r.text[last_off:]
                    break
                r.text = ""
        replaced += 1


def _iter_text_frames(shapes):
    for sh in shapes:
        if sh.shape_type == 6:
            yield from _iter_text_frames(sh.shapes)
            continue
        if sh.has_text_frame:
            yield sh.text_frame
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    yield cell.text_frame


def _blank_chart(chart):
    """Clear the series so the chart renders empty."""
    cd = CategoryChartData()
    cd.categories = [""]
    cd.add_series(chart.plots[0].series[0].name or "Series 1", (None,))
    chart.replace_data(cd)


def _delete_slide(prs, slide):
    id_lst = prs.slides._sldIdLst
    for sldId in list(id_lst):
        if prs.part.related_part(sldId.rId) is slide.part:
            prs.part.drop_rel(sldId.rId)
            id_lst.remove(sldId)
            return


def populate(template_path, out_path, text_values, chart_values, blank_names,
             summary):
    prs = Presentation(template_path)
    blank_set = set(blank_names)
    updates = {name: 0 for name in chart_values}
    blanked = {name: 0 for name in blank_set}
    slide_info = []

    for si, slide in enumerate(prs.slides, 1):
        # Remove table body rows whose every placeholder resolved to empty
        # (static scaffolding like 'Voters / Households' or '100%' doesn't
        # save a row that has no data behind it). Runs before replacement.
        for sh in slide.shapes:
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in list(sh.table.rows)[1:]:
                    names = {n for c in row.cells
                             for n in PLACEHOLDER_RE.findall(c.text)}
                    known = {n for n in names if n in text_values}
                    if known and all(str(text_values[n]) == ""
                                     for n in known):
                        row._tr.getparent().remove(row._tr)
                        summary["table_rows_removed"] += 1
        for tf in _iter_text_frames(slide.shapes):
            for para in tf.paragraphs:
                _replace_in_paragraph(para, text_values)
        n_data = n_blank = 0
        title = ""
        for sh in slide.shapes:
            if sh.has_text_frame and "Profile" in sh.text_frame.text:
                title = sh.text_frame.text.replace("\n", " ").strip()
            if not sh.has_chart:
                continue
            if sh.name in chart_values:
                cats, vals = chart_values[sh.name]
                cd = CategoryChartData()
                cd.categories = cats
                cd.add_series(sh.chart.plots[0].series[0].name or "Series 1",
                              vals)
                sh.chart.replace_data(cd)
                updates[sh.name] += 1
                n_data += 1
            elif sh.name in blank_set:
                _blank_chart(sh.chart)
                blanked[sh.name] += 1
                n_blank += 1
            else:
                log.warning("slide %d: chart shape %r is not a recognized "
                            "chart name — NOT updated (mis-named copy?)",
                            si, sh.name)
        slide_info.append((slide, si, n_data, n_blank, title))

    for name, count in sorted(updates.items()):
        if count == 0:
            log.warning("expected chart %r was not found anywhere in the "
                        "template", name)
        else:
            cats, vals = chart_values[name]
            log.info("chart %-17s x%d <- %s", name, count,
                     dict(zip(cats, vals)))
    for name, count in sorted(blanked.items()):
        if count:
            summary["blanked_charts"].append(f"{name} x{count}")

    # Delete slides where every chart was blanked (chartless slides stay)
    for slide, si, n_data, n_blank, title in slide_info:
        if n_blank > 0 and n_data == 0:
            _delete_slide(prs, slide)
            summary["slides_removed"].append(
                f"#{si}" + (f" ({title})" if title else ""))
            log.info("slide %d removed — every chart on it had no data", si)

    leftovers = set()
    for slide in prs.slides:
        for tf in _iter_text_frames(slide.shapes):
            leftovers.update(PLACEHOLDER_RE.findall(tf.text))
    if leftovers:
        raise ValueError(f"unknown placeholders in template (typo?): "
                         f"{sorted(leftovers)} — no deck written.")

    prs.save(out_path)


# ----------------------------------------------------------------------------
# Run summary
# ----------------------------------------------------------------------------

def print_summary(district, resolved, config, summary, out_path):
    lines = [f"RUN SUMMARY — {district}"]
    for sid in SLOT_IDS:
        if resolved[sid]:
            lines.append(f"  slot {sid:<5} FILLED from universe "
                         f"'{resolved[sid]}'")
        else:
            why = ("universe Total was 0" if sid in summary["zero_slots"]
                   else f"none of {config[sid]['candidates']} in file")
            lines.append(f"  slot {sid:<5} EMPTY  — {why}; its charts "
                         f"blanked, its text cleared")
    g = summary["geo_counts"]
    geo_bits = []
    for kind, noun in (("CityPie", "cities"),
                       ("CouncilPie", "council districts"),
                       ("SupPie", "supervisorial districts")):
        n = g.get(kind, 0)
        geo_bits.append(f"{noun}: {n if n else 'none in file'}")
    lines.append("  geographic slices — " + "; ".join(geo_bits))
    if summary.get("council_multi_city"):
        lines.append("  council districts span multiple cities — slices "
                     "labeled '<City> D<n>' to avoid collisions")
    if summary["unrecognized_geo"]:
        lines.append(f"  unrecognized geographic rows (not charted): "
                     f"{summary['unrecognized_geo']}")
    if summary["blanked_charts"]:
        lines.append(f"  charts blanked: "
                     f"{', '.join(summary['blanked_charts'])}")
    if summary["slides_removed"]:
        lines.append(f"  slides removed (all charts empty): "
                     f"{', '.join(summary['slides_removed'])}")
    if summary["table_rows_removed"]:
        lines.append(f"  empty summary-table rows removed: "
                     f"{summary['table_rows_removed']}")
    if summary["empty_placeholders"]:
        lines.append(f"  placeholders emptied: "
                     f"{sorted(set(summary['empty_placeholders']))}")
    for w in summary["warnings"]:
        lines.append(f"  WARNING: {w}")
    lines.append(f"  deck -> {out_path}")
    for ln in lines:
        log.info(ln)


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def process_file(path, template, config, outdir):
    summary = {"zero_slots": [], "unrecognized_geo": [], "geo_counts": {},
               "blanked_charts": [], "slides_removed": [],
               "table_rows_removed": 0, "empty_placeholders": [],
               "warnings": []}
    district, data, geo_rows, resolved, _ = parse_report(path, config)
    validate(district, data, resolved, summary)
    text_values, chart_values, blanks = compute_values(
        district, data, resolved, config, summary)
    geo_values, geo_blanks = geo_chart_values(
        district, geo_rows, resolved, data, summary)
    chart_values.update(geo_values)
    blanks += geo_blanks

    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    safe = re.sub(r"[^A-Za-z0-9]+", "_", district).strip("_")
    out_path = outdir / f"{safe}_Voter_Profile_{stamp}.pptx"
    populate(template, out_path, text_values, chart_values, blanks, summary)
    print_summary(district, resolved, config, summary, out_path)
    print(out_path)
    return out_path


def main(argv=None):
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+", type=Path)
    ap.add_argument("--template", type=Path,
                    default=Path("template_placeholders.pptx"))
    ap.add_argument("--config", type=Path, default=Path("universes.json"))
    ap.add_argument("--mapping", type=Path,
                    default=Path("placeholder_mapping.json"))
    ap.add_argument("--outdir", type=Path, default=Path("output"))
    args = ap.parse_args(argv)

    args.outdir.mkdir(parents=True, exist_ok=True)
    if not log.handlers and not logging.getLogger().handlers:
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(message)s",
            handlers=[logging.StreamHandler(),
                      logging.FileHandler(args.outdir / "populate_pptx.log")])

    if not args.template.exists():
        log.error("template not found: %s", args.template)
        return 1
    json.loads(args.mapping.read_text())
    try:
        config = load_config(args.config)
    except ValueError as e:
        log.error("SETTINGS ERROR: %s", e)
        return 1
    log.info("universes.json: %s", {sid: config[sid]["candidates"]
                                    for sid in SLOT_IDS})

    failures = 0
    for path in args.files:
        try:
            process_file(path, args.template, config, args.outdir)
        except Exception as e:
            log.error("FAILED %s: %s", path.name, e)
            failures += 1
    if failures:
        log.error("%d of %d file(s) failed", failures, len(args.files))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
