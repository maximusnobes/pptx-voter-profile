#!/usr/bin/env python3
"""
Generate a fake PDI-style tabular export so the pipeline can be tried
without real voter data. Every number here is synthetic.

    python3 examples/make_sample_export.py
    python3 populate_pptx.py "examples/Springfield Unified School District.xls"
"""

import random
from pathlib import Path

random.seed(7)

# universe code -> total registered voters in that universe
UNIVERSES = {"TOTAL": 84_000, "26G5": 41_500, "26P9A": 58_200}

# share of the universe's voters in each segment row
SHARES = {
    "Dem": 0.41, "Rep": 0.29, "NPP/DTS (Only)": 0.24,
    "African-American": 0.04, "Asian": 0.11, "Latino": 0.33,
    "Any Early VBM": 0.38, "Always Late VBM": 0.17, "Homeowner": 0.58,
}
AGES = {"Age 18-24": 0.09, "Age 25-34": 0.15, "Age 35-44": 0.16,
        "Age 45-54": 0.16, "Age 55-64": 0.18, "Age 65": 0.26}
GEO = {
    "CA - City Of Springfield": 0.52,
    "CA - City Of Shelbyville": 0.31,
    "CA - Uninc - Example County": 0.17,
    "CA - Example County - 1st Supervisorial District": 0.44,
    "CA - Example County - 2nd Supervisorial District": 0.56,
    "CA - City Of Springfield Council District 1": 0.18,
    "CA - City Of Springfield Council District 2": 0.20,
    "CA - City Of Springfield Council District 3": 0.14,
}


def jitter(x, code):
    # older, more partisan electorates in the narrower universes
    return x * (1 + random.uniform(-0.08, 0.08)) if code != "TOTAL" else x


def row(label, shares_by_code):
    cells = [label]
    for code, total in UNIVERSES.items():
        voters = round(total * shares_by_code[code])
        cells += [f"{voters:,}", f"{round(voters * 0.62):,}"]
    return cells


def main():
    header = ["Description"]
    for code in UNIVERSES:
        header += [code, f"{code} H"]

    rows = [row("Total", {c: 1.0 for c in UNIVERSES})]
    for label, s in SHARES.items():
        rows.append(row(label, {c: jitter(s, c) for c in UNIVERSES}))
    for label, s in AGES.items():
        rows.append(row(label, {c: s for c in UNIVERSES}))
    for label, s in GEO.items():
        rows.append(row(label, {c: s for c in UNIVERSES}))

    def tr(cells):
        return "<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>"

    html = ("<html><body><table>\n" + tr(header) + "\n"
            + "\n".join(tr(r) for r in rows) + "\n</table></body></html>\n")
    out = Path(__file__).parent / "Springfield Unified School District.xls"
    out.write_text(html)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
