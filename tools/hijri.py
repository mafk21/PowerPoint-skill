#!/usr/bin/env python3
"""Hijri date helper — Gregorian -> arithmetical (tabular) Hijri.

Same algorithm the generator uses for the {hijri} spec token. Tabular
Hijri can differ by ±1 day from sighting-based official calendars
(Umm al-Qura) — fine for deck covers and planning, not for legal dates.

Usage:
  python tools/hijri.py                 # today
  python tools/hijri.py 2027-03-15      # a specific Gregorian date
"""
import datetime
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "generator"))
from generate_pptx import HIJRI_MONTHS, gregorian_to_hijri  # noqa: E402

INDIC = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def main(argv):
    if len(argv) > 1:
        try:
            dt = datetime.date.fromisoformat(argv[1])
        except ValueError:
            print("expected YYYY-MM-DD")
            return 2
    else:
        dt = datetime.date.today()
    hy, hm, hd = gregorian_to_hijri(dt.year, dt.month, dt.day)
    western = f"{hd} {HIJRI_MONTHS[hm - 1]} {hy}هـ"
    print(f"{dt.isoformat()}  ->  {western}  ({western.translate(INDIC)})")
    print(f"numeric: {hy:04d}-{hm:02d}-{hd:02d} AH  (tabular; official may "
          f"differ by a day)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
