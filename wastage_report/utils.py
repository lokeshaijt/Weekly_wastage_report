import re
from datetime import date, timedelta

# Week 1 = the Sunday-starting week that contains 04-Jan-2026 (Sun 28-Dec-2025 .. Sat 03-Jan-2026).
# Same numbering as the PM report app, so week numbers match the "Planned in NN" columns of the source system.
WEEK_ANCHOR = date(2025, 12, 28)


def norm_disp(s):
    """Collapse whitespace / non-breaking spaces, keep case (for display)."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip()


def norm(s):
    """Matching key: normalised and upper-cased."""
    return norm_disp(s).upper()


def week_number_for(d):
    return (d - WEEK_ANCHOR).days // 7 + 1


def week_bounds(n):
    start = WEEK_ANCHOR + timedelta(days=(n - 1) * 7)
    return start, start + timedelta(days=6)
