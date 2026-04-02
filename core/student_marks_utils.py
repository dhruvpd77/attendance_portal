"""Normalize marks from DB or spreadsheet cells for analytics and uploads."""


def normalize_student_mark(raw):
    """
    Return a float for numeric marks, or None if absent.

    Absent: null/blank, textual AB/ABS/ABSENT, dash placeholders, or unparseable values.
    """
    if raw is None:
        return None
    if isinstance(raw, str):
        s = ''.join(raw.split()).upper()
        if s in ('', '-', '--', 'AB', 'ABS', 'ABSENT'):
            return None
    try:
        return float(raw)
    except (TypeError, ValueError):
        return None
