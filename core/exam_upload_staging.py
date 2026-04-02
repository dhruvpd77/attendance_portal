"""Two-step Excel import: parse into session first; user commits to DB explicitly."""
from __future__ import annotations

from datetime import date, datetime

STAGING_VER = 1


def _key(user_id: int, kind: str, phase_id: int) -> str:
    return f'eu_stage_{kind}_{user_id}_{phase_id}'


def clear_staging(request, kind: str, phase_id: int) -> None:
    request.session.pop(_key(request.user.pk, kind, phase_id), None)


def _parse_date(val) -> date | None:
    if val is None or val == '':
        return None
    if isinstance(val, date):
        return val
    return datetime.strptime(str(val)[:10], '%Y-%m-%d').date()


# --- Paper setting ---


def paper_setting_stage_put(request, phase_id: int, rows: list[dict], n_unmatched: int) -> None:
    ser = []
    for r in rows:
        dd = r.get('duty_date')
        dl = r.get('deadline_date')
        ser.append(
            {
                'duty_date': dd.isoformat() if dd else None,
                'deadline_date': dl.isoformat() if dl else None,
                'subject_name': r.get('subject_name') or '',
                'faculty_initial': r.get('faculty_initial') or '',
                'notes': r.get('notes') or '',
            }
        )
    request.session[_key(request.user.pk, 'paper_setting', phase_id)] = {
        'v': STAGING_VER,
        'rows': ser,
        'n_unmatched': n_unmatched,
    }
    request.session.modified = True


def paper_setting_stage_get(request, phase_id: int) -> dict | None:
    blob = request.session.get(_key(request.user.pk, 'paper_setting', phase_id))
    if not blob or blob.get('v') != STAGING_VER:
        return None
    return blob


def paper_setting_stage_deserialize_rows(blob: dict) -> list[dict]:
    out: list[dict] = []
    for r in blob.get('rows') or []:
        dd = _parse_date(r.get('duty_date'))
        dl = _parse_date(r.get('deadline_date'))
        out.append(
            {
                'duty_date': dd,
                'deadline_date': dl,
                'subject_name': r.get('subject_name') or '',
                'faculty_initial': r.get('faculty_initial') or '',
                'notes': r.get('notes') or '',
            }
        )
    return out


# --- Paper checking ---


def paper_checking_stage_put(request, phase_id: int, rows: list[dict], n_unmatched: int) -> None:
    ser = []
    for r in rows:
        ex = r.get('exam_date')
        ser.append(
            {
                'evaluator_initial': r.get('evaluator_initial') or '',
                'exam_date': ex.isoformat() if ex else None,
                'subject_name': r.get('subject_name') or '',
                'total_students': int(r.get('total_students') or 0),
                'allocations': [
                    {
                        'dept_code': (a.get('dept_code') or '').strip(),
                        'block_range': (a.get('block_range') or '').strip(),
                    }
                    for a in (r.get('allocations') or [])
                ],
            }
        )
    request.session[_key(request.user.pk, 'paper_checking', phase_id)] = {
        'v': STAGING_VER,
        'rows': ser,
        'n_unmatched': n_unmatched,
    }
    request.session.modified = True


def paper_checking_stage_get(request, phase_id: int) -> dict | None:
    blob = request.session.get(_key(request.user.pk, 'paper_checking', phase_id))
    if not blob or blob.get('v') != STAGING_VER:
        return None
    return blob


def paper_checking_stage_deserialize_rows(blob: dict) -> list[dict]:
    out: list[dict] = []
    for r in blob.get('rows') or []:
        out.append(
            {
                'evaluator_initial': r.get('evaluator_initial') or '',
                'exam_date': _parse_date(r.get('exam_date')),
                'subject_name': r.get('subject_name') or '',
                'total_students': int(r.get('total_students') or 0),
                'allocations': list(r.get('allocations') or []),
            }
        )
    return out


# --- Supervision ---


def supervision_stage_put(request, phase_id: int, rows: list[dict], n_unmatched: int) -> None:
    ser = []
    for r in rows:
        sd = r.get('supervision_date')
        ser.append(
            {
                'faculty_name': r.get('faculty_name') or '',
                'faculty_initial': r.get('faculty_initial') or '',
                'supervision_date': sd.isoformat() if sd else None,
                'time_slot': r.get('time_slot') or '',
                'subject_name': r.get('subject_name') or '',
                'division_code': r.get('division_code') or '',
            }
        )
    request.session[_key(request.user.pk, 'supervision', phase_id)] = {
        'v': STAGING_VER,
        'rows': ser,
        'n_unmatched': n_unmatched,
    }
    request.session.modified = True


def supervision_stage_get(request, phase_id: int) -> dict | None:
    blob = request.session.get(_key(request.user.pk, 'supervision', phase_id))
    if not blob or blob.get('v') != STAGING_VER:
        return None
    return blob


def supervision_stage_deserialize_rows(blob: dict) -> list[dict]:
    out: list[dict] = []
    for r in blob.get('rows') or []:
        out.append(
            {
                'faculty_name': r.get('faculty_name') or '',
                'faculty_initial': r.get('faculty_initial') or '',
                'supervision_date': _parse_date(r.get('supervision_date')),
                'time_slot': r.get('time_slot') or '',
                'subject_name': r.get('subject_name') or '',
                'division_code': r.get('division_code') or '',
            }
        )
    return out
