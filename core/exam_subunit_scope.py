"""Sub-unit (dept_exam_child) scope for supervision: division codes + hub vs department phases."""
from __future__ import annotations

import re

from django.db.models import Q

from core.models import DepartmentExamProfile, SupervisionExamPhase


def canonical_division_code(raw) -> str:
    """Normalize stream/cell text to a stable form (e.g. SY1, sy_1 → SY_1)."""
    if raw is None:
        return ''
    s = str(raw).strip().upper().replace(' ', '_')
    s = re.sub(r'_+', '_', s).strip('_')
    compact = s.replace('_', '')
    m = re.match(r'^([A-Z]{1,10})(\d+)$', compact)
    if m:
        return f'{m.group(1)}_{m.group(2)}'
    return s


def division_code_variants(code: str) -> list[str]:
    """All strings that should match the same division as sub-unit code or sheet cell (SY1 ⇄ SY_1)."""
    if not (code or '').strip():
        return []
    c0 = canonical_division_code(code)
    compact = c0.replace('_', '')
    out = {c0, compact}
    m = re.match(r'^([A-Z]{1,10})_?(\d+)$', compact)
    if m:
        pref, num = m.group(1), m.group(2)
        out.add(f'{pref}_{num}')
        out.add(f'{pref}{num}')
    return sorted({x for x in out if x})


def division_codes_equivalent(a: str, b: str) -> bool:
    return bool(set(division_code_variants(a)) & set(division_code_variants(b)))


def division_code_match_q(subunit_or_cell_code: str) -> Q:
    variants = division_code_variants(subunit_or_cell_code)
    if not variants:
        return Q(pk__in=[])
    q = Q(division_code__iexact=variants[0])
    for v in variants[1:]:
        q |= Q(division_code__iexact=v)
    return q


def hub_user_id_for_subunit_parent(parent_prof: DepartmentExamProfile | None) -> int | None:
    """Hub user id whose institute-wide phases this sub-unit should see."""
    if not parent_prof:
        return None
    if parent_prof.is_hub_coordinator or not parent_prof.department_id:
        return parent_prof.user_id
    if parent_prof.invited_by_id:
        return parent_prof.invited_by_id
    return None


def subunit_supervision_duty_filter_q(prof: DepartmentExamProfile) -> Q:
    """SupervisionDuty queryset filter: correct division + department phases and/or parent hub phases."""
    code = (prof.subunit_code or '').strip().upper()
    div_q = division_code_match_q(code)
    parent_prof = prof.parent
    parts: list[Q] = []
    if prof.department_id:
        parts.append(Q(phase__department_id=prof.department_id) & div_q)
    hub_uid = hub_user_id_for_subunit_parent(parent_prof)
    if hub_uid and parent_prof and parent_prof.institute_semester_id:
        parts.append(
            Q(phase__hub_coordinator_id=hub_uid)
            & Q(phase__institute_semester_id=parent_prof.institute_semester_id)
            & div_q
        )
    if not parts:
        return Q(pk__in=[])
    out = parts[0]
    for p in parts[1:]:
        out |= p
    return out


def phases_for_subunit_prof(prof: DepartmentExamProfile):
    """Phase picker rows visible to this sub-unit (dept-local + hub phases)."""
    parent_prof = prof.parent
    parts: list[Q] = []
    if prof.department_id:
        parts.append(Q(department_id=prof.department_id))
    hub_uid = hub_user_id_for_subunit_parent(parent_prof)
    if hub_uid and parent_prof and parent_prof.institute_semester_id:
        parts.append(
            Q(hub_coordinator_id=hub_uid)
            & Q(institute_semester_id=parent_prof.institute_semester_id)
        )
    if not parts:
        return SupervisionExamPhase.objects.none()
    out = parts[0]
    for p in parts[1:]:
        out |= p
    return SupervisionExamPhase.objects.filter(out).order_by('name')


def duty_visible_to_subunit(prof: DepartmentExamProfile, duty) -> bool:
    from core.models import SupervisionDuty

    if not isinstance(duty, SupervisionDuty):
        return False
    if not division_codes_equivalent(prof.subunit_code or '', duty.division_code or ''):
        return False
    hub_uid = hub_user_id_for_subunit_parent(prof.parent)
    parent_prof = prof.parent
    if (
        hub_uid
        and duty.phase.hub_coordinator_id == hub_uid
        and parent_prof
        and duty.phase.institute_semester_id == parent_prof.institute_semester_id
    ):
        return True
    if prof.department_id and duty.phase.department_id == prof.department_id:
        return True
    return False
