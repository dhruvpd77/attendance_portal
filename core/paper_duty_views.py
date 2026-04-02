"""Paper checking & paper setting phases, Excel uploads, dashboards (exam section + dept coordinators)."""
from __future__ import annotations

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import IntegrityError, transaction
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_http_methods

from core.faculty_scope import faculty_has_department_access
from core.exam_coordination_views import (
    _dept_child_only,
    _dept_parent_only,
    _exam_section_portal_access,
    _is_hub_coordinator,
    _child_profile,
    _parent_profile,
)
from core.paper_checking_credits import credit_for_completion_request
from core.paper_duty_scope import (
    checking_phases_institute_for_request,
    history_paper_check_completion_requests_for_child,
    history_paper_setting_completion_requests_for_child,
    hub_managed_department_ids,
    paper_checking_duties_for_child_prof,
    paper_setting_duties_for_child_prof,
    pending_paper_check_completion_requests_for_child,
    pending_paper_setting_completion_requests_for_child,
    setting_phases_institute_for_request,
)
from core.paper_setting_credits import credit_for_paper_setting_request
from core.models import (
    Department,
    DepartmentExamCreditRule,
    Faculty,
    PaperCheckingAdjustedShare,
    PaperCheckingAllocation,
    PaperCheckingCompletionRequest,
    PaperCheckingDuty,
    PaperCheckingPhase,
    PaperCheckingSubjectCredit,
    PaperSettingCompletionRequest,
    PaperSettingDuty,
    PaperSettingPhase,
)
from core.paper_checking_excel import (
    default_checking_deadline,
    parse_paper_checking_workbook,
    resolve_department_from_sheet_code,
)
from core.paper_setting_excel import parse_paper_setting_workbook
from core.semester_scope import (
    departments_for_exam_coordination_request,
    institute_semester_for_exam_portal,
    is_exam_section_operator,
)
from core.exam_upload_staging import (
    clear_staging,
    paper_checking_stage_deserialize_rows,
    paper_checking_stage_get,
    paper_checking_stage_put,
    paper_setting_stage_deserialize_rows,
    paper_setting_stage_get,
    paper_setting_stage_put,
)
from core.supervision_excel import match_faculty_for_department, match_faculty_global


def _institute_paper_access(request):
    return _exam_section_portal_access(request)


def _paper_coord_access(request) -> bool:
    return _institute_paper_access(request) or _dept_parent_only(request)


def _checking_phases_for_user(request):
    if _institute_paper_access(request):
        return checking_phases_institute_for_request(request)
    if not _dept_parent_only(request):
        return PaperCheckingPhase.objects.none()
    prof = _parent_profile(request)
    if not prof:
        return PaperCheckingPhase.objects.none()
    if _is_hub_coordinator(prof):
        sem = institute_semester_for_exam_portal(request)
        if not sem:
            return PaperCheckingPhase.objects.none()
        return PaperCheckingPhase.objects.filter(
            hub_coordinator=request.user, institute_semester=sem
        ).order_by('name')
    return PaperCheckingPhase.objects.filter(department=prof.department).order_by('name')


def _setting_phases_for_user(request):
    if _institute_paper_access(request):
        return setting_phases_institute_for_request(request)
    if not _dept_parent_only(request):
        return PaperSettingPhase.objects.none()
    prof = _parent_profile(request)
    if not prof:
        return PaperSettingPhase.objects.none()
    if _is_hub_coordinator(prof):
        sem = institute_semester_for_exam_portal(request)
        if not sem:
            return PaperSettingPhase.objects.none()
        return PaperSettingPhase.objects.filter(
            hub_coordinator=request.user, institute_semester=sem
        ).order_by('name')
    return PaperSettingPhase.objects.filter(department=prof.department).order_by('name')


def _get_checking_phase(request, phase_id) -> PaperCheckingPhase:
    qs = _checking_phases_for_user(request)
    return get_object_or_404(qs.filter(pk=phase_id))


def _get_setting_phase(request, phase_id) -> PaperSettingPhase:
    qs = _setting_phases_for_user(request)
    return get_object_or_404(qs.filter(pk=phase_id))


def _faculty_match_checking(phase: PaperCheckingPhase, coordinator_dept: Department | None, name: str, initial: str):
    if phase.institute_scope or phase.hub_coordinator_id:
        return match_faculty_global(name, initial)
    if coordinator_dept:
        return match_faculty_for_department(coordinator_dept, name, initial)
    if phase.department_id:
        d = Department.objects.filter(pk=phase.department_id).first()
        if d:
            return match_faculty_for_department(d, name, initial)
    return match_faculty_global(name, initial)


def _faculty_match_setting(phase: PaperSettingPhase, coordinator_dept: Department | None, name: str, initial: str):
    if phase.institute_scope or phase.hub_coordinator_id:
        return match_faculty_global(name, initial)
    if coordinator_dept:
        return match_faculty_for_department(coordinator_dept, name, initial)
    if phase.department_id:
        d = Department.objects.filter(pk=phase.department_id).first()
        if d:
            return match_faculty_for_department(d, name, initial)
    return match_faculty_global(name, initial)


def _commit_paper_checking_rows(
    phase: PaperCheckingPhase, coordinator_dept: Department | None, rows: list[dict]
) -> int:
    with transaction.atomic():
        phase.duties.all().delete()
        n = 0
        for row in rows:
            exam_date = row.get('exam_date')
            if not exam_date:
                continue
            fac = _faculty_match_checking(
                phase, coordinator_dept, '', row.get('evaluator_initial') or ''
            )
            deadline = default_checking_deadline(exam_date)
            duty = PaperCheckingDuty.objects.create(
                phase=phase,
                faculty=fac,
                faculty_name_raw='',
                evaluator_short_raw=row.get('evaluator_initial') or '',
                exam_date=exam_date,
                subject_name=row.get('subject_name') or '',
                total_students=int(row.get('total_students') or 0),
                deadline_date=deadline,
            )
            for a in row.get('allocations') or []:
                dept_obj = resolve_department_from_sheet_code(a.get('dept_code') or '')
                PaperCheckingAllocation.objects.create(
                    duty=duty,
                    department=dept_obj,
                    department_code_raw=(a.get('dept_code') or '').strip(),
                    block_range=(a.get('block_range') or '').strip(),
                )
            n += 1
        return n


def _commit_paper_setting_rows(
    phase: PaperSettingPhase, coordinator_dept: Department | None, rows: list[dict]
) -> int:
    with transaction.atomic():
        phase.duties.all().delete()
        n = 0
        for row in rows:
            fac = _faculty_match_setting(
                phase, coordinator_dept, '', row.get('faculty_initial') or ''
            )
            dl = row.get('deadline_date') or row.get('duty_date')
            PaperSettingDuty.objects.create(
                phase=phase,
                faculty=fac,
                faculty_name_raw='',
                faculty_short_raw=row.get('faculty_initial') or '',
                duty_date=row.get('duty_date'),
                deadline_date=dl,
                subject_name=row.get('subject_name') or '',
                notes=row.get('notes') or '',
            )
            n += 1
        return n


@login_required
def paper_checking_dashboard(request):
    if _dept_child_only(request):
        dept_name = ''
        prof = _child_profile(request)
        if not prof:
            messages.error(request, 'No profile.')
            return redirect('accounts:role_redirect')
        if prof.department_id:
            dept_name = prof.department.name
        base = paper_checking_duties_for_child_prof(prof)
        phase_id_set = set(base.values_list('phase_id', flat=True))
        phases = PaperCheckingPhase.objects.filter(id__in=phase_id_set).order_by('name')
        raw_phase = (request.GET.get('phase_id') or '').strip()
        raw_subject = (request.GET.get('subject') or '').strip()
        selected_phase_id = None
        if raw_phase:
            try:
                selected_phase_id = int(raw_phase)
            except (TypeError, ValueError):
                selected_phase_id = None
        duties_qs = base
        subject_choices: list[str] = []
        if selected_phase_id and selected_phase_id in phase_id_set:
            duties_for_phase = base.filter(phase_id=selected_phase_id)
            subject_choices = sorted(
                {s for s in duties_for_phase.values_list('subject_name', flat=True) if (s or '').strip()},
                key=lambda x: x.upper(),
            )
            duties_qs = duties_for_phase
            if raw_subject:
                duties_qs = duties_qs.filter(subject_name__iexact=raw_subject)
            duties_qs = duties_qs.prefetch_related(
                'adjusted_shares__faculty', 'faculty', 'allocations__department'
            )
        else:
            duties_qs = base.none()
        pending = list(pending_paper_check_completion_requests_for_child(prof))
        pending_completions_enriched = [
            {'r': r, 'credit': credit_for_completion_request(r)} for r in pending
        ]
        history_completions = list(history_paper_check_completion_requests_for_child(prof))
        history_completions_enriched = []
        for r in history_completions:
            cr = (
                credit_for_completion_request(r)
                if r.status == PaperCheckingCompletionRequest.APPROVED
                else None
            )
            history_completions_enriched.append({'r': r, 'credit': cr})
        dept_faculty = []
        if prof.department_id:
            dept_faculty = list(
                Faculty.objects.filter(department_id=prof.department_id).order_by('full_name')
            )
        return render(
            request,
            'core/exam_paper/checking_child_view.html',
            {
                'duties': duties_qs,
                'profile': prof,
                'subunit': (prof.subunit_code or '').strip().upper(),
                'department_name': dept_name,
                'phases': phases,
                'selected_phase_id': selected_phase_id,
                'subject_choices': subject_choices,
                'selected_subject': raw_subject,
                'pending_completions': pending,
                'pending_completions_enriched': pending_completions_enriched,
                'history_completions': history_completions,
                'history_completions_enriched': history_completions_enriched,
                'dept_faculty': dept_faculty,
            },
        )
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phases = _checking_phases_for_user(request)
    institute = _institute_paper_access(request)
    return render(
        request,
        'core/exam_paper/checking_dashboard.html',
        {
            'phases': phases,
            'institute_scope_active': institute,
            'page_kind': 'checking',
        },
    )


@login_required
def paper_setting_dashboard(request):
    if _dept_child_only(request):
        prof = _child_profile(request)
        if not prof:
            messages.error(request, 'No profile.')
            return redirect('accounts:role_redirect')
        base = paper_setting_duties_for_child_prof(prof)
        phase_id_set = set(base.values_list('phase_id', flat=True))
        phases = PaperSettingPhase.objects.filter(id__in=phase_id_set).order_by('name')
        raw_phase = (request.GET.get('phase_id') or '').strip()
        selected_phase_id = None
        if raw_phase:
            try:
                selected_phase_id = int(raw_phase)
            except (TypeError, ValueError):
                selected_phase_id = None
        if selected_phase_id and selected_phase_id not in phase_id_set:
            selected_phase_id = None
        duties = base.filter(phase_id=selected_phase_id) if selected_phase_id else base
        dept_name = prof.department.name if prof.department_id else ''
        pending = list(pending_paper_setting_completion_requests_for_child(prof))
        if selected_phase_id:
            pending = [r for r in pending if r.duty.phase_id == selected_phase_id]
        pending_enriched = [
            {'r': r, 'credit': credit_for_paper_setting_request(r)} for r in pending
        ]
        history = list(history_paper_setting_completion_requests_for_child(prof))
        if selected_phase_id:
            history = [r for r in history if r.duty.phase_id == selected_phase_id]
        history_enriched = []
        for r in history:
            cr = (
                credit_for_paper_setting_request(r)
                if r.status == PaperSettingCompletionRequest.APPROVED
                else None
            )
            history_enriched.append({'r': r, 'credit': cr})
        return render(
            request,
            'core/exam_paper/setting_child_view.html',
            {
                'duties': duties,
                'profile': prof,
                'subunit': (prof.subunit_code or '').strip().upper(),
                'department_name': dept_name,
                'pending_setting': pending,
                'pending_setting_enriched': pending_enriched,
                'history_setting_enriched': history_enriched,
                'phases': phases,
                'selected_phase_id': selected_phase_id,
            },
        )
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phases = _setting_phases_for_user(request)
    institute = _institute_paper_access(request)
    return render(
        request,
        'core/exam_paper/setting_dashboard.html',
        {
            'phases': phases,
            'institute_scope_active': institute,
            'page_kind': 'setting',
        },
    )


@login_required
@require_http_methods(['POST'])
def paper_checking_phase_add(request):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    name = (request.POST.get('name') or '').strip()
    if not name:
        messages.error(request, 'Phase name is required.')
        return redirect('core:paper_checking_dashboard')
    if _dept_parent_only(request):
        prof = _parent_profile(request)
        if not prof:
            return redirect('accounts:role_redirect')
        if _is_hub_coordinator(prof):
            sem = institute_semester_for_exam_portal(request)
            if not sem:
                messages.error(
                    request, 'Set an active academic semester first (Academic semesters).'
                )
                return redirect('core:paper_checking_dashboard')
            obj, created = PaperCheckingPhase.objects.get_or_create(
                hub_coordinator=request.user,
                name=name,
                institute_semester=sem,
                department=None,
                institute_scope=False,
                defaults={'created_by': request.user},
            )
        else:
            if not prof.department_id:
                messages.error(request, 'Link a department first.')
                return redirect('core:dept_exam_dashboard')
            dept = prof.department
            obj, created = PaperCheckingPhase.objects.get_or_create(
                department=dept,
                name=name,
                institute_scope=False,
                defaults={
                    'created_by': request.user,
                    'hub_coordinator': None,
                    'institute_semester': dept.institute_semester,
                },
            )
    elif _institute_paper_access(request):
        obj, created = PaperCheckingPhase.objects.get_or_create(
            institute_scope=True,
            name=name,
            defaults={'created_by': request.user},
        )
    else:
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    messages.success(request, f'Phase "{name}" created.') if created else messages.info(request, f'Phase "{name}" already exists.')
    return redirect('core:paper_checking_dashboard')


@login_required
@require_http_methods(['POST'])
def paper_setting_phase_add(request):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    name = (request.POST.get('name') or '').strip()
    if not name:
        messages.error(request, 'Phase name is required.')
        return redirect('core:paper_setting_dashboard')
    if _dept_parent_only(request):
        prof = _parent_profile(request)
        if not prof:
            return redirect('accounts:role_redirect')
        if _is_hub_coordinator(prof):
            sem = institute_semester_for_exam_portal(request)
            if not sem:
                messages.error(
                    request, 'Set an active academic semester first (Academic semesters).'
                )
                return redirect('core:paper_setting_dashboard')
            obj, created = PaperSettingPhase.objects.get_or_create(
                hub_coordinator=request.user,
                name=name,
                institute_semester=sem,
                department=None,
                institute_scope=False,
                defaults={'created_by': request.user},
            )
        else:
            if not prof.department_id:
                messages.error(request, 'Link a department first.')
                return redirect('core:dept_exam_dashboard')
            dept = prof.department
            obj, created = PaperSettingPhase.objects.get_or_create(
                department=dept,
                name=name,
                institute_scope=False,
                defaults={
                    'created_by': request.user,
                    'hub_coordinator': None,
                    'institute_semester': dept.institute_semester,
                },
            )
    elif _institute_paper_access(request):
        obj, created = PaperSettingPhase.objects.get_or_create(
            institute_scope=True,
            name=name,
            defaults={'created_by': request.user},
        )
    else:
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    messages.success(request, f'Phase "{name}" created.') if created else messages.info(request, f'Phase "{name}" already exists.')
    return redirect('core:paper_setting_dashboard')


@login_required
@require_http_methods(['POST'])
def paper_setting_phase_rename(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_setting_phase(request, phase_id)
    name = (request.POST.get('name') or '').strip()
    if not name:
        messages.error(request, 'Phase name is required.')
        return redirect('core:paper_setting_dashboard')
    phase.name = name
    try:
        phase.save(update_fields=['name'])
    except IntegrityError:
        messages.error(request, 'That phase name already exists in your scope.')
        return redirect('core:paper_setting_dashboard')
    messages.success(request, f'Phase renamed to "{name}".')
    return redirect('core:paper_setting_dashboard')


@login_required
@require_http_methods(['POST'])
def paper_setting_phase_delete(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_setting_phase(request, phase_id)
    if phase.duties.exists():
        messages.error(
            request,
            'This phase still has duty rows. Open the phase and clear duties first.',
        )
        return redirect('core:paper_setting_dashboard')
    n = phase.name
    phase.delete()
    messages.success(request, f'Phase "{n}" deleted.')
    return redirect('core:paper_setting_dashboard')


@login_required
@require_http_methods(['POST'])
def paper_checking_phase_rename(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_checking_phase(request, phase_id)
    name = (request.POST.get('name') or '').strip()
    if not name:
        messages.error(request, 'Phase name is required.')
        return redirect('core:paper_checking_dashboard')
    phase.name = name
    try:
        phase.save(update_fields=['name'])
    except IntegrityError:
        messages.error(request, 'That phase name already exists in your scope.')
        return redirect('core:paper_checking_dashboard')
    messages.success(request, f'Phase renamed to "{name}".')
    return redirect('core:paper_checking_dashboard')


@login_required
@require_http_methods(['POST'])
def paper_checking_phase_delete(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_checking_phase(request, phase_id)
    if phase.duties.exists():
        messages.error(
            request,
            'This phase still has duties. Open the phase and clear the sheet first.',
        )
        return redirect('core:paper_checking_dashboard')
    n = phase.name
    phase.delete()
    messages.success(request, f'Phase "{n}" deleted.')
    return redirect('core:paper_checking_dashboard')


@login_required
def paper_checking_phase_detail(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_checking_phase(request, phase_id)
    prof = _parent_profile(request) if _dept_parent_only(request) else None
    dept = prof.department if prof and prof.department_id else None
    staging_blob = paper_checking_stage_get(request, phase.id)
    staging_preview = None
    if staging_blob:
        rows = paper_checking_stage_deserialize_rows(staging_blob)
        staging_preview = []
        for r in rows:
            fac = _faculty_match_checking(phase, dept, '', r['evaluator_initial'])
            staging_preview.append({**r, 'faculty': fac})
    duties = (
        phase.duties.select_related('faculty', 'faculty__department')
        .prefetch_related('allocations__department')
        .order_by('exam_date', 'subject_name', 'evaluator_short_raw')
    )
    subjects = sorted(
        {(d.subject_name or '').strip() for d in phase.duties.all() if (d.subject_name or '').strip()},
        key=lambda x: x.upper(),
    )
    creds = list(phase.subject_credits.all())
    by_l = {c.subject_name.strip().lower(): c for c in creds}
    credit_rows = []
    for s in subjects:
        ex = by_l.get(s.lower())
        credit_rows.append(
            {
                'name': s,
                'is_practical': bool(ex and ex.is_practical),
                'theory': ex.credit_per_paper_theory if ex else '',
                'online': ex.credit_online_per_paper if ex else '',
                'offline': ex.credit_offline_per_paper if ex else '',
            }
        )
    return render(
        request,
        'core/exam_paper/checking_phase_detail.html',
        {
            'phase': phase,
            'duties': duties,
            'profile': prof,
            'credit_rows': credit_rows,
            'staging_preview': staging_preview,
            'staging_unmatched': staging_blob.get('n_unmatched') if staging_blob else None,
        },
    )


@login_required
def paper_setting_phase_detail(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_setting_phase(request, phase_id)
    prof = _parent_profile(request) if _dept_parent_only(request) else None
    dept = prof.department if prof and prof.department_id else None
    staging_blob = paper_setting_stage_get(request, phase.id)
    staging_preview = None
    if staging_blob:
        rows = paper_setting_stage_deserialize_rows(staging_blob)
        staging_preview = []
        for r in rows:
            fac = _faculty_match_setting(phase, dept, '', r['faculty_initial'])
            dl = r.get('deadline_date') or r.get('duty_date')
            staging_preview.append({**r, 'faculty': fac, 'deadline_date': dl})
    duties = phase.duties.select_related('faculty', 'faculty__department').order_by(
        'duty_date', 'subject_name', 'faculty_short_raw'
    )
    return render(
        request,
        'core/exam_paper/setting_phase_detail.html',
        {
            'phase': phase,
            'duties': duties,
            'profile': prof,
            'staging_preview': staging_preview,
            'staging_unmatched': staging_blob.get('n_unmatched') if staging_blob else None,
        },
    )


@login_required
@require_http_methods(['POST'])
def paper_checking_phase_upload(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_checking_phase(request, phase_id)
    prof = _parent_profile(request)
    dept = prof.department if prof and prof.department_id else None
    f = request.FILES.get('paper_file')
    if not f:
        messages.error(request, 'Choose an Excel file to upload.')
        return redirect('core:paper_checking_phase_detail', phase_id=phase.id)
    try:
        rows = parse_paper_checking_workbook(f)
    except Exception as e:
        messages.error(request, str(e))
        return redirect('core:paper_checking_phase_detail', phase_id=phase.id)

    n_unmatched = sum(
        1 for row in rows if not _faculty_match_checking(phase, dept, '', row['evaluator_initial'])
    )
    paper_checking_stage_put(request, phase.id, rows, n_unmatched)
    messages.info(
        request,
        f'Loaded {len(rows)} row(s) for review (not saved yet). '
        f'{"No" if n_unmatched == 0 else n_unmatched} row(s) could not be linked to a faculty record. '
        'Click Save to database to replace current duties for this phase.',
    )
    return redirect('core:paper_checking_phase_detail', phase_id=phase.id)


@login_required
@require_http_methods(['POST'])
def paper_setting_phase_upload(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_setting_phase(request, phase_id)
    prof = _parent_profile(request)
    dept = prof.department if prof and prof.department_id else None
    f = request.FILES.get('paper_file')
    if not f:
        messages.error(request, 'Choose an Excel file to upload.')
        return redirect('core:paper_setting_phase_detail', phase_id=phase.id)
    try:
        rows = parse_paper_setting_workbook(f)
    except Exception as e:
        messages.error(request, str(e))
        return redirect('core:paper_setting_phase_detail', phase_id=phase.id)

    n_unmatched = sum(
        1 for row in rows if not _faculty_match_setting(phase, dept, '', row['faculty_initial'])
    )
    paper_setting_stage_put(request, phase.id, rows, n_unmatched)
    messages.info(
        request,
        f'Loaded {len(rows)} row(s) for review (not saved yet). '
        f'{"No" if n_unmatched == 0 else n_unmatched} row(s) could not be linked to a faculty. '
        'Click Save to database to replace current duties for this phase.',
    )
    return redirect('core:paper_setting_phase_detail', phase_id=phase.id)


@login_required
@require_http_methods(['POST'])
def paper_checking_phase_commit(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_checking_phase(request, phase_id)
    blob = paper_checking_stage_get(request, phase.id)
    if not blob:
        messages.error(request, 'No imported data in this session. Upload an Excel file first.')
        return redirect('core:paper_checking_phase_detail', phase_id=phase.id)
    rows = paper_checking_stage_deserialize_rows(blob)
    prof = _parent_profile(request)
    dept = prof.department if prof and prof.department_id else None
    n = _commit_paper_checking_rows(phase, dept, rows)
    clear_staging(request, 'paper_checking', phase.id)
    messages.success(
        request, f'Saved {n} paper checking row(s) to the database for phase {phase.name}.'
    )
    return redirect('core:paper_checking_phase_detail', phase_id=phase.id)


@login_required
@require_http_methods(['POST'])
def paper_checking_phase_discard_staging(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_checking_phase(request, phase_id)
    if paper_checking_stage_get(request, phase.id):
        clear_staging(request, 'paper_checking', phase.id)
        messages.info(request, 'Discarded the draft import.')
    else:
        messages.info(request, 'No draft import to discard.')
    return redirect('core:paper_checking_phase_detail', phase_id=phase.id)


@login_required
@require_http_methods(['POST'])
def paper_setting_phase_commit(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_setting_phase(request, phase_id)
    blob = paper_setting_stage_get(request, phase.id)
    if not blob:
        messages.error(request, 'No imported data in this session. Upload an Excel file first.')
        return redirect('core:paper_setting_phase_detail', phase_id=phase.id)
    rows = paper_setting_stage_deserialize_rows(blob)
    prof = _parent_profile(request)
    dept = prof.department if prof and prof.department_id else None
    n = _commit_paper_setting_rows(phase, dept, rows)
    clear_staging(request, 'paper_setting', phase.id)
    messages.success(
        request, f'Saved {n} paper setting row(s) to the database for phase {phase.name}.'
    )
    return redirect('core:paper_setting_phase_detail', phase_id=phase.id)


@login_required
@require_http_methods(['POST'])
def paper_setting_phase_discard_staging(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_setting_phase(request, phase_id)
    if paper_setting_stage_get(request, phase.id):
        clear_staging(request, 'paper_setting', phase.id)
        messages.info(request, 'Discarded the draft import.')
    else:
        messages.info(request, 'No draft import to discard.')
    return redirect('core:paper_setting_phase_detail', phase_id=phase.id)


@login_required
@require_http_methods(['POST'])
def paper_checking_phase_clear(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_checking_phase(request, phase_id)
    clear_staging(request, 'paper_checking', phase.id)
    n, _ = phase.duties.all().delete()
    messages.success(request, f'Removed {n} paper checking row(s) for phase {phase.name}.')
    return redirect('core:paper_checking_phase_detail', phase_id=phase.id)


@login_required
@require_http_methods(['POST'])
def paper_setting_phase_clear(request, phase_id):
    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_setting_phase(request, phase_id)
    clear_staging(request, 'paper_setting', phase.id)
    n, _ = phase.duties.all().delete()
    messages.success(request, f'Removed {n} paper setting row(s) for phase {phase.name}.')
    return redirect('core:paper_setting_phase_detail', phase_id=phase.id)


def _child_can_access_paper_completion_request(prof, req: PaperCheckingCompletionRequest) -> bool:
    return paper_checking_duties_for_child_prof(prof).filter(pk=req.duty_id).exists()


@login_required
@require_http_methods(['POST'])
def paper_checking_child_approve_completion(request, pk):
    if not _dept_child_only(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    prof = _child_profile(request)
    if not prof:
        return redirect('accounts:role_redirect')
    req = get_object_or_404(
        PaperCheckingCompletionRequest,
        pk=pk,
        status=PaperCheckingCompletionRequest.PENDING,
    )
    if not _child_can_access_paper_completion_request(prof, req):
        messages.error(request, 'This request is outside your sub-unit scope.')
        return redirect('core:paper_checking_dashboard')
    req.status = PaperCheckingCompletionRequest.APPROVED
    req.decided_at = timezone.now()
    req.decided_by = request.user
    req.save(update_fields=['status', 'decided_at', 'decided_by'])
    messages.success(
        request,
        f'Approved paper checking completion for {req.faculty.full_name} — {req.duty.subject_name} ({req.duty.exam_date}).',
    )
    return redirect('core:paper_checking_dashboard')


@login_required
@require_http_methods(['POST'])
def paper_checking_child_dismiss_completion(request, pk):
    if not _dept_child_only(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    prof = _child_profile(request)
    if not prof:
        return redirect('accounts:role_redirect')
    req = get_object_or_404(
        PaperCheckingCompletionRequest,
        pk=pk,
        status=PaperCheckingCompletionRequest.PENDING,
    )
    if not _child_can_access_paper_completion_request(prof, req):
        messages.error(request, 'This request is outside your sub-unit scope.')
        return redirect('core:paper_checking_dashboard')
    req.status = PaperCheckingCompletionRequest.REJECTED
    req.decided_at = timezone.now()
    req.decided_by = request.user
    req.save(update_fields=['status', 'decided_at', 'decided_by'])
    messages.info(request, 'Request dismissed. The faculty can submit completion again if needed.')
    return redirect('core:paper_checking_dashboard')


@login_required
@require_http_methods(['POST'])
def paper_checking_phase_subject_credits_save(request, phase_id):
    from decimal import Decimal, InvalidOperation

    if not _paper_coord_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    phase = _get_checking_phase(request, phase_id)
    subjects = request.POST.getlist('credit_subject')
    theories = request.POST.getlist('credit_theory')
    onlines = request.POST.getlist('credit_online')
    offlines = request.POST.getlist('credit_offline')
    practical_on = set(request.POST.getlist('credit_is_practical'))

    def to_dec(x) -> Decimal:
        try:
            return Decimal(((x or '') + '').strip() or '0')
        except InvalidOperation:
            return Decimal('0')

    with transaction.atomic():
        PaperCheckingSubjectCredit.objects.filter(phase=phase).delete()
        n = 0
        for i, subj in enumerate(subjects):
            subj = (subj or '').strip()
            if not subj:
                continue
            th = to_dec(theories[i] if i < len(theories) else '0')
            on = to_dec(onlines[i] if i < len(onlines) else '0')
            off = to_dec(offlines[i] if i < len(offlines) else '0')
            is_p = subj in practical_on
            PaperCheckingSubjectCredit.objects.create(
                phase=phase,
                subject_name=subj,
                is_practical=is_p,
                credit_per_paper_theory=th,
                credit_online_per_paper=on if is_p else None,
                credit_offline_per_paper=off if is_p else None,
            )
            n += 1
    messages.success(request, f'Saved credit rules for {n} subject(s) in phase “{phase.name}”.')
    return redirect('core:paper_checking_phase_detail', phase_id=phase.id)


@login_required
@require_http_methods(['POST'])
def paper_checking_child_save_adjustment(request):
    from django.urls import reverse
    from urllib.parse import urlencode

    if not _dept_child_only(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    prof = _child_profile(request)
    if not prof or not prof.department_id:
        messages.error(request, 'No department profile.')
        return redirect('accounts:role_redirect')
    try:
        duty_id = int(request.POST.get('duty_id', ''))
    except (TypeError, ValueError):
        messages.error(request, 'Invalid duty.')
        return redirect('core:paper_checking_dashboard')
    duty = get_object_or_404(PaperCheckingDuty.objects.select_related('phase'), pk=duty_id)
    if not paper_checking_duties_for_child_prof(prof).filter(pk=duty.pk).exists():
        messages.error(request, 'This duty is outside your sub-unit scope.')
        return redirect('core:paper_checking_dashboard')
    raw_ids = request.POST.getlist('share_faculty')
    raw_counts = request.POST.getlist('share_papers')
    if len(raw_ids) != len(raw_counts):
        messages.error(request, 'Each row needs a faculty and paper count.')
        return redirect('core:paper_checking_dashboard')
    if not raw_ids:
        messages.error(request, 'Add at least one faculty and paper count.')
        return redirect('core:paper_checking_dashboard')
    pairs: list[tuple[int, int]] = []
    seen_fac: set[int] = set()
    for sid, scnt in zip(raw_ids, raw_counts):
        sid = (sid or '').strip()
        scnt = (scnt or '').strip()
        if not sid or not scnt:
            messages.error(request, 'Faculty and paper count are required on every row.')
            return redirect('core:paper_checking_dashboard')
        try:
            fid = int(sid)
            pc = int(scnt)
        except ValueError:
            messages.error(request, 'Invalid faculty or paper count.')
            return redirect('core:paper_checking_dashboard')
        if pc < 1:
            messages.error(request, 'Paper counts must be positive.')
            return redirect('core:paper_checking_dashboard')
        if fid in seen_fac:
            messages.error(request, 'Each faculty can appear only once.')
            return redirect('core:paper_checking_dashboard')
        seen_fac.add(fid)
        pairs.append((fid, pc))
    fac_ids = [p[0] for p in pairs]
    valid = Faculty.objects.filter(pk__in=fac_ids, department_id=prof.department_id).count()
    if valid != len(fac_ids):
        messages.error(request, 'All faculty must belong to your department.')
        return redirect('core:paper_checking_dashboard')
    total = sum(p[1] for p in pairs)
    if total != duty.total_students:
        messages.error(
            request,
            f'Sum of assigned papers ({total}) must equal the duty total ({duty.total_students}).',
        )
        return redirect('core:paper_checking_dashboard')
    with transaction.atomic():
        PaperCheckingAdjustedShare.objects.filter(duty=duty).delete()
        PaperCheckingCompletionRequest.objects.filter(duty=duty).delete()
        PaperCheckingAdjustedShare.objects.bulk_create(
            [
                PaperCheckingAdjustedShare(
                    duty=duty,
                    faculty_id=fid,
                    paper_count=pc,
                    created_by_id=request.user.id,
                )
                for fid, pc in pairs
            ]
        )
    messages.success(
        request,
        f'Adjustment saved for {duty.subject_name} — {len(pairs)} faculty share {duty.total_students} papers. Prior completion requests were cleared.',
    )
    phase_kw = (request.POST.get('filter_phase_id') or '').strip()
    sub_kw = (request.POST.get('filter_subject') or '').strip()
    base = reverse('core:paper_checking_dashboard')
    if phase_kw:
        q = urlencode({'phase_id': phase_kw, 'subject': sub_kw})
        return redirect(f'{base}?{q}')
    return redirect(base)


def _after_paper_setting_decision_redirect(request):
    if _dept_child_only(request):
        return redirect('core:paper_setting_dashboard')
    if _dept_parent_only(request):
        return redirect('core:dept_exam_dashboard')
    if _institute_paper_access(request):
        return redirect('core:exam_section_dashboard')
    return redirect('accounts:role_redirect')


def _can_access_paper_setting_completion_request(request, req: PaperSettingCompletionRequest) -> bool:
    if _dept_child_only(request):
        prof = _child_profile(request)
        return bool(
            prof and paper_setting_duties_for_child_prof(prof).filter(pk=req.duty_id).exists()
        )
    if _dept_parent_only(request):
        prof = _parent_profile(request)
        if not prof:
            return False
        if _is_hub_coordinator(prof):
            dept_ids = hub_managed_department_ids(prof)
            if req.duty.phase_id and req.duty.phase.hub_coordinator_id == prof.user_id:
                return True
            for did in dept_ids:
                d = Department.objects.filter(pk=did).first()
                if d and faculty_has_department_access(req.faculty, d):
                    return True
            return False
        return bool(
            prof.department_id
            and req.faculty_id
            and faculty_has_department_access(req.faculty, prof.department)
        )
    if _institute_paper_access(request):
        return True
    return False


@login_required
@require_http_methods(['POST'])
def paper_setting_completion_approve(request, pk):
    if not (_dept_child_only(request) or _dept_parent_only(request) or _institute_paper_access(request)):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    req = get_object_or_404(
        PaperSettingCompletionRequest,
        pk=pk,
        status=PaperSettingCompletionRequest.PENDING,
    )
    if not _can_access_paper_setting_completion_request(request, req):
        messages.error(request, 'This request is outside your scope.')
        return _after_paper_setting_decision_redirect(request)
    req.status = PaperSettingCompletionRequest.APPROVED
    req.decided_at = timezone.now()
    req.decided_by = request.user
    req.save(update_fields=['status', 'decided_at', 'decided_by'])
    messages.success(
        request,
        f'Approved paper setting for {req.faculty.full_name} — {req.duty.subject_name} '
        f'({req.duty.phase.name if req.duty.phase else ""}). Credit: {credit_for_paper_setting_request(req)}.',
    )
    return _after_paper_setting_decision_redirect(request)


@login_required
@require_http_methods(['POST'])
def paper_setting_completion_dismiss(request, pk):
    if not (_dept_child_only(request) or _dept_parent_only(request) or _institute_paper_access(request)):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    req = get_object_or_404(
        PaperSettingCompletionRequest,
        pk=pk,
        status=PaperSettingCompletionRequest.PENDING,
    )
    if not _can_access_paper_setting_completion_request(request, req):
        messages.error(request, 'This request is outside your scope.')
        return _after_paper_setting_decision_redirect(request)
    req.status = PaperSettingCompletionRequest.REJECTED
    req.decided_at = timezone.now()
    req.decided_by = request.user
    req.save(update_fields=['status', 'decided_at', 'decided_by'])
    messages.info(request, 'Paper setting request dismissed. Faculty can submit again if needed.')
    return _after_paper_setting_decision_redirect(request)


def _credit_settings_access(request) -> bool:
    return _dept_parent_only(request) or _institute_paper_access(request)


def _paper_checking_phases_for_credit_settings(request, prof):
    from core.models import PaperCheckingPhase
    from core.semester_scope import institute_semester_for_exam_portal

    if _institute_paper_access(request):
        if is_exam_section_operator(request):
            return checking_phases_institute_for_request(request)
        return PaperCheckingPhase.objects.order_by('name')
    if not prof:
        return PaperCheckingPhase.objects.none()
    if _is_hub_coordinator(prof):
        sem = institute_semester_for_exam_portal(request)
        if not sem:
            return PaperCheckingPhase.objects.none()
        return PaperCheckingPhase.objects.filter(
            hub_coordinator=request.user, institute_semester=sem
        ).order_by('name')
    if prof.department_id:
        return PaperCheckingPhase.objects.filter(department=prof.department).order_by('name')
    return PaperCheckingPhase.objects.none()


def _credit_settings_scope_depts(request, prof) -> list[Department]:
    if _institute_paper_access(request):
        return list(
            departments_for_exam_coordination_request(request)
            .select_related('institute_semester')
            .order_by('name')
        )
    if prof and _is_hub_coordinator(prof):
        return list(
            Department.objects.filter(pk__in=hub_managed_department_ids(prof))
            .select_related('institute_semester')
            .order_by('name')
        )
    if prof and prof.department_id:
        return [prof.department]
    return []


def _parse_credit_scope(request, all_depts: list[Department]) -> str:
    raw = (request.GET.get('credit_scope') or request.POST.get('credit_scope') or '').strip()
    if raw in ('', 'all'):
        return 'all'
    try:
        sid = int(raw)
    except (TypeError, ValueError):
        return 'all'
    if not any(d.pk == sid for d in all_depts):
        return 'all'
    return str(sid)


def _rules_department_for_scope(scope: str, all_depts: list[Department]) -> Department | None:
    if scope == 'all':
        return None
    try:
        pk = int(scope)
    except (TypeError, ValueError):
        return None
    return next((d for d in all_depts if d.pk == pk), None)


@login_required
def exam_credit_settings(request):
    from decimal import Decimal, InvalidOperation
    from django.urls import reverse

    if not _credit_settings_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    prof = _parent_profile(request) if _dept_parent_only(request) else None
    all_depts = _credit_settings_scope_depts(request, prof)
    if _institute_paper_access(request) and not all_depts:
        messages.error(request, 'No departments found.')
        return redirect('core:exam_section_dashboard')
    if prof and _is_hub_coordinator(prof) and not all_depts:
        messages.error(
            request,
            'No departments linked to your hub yet. Create department logins first.',
        )
        return redirect('core:dept_exam_dashboard')
    if prof and not prof.department_id and not _is_hub_coordinator(prof) and not _institute_paper_access(request):
        messages.error(request, 'Link a department first.')
        return redirect('core:dept_exam_dashboard')

    scope = _parse_credit_scope(request, all_depts)
    rules_dep = _rules_department_for_scope(scope, all_depts)
    scope_choices = [('all', 'All departments (institute default)')] + [
        (str(d.pk), d.name) for d in all_depts
    ]

    def to_dec(x) -> Decimal:
        try:
            return Decimal(((x or '') + '').strip() or '0')
        except InvalidOperation:
            return Decimal('0')

    buckets = [
        DepartmentExamCreditRule.BUCKET_T1_T3,
        DepartmentExamCreditRule.BUCKET_SEE,
        DepartmentExamCreditRule.BUCKET_REMEDIAL,
        DepartmentExamCreditRule.BUCKET_FAST_TRACK,
    ]

    def save_bucket_task(form_task: str, post_prefix: str, with_overrides: bool) -> None:
        if rules_dep is None:
            DepartmentExamCreditRule.objects.filter(
                task=form_task, department__isnull=True
            ).delete()
        else:
            DepartmentExamCreditRule.objects.filter(
                task=form_task, department=rules_dep
            ).delete()
        for b in buckets:
            c = to_dec(request.POST.get(f'{post_prefix}_{b}', '0'))
            rem = to_dec(request.POST.get(f'{post_prefix}_rem_{b}', '0'))
            if c == 0 and rem == 0:
                continue
            DepartmentExamCreditRule.objects.create(
                department=rules_dep,
                task=form_task,
                phase_bucket=b,
                subject_name='',
                credit=c,
                remuneration=rem,
            )
        if not with_overrides:
            return
        subj_list = request.POST.getlist(f'{post_prefix}_ov_subject')
        cred_list = request.POST.getlist(f'{post_prefix}_ov_credit')
        rem_list = request.POST.getlist(f'{post_prefix}_ov_remuneration')
        bucket_list = request.POST.getlist(f'{post_prefix}_ov_bucket')
        for i, subj in enumerate(subj_list):
            subj = (subj or '').strip()
            if not subj:
                continue
            b = (bucket_list[i] if i < len(bucket_list) else '') or buckets[0]
            if b not in buckets:
                b = buckets[0]
            c = to_dec(cred_list[i] if i < len(cred_list) else '0')
            rem = to_dec(rem_list[i] if i < len(rem_list) else '0')
            if c == 0 and rem == 0:
                continue
            DepartmentExamCreditRule.objects.create(
                department=rules_dep,
                task=form_task,
                phase_bucket=b,
                subject_name=subj,
                credit=c,
                remuneration=rem,
            )

    if request.method == 'POST':
        task_ps = (request.POST.get('save_task') or '').strip()
        redir = f"{reverse('core:exam_credit_settings')}?credit_scope={scope}"
        if task_ps == DepartmentExamCreditRule.TASK_PAPER_SETTING:
            with transaction.atomic():
                save_bucket_task(
                    DepartmentExamCreditRule.TASK_PAPER_SETTING, 'ps', with_overrides=True
                )
            messages.success(request, 'Paper setting credit rules saved.')
            return redirect(redir)
        if task_ps == DepartmentExamCreditRule.TASK_SUPERVISION:
            with transaction.atomic():
                save_bucket_task(
                    DepartmentExamCreditRule.TASK_SUPERVISION, 'sv', with_overrides=False
                )
            messages.success(request, 'Supervision credit rules saved.')
            return redirect(redir)
        if task_ps == DepartmentExamCreditRule.TASK_PAPER_CHECKING:
            with transaction.atomic():
                save_bucket_task(
                    DepartmentExamCreditRule.TASK_PAPER_CHECKING, 'pc', with_overrides=True
                )
            messages.success(
                request,
                'Paper checking fallback credits saved (theory: per paper when no phase subject row).',
            )
            return redirect(redir)

    rule_filter = {'department': rules_dep} if rules_dep is not None else {'department__isnull': True}

    rules_ps = list(
        DepartmentExamCreditRule.objects.filter(
            task=DepartmentExamCreditRule.TASK_PAPER_SETTING, **rule_filter
        ).order_by('phase_bucket', 'subject_name')
    )
    rules_sv = list(
        DepartmentExamCreditRule.objects.filter(
            task=DepartmentExamCreditRule.TASK_SUPERVISION, **rule_filter
        ).order_by('phase_bucket', 'subject_name')
    )
    rules_pc = list(
        DepartmentExamCreditRule.objects.filter(
            task=DepartmentExamCreditRule.TASK_PAPER_CHECKING, **rule_filter
        ).order_by('phase_bucket', 'subject_name')
    )

    def split_defaults(rules):
        defaults = {b: '' for b in buckets}
        rem_defaults = {b: '' for b in buckets}
        overrides = []
        for r in rules:
            if (r.subject_name or '').strip():
                overrides.append(r)
            else:
                defaults[r.phase_bucket] = str(r.credit).rstrip('0').rstrip('.') if r.credit else ''
                rem_defaults[r.phase_bucket] = str(r.remuneration).rstrip('0').rstrip('.')
        return defaults, rem_defaults, overrides

    ps_defaults, ps_rem_defaults, ps_overrides = split_defaults(rules_ps)
    sv_defaults = {b: '' for b in buckets}
    sv_rem_defaults = {b: '' for b in buckets}
    for r in rules_sv:
        if not (r.subject_name or '').strip():
            sv_defaults[r.phase_bucket] = str(r.credit).rstrip('0').rstrip('.') if r.credit else ''
            sv_rem_defaults[r.phase_bucket] = str(r.remuneration).rstrip('0').rstrip('.')

    pc_defaults, pc_rem_defaults, pc_overrides = split_defaults(rules_pc)

    bl = dict(DepartmentExamCreditRule.BUCKET_CHOICES)
    bucket_key_labels = [(b, bl[b]) for b in buckets]
    ps_default_rows = [(b, bl[b], ps_defaults.get(b, ''), ps_rem_defaults.get(b, '')) for b in buckets]
    sv_default_rows = [(b, bl[b], sv_defaults.get(b, ''), sv_rem_defaults.get(b, '')) for b in buckets]
    pc_default_rows = [(b, bl[b], pc_defaults.get(b, ''), pc_rem_defaults.get(b, '')) for b in buckets]

    checking_phases = _paper_checking_phases_for_credit_settings(request, prof)
    scope_label = (
        'All departments (institute default)' if scope == 'all' else (rules_dep.name if rules_dep else '')
    )
    return render(
        request,
        'core/exam_paper/credit_settings.html',
        {
            'credit_scope': scope,
            'scope_choices': scope_choices,
            'scope_label': scope_label,
            'buckets': buckets,
            'bucket_key_labels': bucket_key_labels,
            'ps_default_rows': ps_default_rows,
            'ps_overrides': ps_overrides,
            'sv_default_rows': sv_default_rows,
            'pc_default_rows': pc_default_rows,
            'pc_overrides': pc_overrides,
            'checking_phases': checking_phases,
            'institute_scope': _institute_paper_access(request),
        },
    )
