"""Merge one Subject into another within a department (shared by management command and admin UI)."""
from django.db import transaction

from core.models import (
    Department,
    Subject,
    ScheduleSlot,
    FacultyCombineDrCache,
    LectureAdjustment,
    ExtraLecture,
    ExamPhaseSubject,
    StudentMark,
    PaperCheckingSubjectCredit,
    PaperCheckingDuty,
    PaperCheckingAllocation,
    SupervisionDuty,
    PaperSettingDuty,
    DepartmentExamCreditRule,
)


class SubjectMergeError(Exception):
    """Invalid merge request or missing subjects."""


def resolve_merge_subjects(
    dept: Department,
    from_name: str,
    to_name: str,
) -> tuple[Subject, Subject, str]:
    """Return (source, target, target_label). Raises SubjectMergeError."""
    from_name = (from_name or '').strip()
    to_name = (to_name or '').strip()
    if not from_name or not to_name:
        raise SubjectMergeError('Enter both source and target subject names.')
    if from_name.lower() == to_name.lower():
        raise SubjectMergeError('Source and target must be different subjects.')

    src = Subject.objects.filter(department=dept, name__iexact=from_name).first()
    dst = Subject.objects.filter(department=dept, name__iexact=to_name).first()
    if not src:
        raise SubjectMergeError(f'Source subject "{from_name}" was not found in {dept.name}.')
    if not dst:
        raise SubjectMergeError(f'Target subject "{to_name}" was not found in {dept.name}.')
    if src.pk == dst.pk:
        raise SubjectMergeError('Source and target are the same record.')

    return src, dst, dst.name


def subject_merge_counts(dept: Department, src: Subject, from_name: str) -> dict[str, int]:
    """Row counts for preview (from_name used for string-backed rows)."""
    d_ids = PaperCheckingAllocation.objects.filter(department=dept).values_list('duty_id', flat=True).distinct()
    return {
        'schedule_slot': ScheduleSlot.objects.filter(subject=src).count(),
        'faculty_combine_dr_cache': FacultyCombineDrCache.objects.filter(subject=src).count(),
        'lecture_adjustment_original': LectureAdjustment.objects.filter(original_subject=src).count(),
        'lecture_adjustment_new': LectureAdjustment.objects.filter(new_subject=src).count(),
        'extra_lecture': ExtraLecture.objects.filter(subject=src).count(),
        'exam_phase_subject': ExamPhaseSubject.objects.filter(subject=src).count(),
        'student_mark': StudentMark.objects.filter(subject=src).count(),
        'paper_checking_subject_credit': PaperCheckingSubjectCredit.objects.filter(
            phase__department=dept,
            subject_name__iexact=from_name,
        ).count(),
        'paper_checking_duty': PaperCheckingDuty.objects.filter(
            pk__in=d_ids,
            subject_name__iexact=from_name,
        ).count(),
        'supervision_duty': SupervisionDuty.objects.filter(
            phase__department=dept,
            subject_name__iexact=from_name,
        ).count(),
        'paper_setting_duty': PaperSettingDuty.objects.filter(
            phase__department=dept,
            subject_name__iexact=from_name,
        ).count(),
        'department_exam_credit_rule': DepartmentExamCreditRule.objects.filter(
            department=dept,
        )
        .exclude(subject_name='')
        .filter(subject_name__iexact=from_name)
        .count(),
    }


def execute_subject_merge(dept: Department, from_name: str, to_name: str) -> dict[str, int]:
    """
    Re-point all references from source subject to target and delete source.
    Returns counts of updated rows (approximate aggregates).
    """
    src, dst, target_label = resolve_merge_subjects(dept, from_name, to_name)
    d_ids = list(
        PaperCheckingAllocation.objects.filter(department=dept).values_list('duty_id', flat=True).distinct()
    )
    stats: dict[str, int] = {}

    with transaction.atomic():
        stats['schedule_slot'] = ScheduleSlot.objects.filter(subject=src).update(subject=dst)
        stats['faculty_combine_dr_cache'] = FacultyCombineDrCache.objects.filter(subject=src).update(subject=dst)
        stats['lecture_adjustment_original'] = LectureAdjustment.objects.filter(original_subject=src).update(
            original_subject=dst
        )
        stats['lecture_adjustment_new'] = LectureAdjustment.objects.filter(new_subject=src).update(new_subject=dst)
        stats['extra_lecture'] = ExtraLecture.objects.filter(subject=src).update(subject=dst)

        eps_list = list(ExamPhaseSubject.objects.filter(subject=src).select_related('exam_phase'))
        stats['exam_phase_subject'] = len(eps_list)
        for eps in eps_list:
            ExamPhaseSubject.objects.get_or_create(exam_phase=eps.exam_phase, subject=dst)
            eps.delete()

        sm_count = 0
        for sm in StudentMark.objects.filter(subject=src).select_related('student', 'exam_phase'):
            twin = StudentMark.objects.filter(
                student=sm.student,
                exam_phase=sm.exam_phase,
                subject=dst,
            ).first()
            if twin:
                if twin.marks_obtained is None and sm.marks_obtained is not None:
                    twin.marks_obtained = sm.marks_obtained
                    twin.save(update_fields=['marks_obtained', 'updated_at'])
                sm.delete()
            else:
                sm.subject = dst
                sm.save(update_fields=['subject', 'updated_at'])
            sm_count += 1
        stats['student_mark'] = sm_count

        pcc = 0
        for pc in PaperCheckingSubjectCredit.objects.filter(
            phase__department=dept,
            subject_name__iexact=from_name,
        ):
            exists = PaperCheckingSubjectCredit.objects.filter(
                phase=pc.phase,
                subject_name__iexact=target_label,
            ).exclude(pk=pc.pk).exists()
            if exists:
                pc.delete()
            else:
                pc.subject_name = target_label
                pc.save(update_fields=['subject_name'])
            pcc += 1
        stats['paper_checking_subject_credit'] = pcc

        stats['paper_checking_duty'] = PaperCheckingDuty.objects.filter(
            pk__in=d_ids,
            subject_name__iexact=from_name,
        ).update(subject_name=target_label)

        stats['supervision_duty'] = SupervisionDuty.objects.filter(
            phase__department=dept,
            subject_name__iexact=from_name,
        ).update(subject_name=target_label)

        stats['paper_setting_duty'] = PaperSettingDuty.objects.filter(
            phase__department=dept,
            subject_name__iexact=from_name,
        ).update(subject_name=target_label)

        dcr = 0
        for rule in DepartmentExamCreditRule.objects.filter(
            department=dept,
        ).exclude(subject_name='').filter(subject_name__iexact=from_name):
            conflict = DepartmentExamCreditRule.objects.filter(
                department=dept,
                task=rule.task,
                phase_bucket=rule.phase_bucket,
                subject_name__iexact=target_label,
            ).exclude(pk=rule.pk).exists()
            if conflict:
                rule.delete()
            else:
                rule.subject_name = target_label
                rule.save(update_fields=['subject_name', 'updated_at'])
            dcr += 1
        stats['department_exam_credit_rule'] = dcr

        src.delete()
        stats['source_deleted'] = 1

    return stats
