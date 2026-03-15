"""
Schedule versioning helpers. Used by views and management commands.
"""
from datetime import date

from django.db.models import Max

from .models import ScheduleSlot


def get_effective_slots_for_date(dept, target_date, extra_filters=None):
    """Return list of ScheduleSlot objects effective on this date.
    Uses version-based logic: for date D, find the active version (max effective_from <= D)
    and return ONLY slots from that version. This ensures old slots (e.g. Tuesday) don't
    appear after a timetable change that moved them (e.g. to Wednesday)."""
    if not isinstance(target_date, date):
        target_date = target_date
    qs = ScheduleSlot.objects.filter(
        department=dept,
        effective_from__lte=target_date
    )
    if extra_filters:
        qs = qs.filter(**extra_filters)
    # Find active version: max effective_from for this date
    active_version = qs.aggregate(m=Max('effective_from'))['m']
    if active_version is None:
        return []
    # Return only slots from the active version (complete schedule, no mixing)
    slots_qs = ScheduleSlot.objects.filter(
        department=dept,
        effective_from=active_version
    ).select_related('batch', 'subject', 'faculty')
    if extra_filters:
        slots_qs = slots_qs.filter(**extra_filters)
    return list(slots_qs)


def get_effective_day_set(dept, target_date):
    """Return set of weekday names that have schedule effective on this date."""
    slots = get_effective_slots_for_date(dept, target_date)
    return {s.day for s in slots if s.day}


def get_all_schedule_days(dept):
    """Return set of weekday names that have schedule in ANY version.
    Use when building phase dates so past dates (old timetable) aren't wrongly excluded
    when a new timetable removes certain weekdays."""
    return set(
        ScheduleSlot.objects.filter(department=dept)
        .values_list('day', flat=True)
        .distinct()
    ) - {None, ''}
