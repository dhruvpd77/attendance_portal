from django.db.models.signals import post_delete, post_save
from django.dispatch import receiver

from .models import Faculty, FacultyAttendance, FacultyCombineDrCache, FacultyDepartmentMembership


@receiver(post_save, sender=Faculty)
def faculty_sync_primary_department_membership(sender, instance, **kwargs):
    """Ensure Faculty.department always has a membership row."""
    if not instance.department_id:
        return
    FacultyDepartmentMembership.objects.get_or_create(
        faculty_id=instance.pk,
        department_id=instance.department_id,
    )
    if instance.portal_canonical_id:
        FacultyDepartmentMembership.objects.get_or_create(
            faculty_id=instance.portal_canonical_id,
            department_id=instance.department_id,
        )


@receiver(post_delete, sender=FacultyAttendance)
def faculty_attendance_delete_combine_cache(sender, instance, **kwargs):
    FacultyCombineDrCache.objects.filter(
        faculty_id=instance.faculty_id,
        date=instance.date,
        batch_id=instance.batch_id,
        lecture_slot=instance.lecture_slot,
    ).delete()
