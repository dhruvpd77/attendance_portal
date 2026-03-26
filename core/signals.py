from django.db.models.signals import post_delete
from django.dispatch import receiver

from .models import FacultyAttendance, FacultyCombineDrCache


@receiver(post_delete, sender=FacultyAttendance)
def faculty_attendance_delete_combine_cache(sender, instance, **kwargs):
    FacultyCombineDrCache.objects.filter(
        faculty_id=instance.faculty_id,
        date=instance.date,
        batch_id=instance.batch_id,
        lecture_slot=instance.lecture_slot,
    ).delete()
