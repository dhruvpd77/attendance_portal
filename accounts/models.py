"""
User roles: Admin, Faculty, Student.
"""
from django.db import models
from django.contrib.auth.models import User


class UserRole(models.Model):
    ROLE_CHOICES = [
        ('admin', 'Admin'),
        ('hod', 'HOD'),
        ('exam_admin', 'Exam Admin'),
        ('faculty', 'Faculty'),
        ('student', 'Student'),
    ]
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='role_profile')
    role = models.CharField(max_length=20, choices=ROLE_CHOICES)
    # For role='admin': if set, this user is a departmental admin (only this dept); if null, super admin (all depts).
    department = models.ForeignKey(
        'core.Department', on_delete=models.SET_NULL, null=True, blank=True, related_name='admin_users'
    )

    def __str__(self):
        return f"{self.user.username} ({self.get_role_display()})"
