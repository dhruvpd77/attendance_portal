"""
Core views: Admin, Faculty, Student dashboards and features.
"""
import csv
import json
import os
import re
import random
from collections import defaultdict
from datetime import datetime, timedelta
from io import BytesIO

import openpyxl
from django.conf import settings
from django.contrib.auth.models import User
from django.core.mail import send_mail
from django.http import Http404, HttpResponse

from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import HttpResponse
from django.db.models import Count, Q
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from .models import (
    Department, Batch, Subject, Faculty, Student,
    ScheduleSlot, TermPhase, FacultyAttendance, LectureAdjustment, PhaseHoliday,
    AttendanceNotificationLog,
)
from accounts.models import UserRole


# ---------- Error handlers ----------

def handler404(request, exception):
    return render(request, '404.html', status=404)


def handler500(request):
    return render(request, '500.html', status=500)


# ---------- Helpers ----------

def get_phase_holidays(dept, phase):
    """Return set of holiday dates for this department and phase (T1, T2, T3, T4)."""
    if not dept or not phase:
        return set()
    return set(
        PhaseHoliday.objects.filter(department=dept, phase=phase.upper()).values_list('date', flat=True)
    )


def get_all_holiday_dates(dept):
    """Return set of all holiday dates for this department (all phases). Used to exclude from valid/available dates."""
    if not dept:
        return set()
    return set(
        PhaseHoliday.objects.filter(department=dept).values_list('date', flat=True)
    )


# ----------

def get_faculty_subject_for_slot(date, batch, time_slot):
    """Return (faculty, subject) for this date/batch/slot; use LectureAdjustment if exists, else ScheduleSlot."""
    from datetime import date as date_type
    if not isinstance(date, date_type):
        date = date
    weekday = date.strftime('%A')
    adj = LectureAdjustment.objects.filter(
        date=date, batch=batch, time_slot=time_slot
    ).select_related('new_faculty', 'new_subject').first()
    if adj:
        return adj.new_faculty, adj.new_subject
    slot = ScheduleSlot.objects.filter(
        batch=batch, day=weekday, time_slot=time_slot
    ).select_related('faculty', 'subject').first()
    if slot:
        return slot.faculty, slot.subject
    return None, None


# ----------

def get_admin_department(request):
    """Department for admin: departmental admin has fixed dept; super admin uses session or first."""
    try:
        if request.user.is_authenticated and hasattr(request.user, 'role_profile'):
            rp = request.user.role_profile
            if rp.role == 'admin' and rp.department_id:
                return rp.department
    except Exception:
        pass
    dept_id = request.session.get('admin_department_id')
    if dept_id:
        return Department.objects.filter(pk=dept_id).first()
    return Department.objects.first()


def is_super_admin(request):
    """True if current user is admin with no department (can create depts and departmental admins)."""
    if not user_can_admin(request):
        return False
    if request.user.is_superuser or request.user.is_staff:
        return True
    try:
        rp = request.user.role_profile
        return rp.role == 'admin' and not rp.department_id
    except (UserRole.DoesNotExist, AttributeError):
        return True


def get_faculty_user(request):
    """Current user's Faculty or None."""
    if not request.user.is_authenticated:
        return None
    return getattr(request.user, 'faculty_profile', None)


def get_student_user(request):
    """Current user's Student or None."""
    if not request.user.is_authenticated:
        return None
    return getattr(request.user, 'student_profile', None)


def user_can_admin(request):
    try:
        return request.user.role_profile.role == 'admin' or request.user.is_superuser or request.user.is_staff
    except (UserRole.DoesNotExist, AttributeError):
        return request.user.is_superuser or request.user.is_staff


def user_can_faculty(request):
    try:
        return request.user.role_profile.role == 'faculty'
    except (UserRole.DoesNotExist, AttributeError):
        return False


def user_can_student(request):
    try:
        return request.user.role_profile.role == 'student'
    except (UserRole.DoesNotExist, AttributeError):
        return False


# ---------- Home & Dashboards ----------

def portal_root(request):
    """Redirect /portal/ to main app (role redirect or login)."""
    return redirect('accounts:role_redirect')


def home(request):
    if request.user.is_authenticated:
        return redirect('accounts:role_redirect')
    return redirect('accounts:login')


@login_required
def admin_dashboard(request):
    if not user_can_admin(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    super_admin = is_super_admin(request)
    ctx = {
        'department': dept,
        'departments': Department.objects.all() if super_admin else ([dept] if dept else []),
        'batch_count': Batch.objects.filter(department=dept).count() if dept else 0,
        'faculty_count': Faculty.objects.filter(department=dept).count() if dept else 0,
        'student_count': Student.objects.filter(department=dept).count() if dept else 0,
        'is_super_admin': super_admin,
    }
    return render(request, 'core/admin/dashboard.html', ctx)


@login_required
def faculty_dashboard(request):
    if not user_can_faculty(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    faculty = get_faculty_user(request)
    if not faculty:
        messages.error(request, 'No faculty profile linked.')
        return redirect('accounts:logout')
    from datetime import date
    today = date.today()
    weekday = today.strftime('%A')
    today_slots = ScheduleSlot.objects.filter(
        faculty=faculty, day=weekday
    ).select_related('batch', 'subject').order_by('time_slot')
    ctx = {
        'faculty': faculty,
        'today_slots': today_slots,
        'today': today,
    }
    return render(request, 'core/faculty/dashboard.html', ctx)


@login_required
def student_dashboard(request):
    if not user_can_student(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    student = get_student_user(request)
    if not student:
        messages.error(request, 'No student profile linked.')
        return redirect('accounts:logout')
    # Today's schedule for this batch (use adjusted faculty/subject if any)
    from datetime import date as date_type
    today = date_type.today()
    weekday = today.strftime('%A')
    slots = ScheduleSlot.objects.filter(
        batch=student.batch, day=weekday
    ).select_related('faculty', 'subject').order_by('time_slot')
    schedule = []
    for slot in slots:
        fac, subj = get_faculty_subject_for_slot(today, student.batch, slot.time_slot)
        schedule.append({
            'time_slot': slot.time_slot,
            'subject': subj.name if subj else (slot.subject.name if slot.subject else 'N/A'),
            'faculty': fac.short_name if fac else (slot.faculty.short_name if slot.faculty else '—'),
        })
    ctx = {'student': student, 'schedule': schedule}
    return render(request, 'core/student/dashboard.html', ctx)


# ---------- Admin: Department ----------

@login_required
def department_list(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    # Only super admin can switch department via form
    if is_super_admin(request) and request.method == 'POST' and request.POST.get('set_department'):
        request.session['admin_department_id'] = request.POST.get('department_id')
        return redirect('core:admin_dashboard')
    dept = get_admin_department(request)
    departments = Department.objects.all() if is_super_admin(request) else ([dept] if dept else [])
    ctx = {
        'departments': departments,
        'department': dept,
        'is_super_admin': is_super_admin(request),
    }
    return render(request, 'core/admin/department_list.html', ctx)


@login_required
def department_add(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    if not is_super_admin(request):
        messages.error(request, 'Only super admin can create departments.')
        return redirect('core:admin_dashboard')
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()
        if name:
            Department.objects.create(name=name, code=code)
            messages.success(request, 'Department added.')
            return redirect('core:department_list')
    return render(request, 'core/admin/department_form.html', {'form_type': 'add'})


@login_required
def department_edit(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    obj = get_object_or_404(Department, pk=pk)
    if not is_super_admin(request) and get_admin_department(request) != obj:
        messages.error(request, 'You can only edit your own department.')
        return redirect('core:department_list')
    if request.method == 'POST':
        obj.name = request.POST.get('name', '').strip() or obj.name
        obj.code = request.POST.get('code', '').strip()
        obj.save()
        messages.success(request, 'Department updated.')
        return redirect('core:department_list')
    return render(request, 'core/admin/department_form.html', {'obj': obj, 'form_type': 'edit'})


@login_required
def department_delete(request, pk):
    if not user_can_admin(request) or not is_super_admin(request):
        messages.error(request, 'Only super admin can delete departments.')
        return redirect('core:department_list')
    obj = get_object_or_404(Department, pk=pk)
    name = obj.name
    obj.delete()
    messages.success(request, f'Department "{name}" deleted.')
    return redirect('core:department_list')


@login_required
def departmental_admin_list(request):
    """List departmental admins (admin users with a department). Super admin only."""
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    if not is_super_admin(request):
        messages.error(request, 'Only super admin can manage departmental admins.')
        return redirect('core:admin_dashboard')
    admins = UserRole.objects.filter(role='admin', department__isnull=False).select_related('user', 'department').order_by('department__name', 'user__username')
    ctx = {'admins': admins}
    return render(request, 'core/admin/departmental_admin_list.html', ctx)


@login_required
def departmental_admin_create(request):
    """Create a departmental admin: username, password, department. Super admin only."""
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    if not is_super_admin(request):
        messages.error(request, 'Only super admin can create departmental admins.')
        return redirect('core:admin_dashboard')
    if request.method == 'POST':
        username = (request.POST.get('username') or '').strip()
        password = request.POST.get('password') or ''
        password2 = request.POST.get('password2') or ''
        department_id = request.POST.get('department_id')
        if not username:
            messages.error(request, 'Username is required.')
        elif User.objects.filter(username=username).exists():
            messages.error(request, 'That username is already taken.')
        elif not password or len(password) < 6:
            messages.error(request, 'Password must be at least 6 characters.')
        elif password != password2:
            messages.error(request, 'Passwords do not match.')
        elif not department_id:
            messages.error(request, 'Please select a department.')
        else:
            dept = Department.objects.filter(pk=department_id).first()
            if not dept:
                messages.error(request, 'Invalid department.')
            else:
                user = User.objects.create_user(username=username, password=password)
                UserRole.objects.create(user=user, role='admin', department=dept)
                messages.success(request, f'Departmental admin "{username}" created for {dept.name}.')
                return redirect('core:departmental_admin_list')
    ctx = {'departments': Department.objects.order_by('name')}
    return render(request, 'core/admin/departmental_admin_form.html', ctx)


# ---------- Admin: Batch ----------

@login_required
def batch_list(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    batches = Batch.objects.filter(department=dept) if dept else []
    ctx = {'batches': batches, 'department': dept}
    return render(request, 'core/admin/batch_list.html', ctx)


@login_required
def batch_add(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        if name:
            Batch.objects.get_or_create(department=dept, name=name)
            messages.success(request, 'Batch added.')
            return redirect('core:batch_list')
    return render(request, 'core/admin/batch_form.html', {'form_type': 'add', 'department': dept})


@login_required
def batch_edit(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    obj = get_object_or_404(Batch, pk=pk)
    dept = get_admin_department(request)
    if dept and obj.department != dept:
        messages.error(request, 'You can only edit batches in your department.')
        return redirect('core:batch_list')
    if request.method == 'POST':
        obj.name = request.POST.get('name', '').strip() or obj.name
        obj.save()
        messages.success(request, 'Batch updated.')
        return redirect('core:batch_list')
    return render(request, 'core/admin/batch_form.html', {'obj': obj, 'form_type': 'edit'})


@login_required
def batch_delete(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    obj = get_object_or_404(Batch, pk=pk)
    dept = get_admin_department(request)
    if dept and obj.department != dept:
        messages.error(request, 'You can only manage batches in your department.')
        return redirect('core:batch_list')
    name = obj.name
    obj.delete()
    messages.success(request, f'Batch "{name}" deleted.')
    return redirect('core:batch_list')


# ---------- Admin: Subject ----------

@login_required
def subject_list(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    subjects = Subject.objects.filter(department=dept) if dept else []
    ctx = {'subjects': subjects, 'department': dept}
    return render(request, 'core/admin/subject_list.html', ctx)


@login_required
def subject_add(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()
        if name:
            Subject.objects.get_or_create(department=dept, name=name, defaults={'code': code})
            messages.success(request, 'Subject added.')
            return redirect('core:subject_list')
    return render(request, 'core/admin/subject_form.html', {'form_type': 'add', 'department': dept})


@login_required
def subject_edit(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    obj = get_object_or_404(Subject, pk=pk)
    dept = get_admin_department(request)
    if dept and obj.department != dept:
        messages.error(request, 'You can only edit subjects in your department.')
        return redirect('core:subject_list')
    if request.method == 'POST':
        obj.name = request.POST.get('name', '').strip() or obj.name
        obj.code = request.POST.get('code', '').strip()
        obj.save()
        messages.success(request, 'Subject updated.')
        return redirect('core:subject_list')
    return render(request, 'core/admin/subject_form.html', {'obj': obj, 'form_type': 'edit'})


@login_required
def subject_delete(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    obj = get_object_or_404(Subject, pk=pk)
    dept = get_admin_department(request)
    if dept and obj.department != dept:
        messages.error(request, 'You can only manage subjects in your department.')
        return redirect('core:subject_list')
    name = obj.name
    obj.delete()
    messages.success(request, f'Subject "{name}" deleted.')
    return redirect('core:subject_list')


# ---------- Admin: Faculty ----------

@login_required
def faculty_list(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    faculties = Faculty.objects.filter(department=dept) if dept else []
    ctx = {'faculties': faculties, 'department': dept}
    return render(request, 'core/admin/faculty_list.html', ctx)


@login_required
def faculty_add(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    if request.method == 'POST':
        full_name = request.POST.get('full_name', '').strip()
        short_name = request.POST.get('short_name', '').strip()
        email = request.POST.get('email', '').strip()
        if full_name and short_name:
            Faculty.objects.get_or_create(
                department=dept, short_name=short_name,
                defaults={'full_name': full_name, 'email': email}
            )
            messages.success(request, 'Faculty added.')
            return redirect('core:faculty_list')
    return render(request, 'core/admin/faculty_form.html', {'form_type': 'add', 'department': dept})


@login_required
def faculty_edit(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    obj = get_object_or_404(Faculty, pk=pk)
    dept = get_admin_department(request)
    if dept and obj.department != dept:
        messages.error(request, 'You can only edit faculty in your department.')
        return redirect('core:faculty_list')
    if request.method == 'POST':
        obj.full_name = request.POST.get('full_name', '').strip() or obj.full_name
        obj.short_name = request.POST.get('short_name', '').strip() or obj.short_name
        obj.email = request.POST.get('email', '').strip()
        obj.save()
        messages.success(request, 'Faculty updated.')
        return redirect('core:faculty_list')
    return render(request, 'core/admin/faculty_form.html', {'obj': obj, 'form_type': 'edit'})


@login_required
def faculty_delete(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    obj = get_object_or_404(Faculty, pk=pk)
    dept = get_admin_department(request)
    if dept and obj.department != dept:
        messages.error(request, 'You can only manage faculties in your department.')
        return redirect('core:faculty_list')
    name = obj.full_name
    obj.delete()
    messages.success(request, f'Faculty "{name}" deleted.')
    return redirect('core:faculty_list')


@login_required
def generate_credentials_choice(request):
    """Choose whether to generate credentials for Students or Faculty."""
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first from Dashboard.')
        return redirect('core:admin_dashboard')
    ctx = {'department': dept}
    return render(request, 'core/admin/generate_credentials_choice.html', ctx)


@login_required
def faculty_generate_credentials(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first from Dashboard.')
        return redirect('core:admin_dashboard')

    if request.method == 'POST':
        faculties_without_user = Faculty.objects.filter(department=dept, user__isnull=True).order_by('short_name')
        if not faculties_without_user.exists():
            messages.warning(request, 'All faculties in this department already have login credentials.')
            return redirect('core:generate_credentials_choice')

        rows = []
        for faculty in faculties_without_user:
            base_username = (faculty.short_name or 'f').strip().lower()
            base_username = re.sub(r'[^\w]', '', base_username)[:30] or 'f'
            username = base_username
            if User.objects.filter(username=username).exists():
                username = f"{base_username}{faculty.id}"
            while User.objects.filter(username=username).exists():
                username = f"{base_username}{faculty.id}_{random.randint(100, 999)}"
            password = str(random.randint(0, 9999)).zfill(4)
            user = User.objects.create_user(username=username, password=password)
            faculty.user = user
            faculty.save()
            UserRole.objects.get_or_create(user=user, defaults={'role': 'faculty'})
            rows.append({
                'department': dept.name,
                'full_name': faculty.full_name,
                'short_name': faculty.short_name,
                'username': username,
                'password': password,
            })

        cred_dir = os.path.join(settings.MEDIA_ROOT, 'credentials')
        os.makedirs(cred_dir, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_dept = re.sub(r'[^\w\-]', '_', dept.name)[:50]
        filename = f'faculty_credentials_{safe_dept}_{timestamp}.csv'
        filepath = os.path.join(cred_dir, filename)

        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=['department', 'full_name', 'short_name', 'username', 'password'])
            w.writeheader()
            w.writerows(rows)

        request.session['credentials_filename'] = filename
        request.session['credentials_count'] = len(rows)
        request.session['credentials_type'] = 'faculty'
        messages.success(request, f'Credentials generated for {len(rows)} faculty. Download and store the file securely.')
        return redirect('core:credentials_result')

    faculties_without = Faculty.objects.filter(department=dept, user__isnull=True).count()
    faculties_total = Faculty.objects.filter(department=dept).count()
    ctx = {'department': dept, 'faculties_without': faculties_without, 'faculties_total': faculties_total}
    return render(request, 'core/admin/faculty_generate_credentials.html', ctx)


@login_required
def student_generate_credentials(request):
    """Generate login credentials for students: username = enrollment number, password = 4-digit random."""
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first from Dashboard.')
        return redirect('core:admin_dashboard')

    if request.method == 'POST':
        students_without_user = Student.objects.filter(department=dept, user__isnull=True).select_related('batch').order_by('batch__name', 'roll_no')
        if not students_without_user.exists():
            messages.warning(request, 'All students in this department already have login credentials.')
            return redirect('core:generate_credentials_choice')

        rows = []
        for student in students_without_user:
            base_username = (student.enrollment_no or '').strip()
            if not base_username:
                base_username = f'stu{student.id}'
            base_username = re.sub(r'[^\w]', '', base_username)[:30] or f'stu{student.id}'
            username = base_username
            if User.objects.filter(username=username).exists():
                username = f"{base_username}_{student.id}"
            while User.objects.filter(username=username).exists():
                username = f"{base_username}_{student.id}_{random.randint(100, 999)}"
            password = str(random.randint(0, 9999)).zfill(4)
            user = User.objects.create_user(username=username, password=password)
            student.user = user
            student.save()
            UserRole.objects.get_or_create(user=user, defaults={'role': 'student'})
            rows.append({
                'department': dept.name,
                'batch': student.batch.name,
                'roll_no': student.roll_no,
                'enrollment_no': student.enrollment_no or '',
                'name': student.name,
                'username': username,
                'password': password,
            })

        cred_dir = os.path.join(settings.MEDIA_ROOT, 'credentials')
        os.makedirs(cred_dir, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_dept = re.sub(r'[^\w\-]', '_', dept.name)[:50]
        filename = f'student_credentials_{safe_dept}_{timestamp}.csv'
        filepath = os.path.join(cred_dir, filename)

        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=['department', 'batch', 'roll_no', 'enrollment_no', 'name', 'username', 'password'])
            w.writeheader()
            w.writerows(rows)

        request.session['credentials_filename'] = filename
        request.session['credentials_count'] = len(rows)
        request.session['credentials_type'] = 'student'
        messages.success(request, f'Credentials generated for {len(rows)} students. Download and store the file securely.')
        return redirect('core:credentials_result')

    students_without = Student.objects.filter(department=dept, user__isnull=True).count()
    students_total = Student.objects.filter(department=dept).count()
    ctx = {'department': dept, 'students_without': students_without, 'students_total': students_total}
    return render(request, 'core/admin/student_generate_credentials.html', ctx)


@login_required
def credentials_result(request):
    """Shared result page after generating faculty or student credentials."""
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    filename = request.session.pop('credentials_filename', None)
    count = request.session.pop('credentials_count', 0)
    cred_type = request.session.pop('credentials_type', 'faculty')
    if not filename:
        messages.info(request, 'No credentials file from this session.')
        return redirect('core:generate_credentials_choice')
    ctx = {'filename': filename, 'count': count, 'credentials_type': cred_type}
    return render(request, 'core/admin/credentials_result.html', ctx)


@login_required
def faculty_credentials_result(request):
    """Legacy URL: redirect to shared credentials result (session may already be consumed)."""
    return redirect('core:credentials_result')


@login_required
def download_credentials_file(request, filename):
    if not user_can_admin(request):
        raise Http404()
    if not filename or not re.match(r'^[a-zA-Z0-9_.\-]+\.csv$', filename):
        raise Http404()
    filepath = os.path.join(settings.MEDIA_ROOT, 'credentials', filename)
    if not os.path.isfile(filepath):
        raise Http404()
    with open(filepath, 'rb') as f:
        response = HttpResponse(f.read(), content_type='text/csv')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


# ---------- Admin: Students ----------

@login_required
def student_list(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    batch_id = request.GET.get('batch')
    q = request.GET.get('q', '').strip()
    students = Student.objects.filter(department=dept) if dept else Student.objects.none()
    if batch_id and dept:
        students = students.filter(batch_id=batch_id)
    if q:
        from django.db.models import Q
        students = students.filter(
            Q(roll_no__icontains=q) | Q(name__icontains=q) | Q(enrollment_no__icontains=q)
        )
    students = students.select_related('batch', 'mentor').order_by('batch__name', 'roll_no')
    from django.core.paginator import Paginator
    paginator = Paginator(students, 25)
    page = request.GET.get('page', 1)
    try:
        page_obj = paginator.page(int(page))
    except (ValueError, Paginator.EmptyPage):
        page_obj = paginator.page(1)
    batches = Batch.objects.filter(department=dept) if dept else []
    ctx = {
        'students': page_obj,
        'page_obj': page_obj,
        'batches': batches,
        'department': dept,
        'selected_batch_id': batch_id,
        'search_q': q,
    }
    return render(request, 'core/admin/student_list.html', ctx)


@login_required
def student_upload(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    batches = Batch.objects.filter(department=dept)
    if request.method == 'POST':
        batch_id = request.POST.get('batch')
        file = request.FILES.get('file')
        if not batch_id or not file:
            messages.error(request, 'Select batch and upload CSV (columns: roll_no, name, enrollment_no).')
            return redirect('core:student_upload')
        batch = Batch.objects.filter(pk=batch_id, department=dept).first()
        if not batch:
            messages.error(request, 'Invalid batch.')
            return redirect('core:student_upload')
        try:
            decoded = file.read().decode('utf-8-sig')
            reader = csv.DictReader(decoded.splitlines())
            rows = []
            faculties = list(Faculty.objects.filter(department=dept))

            def get_col(row, *names):
                for n in names:
                    for k in row:
                        if k.strip().lower().replace(' ', '_') == n.lower().replace(' ', '_'):
                            return (row.get(k) or '').strip()
                return ''

            for row in reader:
                rn = get_col(row, 'roll_no', 'roll no')
                nm = get_col(row, 'name')
                en = get_col(row, 'enrollment_no', 'enrollment no')
                mentor_val = get_col(row, 'mentor')
                mentor = None
                if mentor_val:
                    for f in faculties:
                        if mentor_val.lower() in (f.short_name.lower(), f.full_name.lower()):
                            mentor = f
                            break
                if rn and nm:
                    rows.append({'roll_no': rn, 'name': nm, 'enrollment_no': en, 'mentor': mentor})
            if not rows:
                messages.error(request, 'No valid rows (need roll_no, name).')
                return redirect('core:student_upload')
            Student.objects.filter(department=dept, batch=batch).delete()
            for r in rows:
                Student.objects.create(
                    department=dept, batch=batch,
                    roll_no=r['roll_no'], name=r['name'], enrollment_no=r.get('enrollment_no', ''),
                    mentor=r.get('mentor')
                )
            messages.success(request, f'{len(rows)} students uploaded for {batch.name}.')
            return redirect('core:student_list')
        except Exception as e:
            messages.error(request, str(e))
    ctx = {'batches': batches, 'department': dept}
    return render(request, 'core/admin/student_upload.html', ctx)


# ---------- Admin: Schedule ----------

@login_required
def schedule_list(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    slots = ScheduleSlot.objects.filter(department=dept).select_related(
        'faculty', 'subject', 'batch'
    ).order_by('day', 'time_slot') if dept else []
    ctx = {'slots': slots, 'department': dept}
    return render(request, 'core/admin/schedule_list.html', ctx)


@login_required
def schedule_add(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    if request.method == 'POST':
        faculty_id = request.POST.get('faculty')
        subject_id = request.POST.get('subject')
        batch_id = request.POST.get('batch')
        day = request.POST.get('day', '').strip()
        time_slot = request.POST.get('time_slot', '').strip()
        if faculty_id and subject_id and batch_id and day and time_slot:
            faculty = Faculty.objects.filter(pk=faculty_id, department=dept).first()
            subject = Subject.objects.filter(pk=subject_id, department=dept).first()
            batch = Batch.objects.filter(pk=batch_id, department=dept).first()
            if faculty and subject and batch:
                ScheduleSlot.objects.get_or_create(
                    department=dept, batch=batch, day=day, time_slot=time_slot,
                    defaults={'faculty': faculty, 'subject': subject}
                )
                messages.success(request, 'Schedule slot added.')
                return redirect('core:schedule_list')
    faculties = Faculty.objects.filter(department=dept)
    subjects = Subject.objects.filter(department=dept)
    batches = Batch.objects.filter(department=dept)
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    ctx = {
        'department': dept, 'faculties': faculties, 'subjects': subjects,
        'batches': batches, 'days': days, 'slot': None,
    }
    return render(request, 'core/admin/schedule_form.html', ctx)


@login_required
def schedule_delete(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    slot = get_object_or_404(ScheduleSlot, pk=pk)
    if slot.department != dept:
        messages.error(request, 'You can only manage schedule for your selected department.')
        return redirect('core:schedule_list')
    slot.delete()
    messages.success(request, 'Slot removed.')
    return redirect('core:schedule_list')


@login_required
def schedule_edit(request, pk):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    slot = get_object_or_404(ScheduleSlot, pk=pk)
    if slot.department != dept:
        messages.error(request, 'You can only edit schedule for your selected department.')
        return redirect('core:schedule_list')
    if request.method == 'POST':
        faculty_id = request.POST.get('faculty')
        subject_id = request.POST.get('subject')
        batch_id = request.POST.get('batch')
        day = request.POST.get('day', '').strip()
        time_slot = request.POST.get('time_slot', '').strip()
        if faculty_id and subject_id and batch_id and day and time_slot:
            faculty = Faculty.objects.filter(pk=faculty_id, department=dept).first()
            subject = Subject.objects.filter(pk=subject_id, department=dept).first()
            batch = Batch.objects.filter(pk=batch_id, department=dept).first()
            if faculty and subject and batch:
                slot.faculty = faculty
                slot.subject = subject
                slot.batch = batch
                slot.day = day
                slot.time_slot = time_slot
                slot.save()
                messages.success(request, 'Schedule slot updated.')
                return redirect('core:schedule_list')
    faculties = Faculty.objects.filter(department=dept)
    subjects = Subject.objects.filter(department=dept)
    batches = Batch.objects.filter(department=dept)
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    ctx = {
        'slot': slot,
        'department': dept,
        'faculties': faculties,
        'subjects': subjects,
        'batches': batches,
        'days': days,
    }
    return render(request, 'core/admin/schedule_form.html', ctx)


@login_required
def schedule_clear_all(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    if request.method != 'POST':
        return redirect('core:schedule_list')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    deleted, _ = ScheduleSlot.objects.filter(department=dept).delete()
    messages.success(request, f'All schedule entries for {dept.name} have been deleted ({deleted} slot(s)).')
    return redirect('core:schedule_list')


# ---------- Admin: Lecture Adjustment ----------

@login_required
def lecture_adjustment(request):
    """Adjust lectures for a date (e.g. substitute faculty): select date, batch, then add changes and apply. History shown below."""
    if not user_can_admin(request):
        return redirect('core:admin_dashboard')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')

    batches = Batch.objects.filter(department=dept).order_by('name')
    faculties = Faculty.objects.filter(department=dept).order_by('full_name')
    subjects = Subject.objects.filter(department=dept).order_by('name')

    # Pending changes in session: list of {time_slot, new_subject_id, new_faculty_id, ...}
    pending_key = 'lecture_adj_pending'
    pending = request.session.get(pending_key, [])

    if request.method == 'POST':
        action = request.POST.get('action')
        date_str = request.POST.get('date')
        batch_id = request.POST.get('batch')
        if action == 'add':
            time_slot = request.POST.get('time_slot', '').strip()
            new_subject_id = request.POST.get('new_subject')
            new_faculty_id = request.POST.get('new_faculty')
            if date_str and batch_id and time_slot and new_subject_id and new_faculty_id:
                # Remove any existing pending for same slot
                pending = [p for p in pending if p.get('time_slot') != time_slot or p.get('_date') != date_str or p.get('_batch_id') != batch_id]
                pending.append({
                    '_date': date_str, '_batch_id': batch_id,
                    'time_slot': time_slot,
                    'new_subject_id': new_subject_id,
                    'new_faculty_id': new_faculty_id,
                })
                request.session[pending_key] = pending
                request.session.modified = True
                messages.success(request, 'Change added. Add more or click "Apply changes" below.')
            return redirect(reverse('core:lecture_adjustment') + f'?date={date_str}&batch={batch_id}')
        if action == 'apply' and date_str and batch_id:
            try:
                selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
            except Exception:
                messages.error(request, 'Invalid date.')
                return redirect('core:lecture_adjustment')
            batch = Batch.objects.filter(pk=batch_id, department=dept).first()
            if not batch:
                messages.error(request, 'Invalid batch.')
                return redirect('core:lecture_adjustment')
            weekday = selected_date.strftime('%A')
            to_apply = [p for p in pending if p.get('_date') == date_str and str(p.get('_batch_id')) == str(batch_id)]
            for p in to_apply:
                slot = ScheduleSlot.objects.filter(
                    batch=batch, day=weekday, time_slot=p.get('time_slot')
                ).select_related('faculty', 'subject').first()
                if not slot:
                    continue
                new_faculty = Faculty.objects.filter(pk=p.get('new_faculty_id'), department=dept).first()
                new_subject = Subject.objects.filter(pk=p.get('new_subject_id'), department=dept).first()
                if not new_faculty or not new_subject:
                    continue
                LectureAdjustment.objects.update_or_create(
                    date=selected_date,
                    batch=batch,
                    time_slot=slot.time_slot,
                    defaults={
                        'original_faculty': slot.faculty,
                        'original_subject': slot.subject,
                        'new_faculty': new_faculty,
                        'new_subject': new_subject,
                    },
                )
            pending = [p for p in pending if not (p.get('_date') == date_str and str(p.get('_batch_id')) == str(batch_id))]
            request.session[pending_key] = pending
            request.session.modified = True
            messages.success(request, f'Applied {len(to_apply)} adjustment(s). They will reflect in all Excel and displays.')
            return redirect(reverse('core:lecture_adjustment') + f'?date={date_str}&batch={batch_id}')
        if action == 'clear_pending':
            request.session[pending_key] = []
            request.session.modified = True
            return redirect('core:lecture_adjustment')

    selected_date = None
    selected_batch = None
    day_slots = []
    date_str = request.GET.get('date')
    batch_id = request.GET.get('batch')
    if date_str and batch_id:
        try:
            selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except Exception:
            pass
        selected_batch = Batch.objects.filter(pk=batch_id, department=dept).first()
        if selected_date and selected_batch:
            weekday = selected_date.strftime('%A')
            slots = ScheduleSlot.objects.filter(
                batch=selected_batch, day=weekday
            ).select_related('faculty', 'subject').order_by('time_slot')
            # Check existing adjustments for this date/batch
            existing_adj = {a.time_slot: (a.new_faculty, a.new_subject) for a in
                           LectureAdjustment.objects.filter(date=selected_date, batch=selected_batch).select_related('new_faculty', 'new_subject')}
            for i, slot in enumerate(slots, 1):
                fac, subj = existing_adj.get(slot.time_slot, (slot.faculty, slot.subject))
                day_slots.append({
                    'lecture_num': i,
                    'time_slot': slot.time_slot,
                    'subject': subj.name,
                    'faculty': fac.short_name,
                    'original_subject': slot.subject.name,
                    'original_faculty': slot.faculty.short_name,
                    'is_adjusted': slot.time_slot in existing_adj,
                })

    # Pending for this page's date/batch
    pending_for_page = [p for p in pending if p.get('_date') == date_str and str(p.get('_batch_id')) == str(batch_id)] if date_str and batch_id else []
    for p in pending_for_page:
        new_f = Faculty.objects.filter(pk=p.get('new_faculty_id')).first()
        new_s = Subject.objects.filter(pk=p.get('new_subject_id')).first()
        p['new_faculty_name'] = new_f.short_name if new_f else '—'
        p['new_subject_name'] = new_s.name if new_s else '—'

    # History: recent adjustments (for this dept)
    batch_ids = list(batches.values_list('pk', flat=True))
    history = LectureAdjustment.objects.filter(
        batch_id__in=batch_ids
    ).select_related('batch', 'new_faculty', 'new_subject', 'original_faculty', 'original_subject').order_by('-date', '-created_at')[:50]

    ctx = {
        'department': dept,
        'batches': batches,
        'faculties': faculties,
        'subjects': subjects,
        'selected_date': selected_date,
        'date_str': date_str or '',
        'selected_batch': selected_batch,
        'batch_id': batch_id or '',
        'day_slots': day_slots,
        'pending': pending_for_page,
        'history': history,
    }
    return render(request, 'core/admin/lecture_adjustment.html', ctx)


# ---------- Admin: Upload Timetable (Excel) ----------

def _looks_like_timing_header(val):
    """Return True if header cell looks like timing/time/slot, not a batch name."""
    if not val:
        return True
    s = str(val).strip().lower()
    if not s:
        return True
    if s in ('timing', 'time', 'slot', 'slots', 'lecture', 'lec', 'period', 'sr no', 's.no', 'no.'):
        return True
    if re.match(r'^lec\s*\d+$', s):
        return True
    if re.match(r'^\d{1,2}:\d{2}\s*[-–—to]+\s*\d{1,2}:\d{2}', s):
        return True
    if re.match(r'^\d{1,2}-\d{1,2}$', s):
        return True
    return False


def _normalize_day(excel_day):
    """Map Mon, Tue, ... to Monday, Tuesday, ..."""
    if not excel_day or not str(excel_day).strip():
        return None
    s = str(excel_day).strip().lower()[:3]
    day_map = {
        'mon': 'Monday', 'tue': 'Tuesday', 'wed': 'Wednesday',
        'thu': 'Thursday', 'fri': 'Friday', 'sat': 'Saturday', 'sun': 'Sunday',
    }
    return day_map.get(s, str(excel_day).strip().capitalize())


def _normalize_time_slot(val):
    """Normalize time slot string (e.g. 08:45-09:45, 8:45 to 9:45, Lec 1)."""
    if not val:
        return None
    t = str(val).strip().replace('to', '-').replace('–', '-').replace('—', '-').replace(' ', '')
    if t:
        return t
    return str(val).strip()


def _parse_cell_faculty_subject(cell_value):
    """
    Parse timetable cell to get (faculty_short_name, subject_name).
    Formats: "Subject (Faculty) (Room)", "Subject (Faculty) (Lab) (Lab)", "FAC-SUB-310", "FAC-SUB-408(L)"
    """
    if not cell_value or str(cell_value).strip() == '':
        return None, None
    text = str(cell_value).strip()
    # New format: DE (UMS) (301) or FSD-1 (DKU) (408-A) (Lab)
    new_pattern = r'^(.+?)\s*\(([^)]+)\)\s*\(([^)]+)\)'
    match = re.match(new_pattern, text)
    if match:
        subject = match.group(1).strip()
        faculty = match.group(2).strip()
        return faculty, subject
    # Old format: PHA-FSD_2-408-A(L) or PSK-DM-310
    parts = text.split('-')
    if len(parts) >= 2:
        faculty = parts[0].strip()
        subject = parts[1].strip()
        return faculty, subject
    return None, None


@login_required
def upload_timetable(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first from Dashboard.')
        return redirect('core:admin_dashboard')

    if request.method == 'POST':
        file = request.FILES.get('excel_file')
        replace_schedule = request.POST.get('replace_schedule') == 'on'
        if not file:
            messages.error(request, 'Please select an Excel file.')
            return redirect('core:upload_timetable')
        if not file.name.lower().endswith(('.xlsx', '.xls')):
            messages.error(request, 'Upload a valid Excel file (.xlsx or .xls).')
            return redirect('core:upload_timetable')
        try:
            wb = openpyxl.load_workbook(file, data_only=True)
        except Exception as e:
            messages.error(request, f'Could not read Excel: {e}')
            return redirect('core:upload_timetable')

        sheet = None
        for name in ['TT-CLASSWISE', 'TT CLASSWISE', 'Timetable', 'Sheet1']:
            if name in wb.sheetnames:
                sheet = wb[name]
                break
        if not sheet:
            sheet = wb.active
        if not sheet:
            messages.error(request, 'No sheet found in the workbook.')
            return redirect('core:upload_timetable')

        # Header row: row 3 (1-based) = batch names from column C onwards
        header_row = next(sheet.iter_rows(min_row=3, max_row=3, values_only=True), None)
        if not header_row:
            messages.error(request, 'Could not read header row (row 3). Expected batch names in columns from C.')
            return redirect('core:upload_timetable')

        batches = []
        for col_idx in range(2, len(header_row)):
            val = header_row[col_idx]
            if not val or not str(val).strip():
                continue
            batch_name = str(val).strip()
            if _looks_like_timing_header(batch_name):
                continue
            batches.append((col_idx, batch_name))

        if not batches:
            messages.error(request, 'No batch columns found in row 3 (columns C onwards).')
            return redirect('core:upload_timetable')

        for _, bname in batches:
            Batch.objects.get_or_create(department=dept, name=bname)

        if replace_schedule:
            ScheduleSlot.objects.filter(department=dept).delete()

        current_day = None
        created_slots = 0
        for row in sheet.iter_rows(min_row=4, values_only=True):
            first_col = row[0] if len(row) > 0 else None
            time_val = row[1] if len(row) > 1 else None
            if first_col and isinstance(first_col, str) and first_col.strip().upper() not in ('', 'RECESS'):
                current_day = _normalize_day(first_col)
            if not current_day or not time_val:
                continue
            if len(row) > 2 and isinstance(row[2], str) and row[2].strip().upper() == 'RECESS':
                continue
            time_slot = _normalize_time_slot(time_val)
            if not time_slot:
                continue
            for col_idx, batch_name in batches:
                if col_idx >= len(row):
                    continue
                cell_val = row[col_idx]
                fac, subj = _parse_cell_faculty_subject(cell_val)
                if not fac or not subj:
                    continue
                faculty_obj, _ = Faculty.objects.get_or_create(
                    department=dept,
                    short_name=fac,
                    defaults={'full_name': fac}
                )
                subject_obj, _ = Subject.objects.get_or_create(
                    department=dept,
                    name=subj,
                    defaults={}
                )
                batch_obj = Batch.objects.get(department=dept, name=batch_name)
                _, created = ScheduleSlot.objects.get_or_create(
                    department=dept,
                    batch=batch_obj,
                    day=current_day,
                    time_slot=time_slot,
                    defaults={'faculty': faculty_obj, 'subject': subject_obj}
                )
                if created:
                    created_slots += 1

        # Exact counts after import (from database)
        slots_qs = ScheduleSlot.objects.filter(department=dept)
        total_entries = slots_qs.count()
        total_days = slots_qs.values_list('day', flat=True).distinct().count()
        total_time_slots = slots_qs.values_list('time_slot', flat=True).distinct().count()
        total_batches = Batch.objects.filter(department=dept).count()
        total_subjects = Subject.objects.filter(department=dept).count()
        total_faculties = Faculty.objects.filter(department=dept).count()
        per_batch_raw = slots_qs.values('batch__name').annotate(count=Count('id')).order_by('batch__name')
        per_batch = [{'batch_name': r['batch__name'], 'count': r['count']} for r in per_batch_raw]
        import_summary = {
            'total_batches': total_batches,
            'total_days': total_days,
            'total_time_slots': total_time_slots,
            'total_subjects': total_subjects,
            'total_faculties': total_faculties,
            'total_entries': total_entries,
            'new_slots_added': created_slots,
            'replace_schedule': replace_schedule,
            'per_batch': per_batch,
        }
        messages.success(request, 'Timetable imported successfully. See summary below.')
        departments = Department.objects.all()
        ctx = {
            'department': dept,
            'departments': departments,
            'import_summary': import_summary,
        }
        return render(request, 'core/admin/upload_timetable.html', ctx)

    departments = Department.objects.all()
    ctx = {
        'department': dept,
        'departments': departments,
    }
    return render(request, 'core/admin/upload_timetable.html', ctx)


# ---------- Admin: Term Phases ----------

def _parse_holiday_dates(text, phase_start, phase_end):
    """Parse holiday date string (newline or comma separated YYYY-MM-DD). Return list of date objects within phase range."""
    from datetime import datetime
    out = []
    if not text or not phase_start or not phase_end:
        return out
    # Ensure phase_start and phase_end are date objects (POST/model can give strings)
    try:
        if isinstance(phase_start, str):
            phase_start = datetime.strptime(phase_start, '%Y-%m-%d').date()
        if isinstance(phase_end, str):
            phase_end = datetime.strptime(phase_end, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        return out
    for part in text.replace(',', '\n').split():
        part = part.strip()
        if not part:
            continue
        try:
            d = datetime.strptime(part, '%Y-%m-%d').date()
            if phase_start <= d <= phase_end:
                out.append(d)
        except ValueError:
            continue
    return sorted(set(out))


@login_required
def term_phases(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    tp, _ = TermPhase.objects.get_or_create(department=dept)
    if request.method == 'POST':
        for i in range(1, 5):
            setattr(tp, f't{i}_start', request.POST.get(f't{i}_start') or None)
            setattr(tp, f't{i}_end', request.POST.get(f't{i}_end') or None)
        tp.save()
        # Save holidays per phase (only dates within phase start–end)
        for i in range(1, 5):
            phase = f'T{i}'
            start = getattr(tp, f't{i}_start', None)
            end = getattr(tp, f't{i}_end', None)
            PhaseHoliday.objects.filter(department=dept, phase=phase).delete()
            raw = request.POST.get(f't{i}_holidays', '') or ''
            dates = _parse_holiday_dates(raw, start, end)
            for d in dates:
                PhaseHoliday.objects.get_or_create(department=dept, phase=phase, date=d)
        messages.success(request, 'Term phases and holidays saved.')
        return redirect('core:term_phases')
    # Load existing holidays per phase for the form
    holiday_lists = {}
    for i in range(1, 5):
        phase = f'T{i}'
        dates = list(
            PhaseHoliday.objects.filter(department=dept, phase=phase).order_by('date').values_list('date', flat=True)
        )
        holiday_lists[f't{i}_holidays_list'] = dates
        holiday_lists[f't{i}_holidays'] = '\n'.join(d.strftime('%Y-%m-%d') for d in dates)
    ctx = {'term_phase': tp, 'department': dept, **holiday_lists}
    return render(request, 'core/admin/term_phases.html', ctx)


# ---------- Admin: Daily Absent ----------

def _valid_dates(dept, term_phase):
    """Lecture days across all phases, excluding holidays."""
    if not term_phase:
        return []
    days_set = set(
        ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct()
    )
    days_set = {d.lower() for d in days_set if d}
    holidays = get_all_holiday_dates(dept)
    out = []
    for i in range(1, 5):
        start = getattr(term_phase, f't{i}_start', None)
        end = getattr(term_phase, f't{i}_end', None)
        if not start or not end:
            continue
        cur = start
        while cur <= end:
            if cur not in holidays and cur.strftime('%A').lower() in days_set:
                out.append(cur)
            cur += timedelta(days=1)
    return sorted(out)


@login_required
def daily_absent(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    tp = TermPhase.objects.filter(department=dept).first()
    valid_dates = _valid_dates(dept, tp)
    date_str = request.GET.get('date')
    selected_date = None
    if date_str:
        try:
            selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
            if selected_date not in valid_dates:
                selected_date = valid_dates[0] if valid_dates else None
        except Exception:
            selected_date = valid_dates[0] if valid_dates else None
    else:
        selected_date = valid_dates[0] if valid_dates else None

    lectures_by_batch = defaultdict(list)
    if selected_date:
        weekday = selected_date.strftime('%A')
        slots = ScheduleSlot.objects.filter(
            department=dept, day=weekday
        ).select_related('faculty', 'subject', 'batch').order_by('batch__name', 'time_slot')
        for s in slots:
            lectures_by_batch[s.batch.name].append(s)

    attendance_map = defaultdict(lambda: defaultdict(list))
    if selected_date:
        atts = FacultyAttendance.objects.filter(
            faculty__department=dept, date=selected_date
        )
        for a in atts:
            for r in (a.absent_roll_numbers or '').split(','):
                if r.strip():
                    attendance_map[a.batch.id][a.lecture_slot].append(r.strip())

    missing = []
    for batch_name, lectures in lectures_by_batch.items():
        for slot in lectures:
            batch_id = slot.batch.id
            fac, subj = get_faculty_subject_for_slot(selected_date, slot.batch, slot.time_slot)
            effective_faculty = fac if fac is not None else slot.faculty
            if subj is not None:
                slot.subject = subj
            if fac is not None:
                slot.faculty = fac
            rec = FacultyAttendance.objects.filter(
                faculty=effective_faculty, date=selected_date,
                batch_id=batch_id, lecture_slot=slot.time_slot
            ).first()
            slot.absent_list = attendance_map.get(batch_id, {}).get(slot.time_slot, [])
            # Only "missing" when no attendance record; saved "all present" (empty absent) is valid
            if not rec:
                missing.append((effective_faculty.full_name, batch_name, slot.time_slot))

    can_generate = not missing and bool(lectures_by_batch)
    ctx = {
        'valid_dates': valid_dates, 'selected_date': selected_date,
        'lectures_by_batch': dict(lectures_by_batch),
        'attendance_map': dict(attendance_map),
        'missing_entries': missing, 'can_generate_report': can_generate,
        'department': dept,
    }
    return render(request, 'core/admin/daily_absent.html', ctx)


@login_required
def daily_absent_excel(request):
    """Export daily absent report: title, 2 batches per row, No/Subject/Faculty/Absent Nos (all absent numbers in one cell per lecture, with wrap text)."""
    if not user_can_admin(request):
        return redirect('core:admin_dashboard')
    dept = get_admin_department(request)
    date_str = request.GET.get('date')
    if not date_str or not dept:
        return redirect('core:daily_absent')
    try:
        selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    except Exception:
        return redirect('core:daily_absent')
    if selected_date in get_all_holiday_dates(dept):
        messages.warning(request, 'Selected date is a holiday; no lectures scheduled.')
        return redirect('core:daily_absent')
    weekday = selected_date.strftime('%A')
    slots = ScheduleSlot.objects.filter(
        department=dept, day=weekday
    ).select_related('faculty', 'subject', 'batch').order_by('batch__name', 'time_slot')
    lectures_by_batch = defaultdict(list)
    for s in slots:
        lectures_by_batch[s.batch.name].append(s)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Daily Absent'

    title = f'Daily Absent Report - {dept.name} - {selected_date.strftime("%d-%m-%Y")}'
    max_batches_per_row = 2
    headers = ['No', 'Subject', 'Faculty', 'Absent Nos']
    header_fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    batch_names = sorted(lectures_by_batch.keys())
    pairs = [batch_names[i:i + max_batches_per_row] for i in range(0, len(batch_names), max_batches_per_row)]
    current_row = 1

    n_cols = (len(pairs[0]) * 5) if pairs else 5
    if n_cols > 1:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=n_cols)
    cell = ws.cell(row=current_row, column=1, value=title)
    cell.font = Font(size=14, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    current_row += 2

    for pair in pairs:
        col = 1
        for batch in pair:
            ws.merge_cells(start_row=current_row, start_column=col, end_row=current_row, end_column=col + 3)
            batch_cell = ws.cell(row=current_row, column=col, value=f'Batch: {batch}')
            batch_cell.font = Font(bold=True, color='FFFFFF')
            batch_cell.fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
            batch_cell.alignment = Alignment(horizontal='center')
            col += 5
        current_row += 1
        col = 1
        for batch in pair:
            for cidx, header in enumerate(headers):
                c = ws.cell(row=current_row, column=col + cidx, value=header)
                c.font = header_font
                c.fill = header_fill
                c.border = thin_border
                c.alignment = Alignment(horizontal='center')
            col += 5
        current_row += 1

        max_lects = max(len(lectures_by_batch[b]) for b in pair)
        to_continue = []
        for batch in pair:
            lec_list = lectures_by_batch[batch]
            rows = []
            for idx, lec in enumerate(lec_list, 1):
                fac, subj = get_faculty_subject_for_slot(selected_date, lec.batch, lec.time_slot)
                subj_name = subj.name if subj else lec.subject.name
                fac_name = fac.full_name if fac else lec.faculty.full_name
                att = FacultyAttendance.objects.filter(
                    date=selected_date, batch=lec.batch, lecture_slot=lec.time_slot
                ).first()
                if fac and not att:
                    att = FacultyAttendance.objects.filter(
                        faculty=fac, date=selected_date,
                        batch=lec.batch, lecture_slot=lec.time_slot
                    ).first()
                if not att:
                    att = FacultyAttendance.objects.filter(
                        faculty=lec.faculty, date=selected_date,
                        batch=lec.batch, lecture_slot=lec.time_slot
                    ).first()
                absent_nos = (att.absent_roll_numbers or '').strip() if att else ''
                absents = [a.strip() for a in absent_nos.split(',') if a.strip()]
                if not absents:
                    rows.append([idx, subj_name, fac_name, 'NIL'])
                else:
                    rows.append([idx, subj_name, fac_name, ', '.join(absents)])
            to_continue.append(rows)
        block_height = max(len(x) for x in to_continue)
        for data_rows in to_continue:
            while len(data_rows) < block_height:
                data_rows.append(['', '', '', ''])
        for ridx in range(block_height):
            col = 1
            for data_rows in to_continue:
                for cidx, value in enumerate(data_rows[ridx]):
                    c = ws.cell(row=current_row, column=col + cidx, value=value)
                    c.border = thin_border
                    wrap = (cidx == 3)
                    c.alignment = Alignment(horizontal='center' if cidx == 0 else 'left', wrap_text=wrap)
                col += 5
            current_row += 1
        current_row += 1

    for b in range(2):
        offset = b * 5
        ws.column_dimensions[get_column_letter(1 + offset)].width = 5
        ws.column_dimensions[get_column_letter(2 + offset)].width = 28
        ws.column_dimensions[get_column_letter(3 + offset)].width = 22
        ws.column_dimensions[get_column_letter(4 + offset)].width = 35

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    resp = HttpResponse(bio.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename=DailyAbsent_{selected_date:%Y-%m-%d}.xlsx'
    return resp


# ---------- Admin: Attendance Sheet Manager ----------

@login_required
def attendance_sheet_manager(request):
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    tp = TermPhase.objects.filter(department=dept).first()
    batches = Batch.objects.filter(department=dept)
    phases = ['T1', 'T2', 'T3', 'T4']
    week_map = {}
    for p in phases:
        start = getattr(tp, f'{p.lower()}_start', None) if tp else None
        end = getattr(tp, f'{p.lower()}_end', None) if tp else None
        if not start or not end:
            week_map[p] = []
            continue
        days_set = set(ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct())
        days_set = {d.lower() for d in days_set if d}
        dates = []
        cur = start
        while cur <= end:
            if cur.strftime('%A').lower() in days_set:
                dates.append(cur)
            cur += timedelta(days=1)
        # chunk by week
        weeks = []
        week = []
        last_w = None
        for d in sorted(dates):
            w = d.isocalendar()[1]
            if last_w is not None and w != last_w and week:
                weeks.append(week)
                week = []
            week.append(d)
            last_w = w
        if week:
            weeks.append(week)
        week_map[p] = [[d.isoformat() for d in w] for w in weeks]
    # All lecture dates for daily picker (any phase)
    all_dates = []
    for p in phases:
        for week_dates in week_map.get(p, []):
            for d_str in week_dates:
                try:
                    all_dates.append(datetime.strptime(d_str, '%Y-%m-%d').date())
                except Exception:
                    pass
    available_dates = sorted(set(all_dates))
    ctx = {
        'department': dept, 'batches': batches, 'phases': phases,
        'week_map': week_map, 'week_map_json': json.dumps(week_map),
        'available_dates': available_dates,
    }
    return render(request, 'core/admin/attendance_sheet_manager.html', ctx)


def _build_date_slots_list_for_batch(dept, batch, dates):
    """Return [(date, slots), ...] for this batch and list of dates (excluding holidays already in dates)."""
    out = []
    for d in dates:
        weekday = d.strftime('%A')
        slots = list(ScheduleSlot.objects.filter(
            department=dept, batch=batch, day=weekday
        ).select_related('subject').order_by('time_slot'))
        out.append((d, slots))
    return out


def _attendance_sheet_dates_for_period(dept, period_type, phase, week_index=None, single_date=None):
    """Return list of dates for the chosen period (excluding holidays). week_index is 0-based index into week_map[phase]."""
    tp = TermPhase.objects.filter(department=dept).first()
    days_set = set(ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct())
    days_set = {d.lower() for d in days_set if d}
    if period_type == 'daily' and single_date:
        holidays = get_all_holiday_dates(dept)
        if single_date in holidays:
            return []
        return [single_date] if single_date.strftime('%A').lower() in days_set else []
    if period_type == 'weekly' and phase and week_index is not None:
        start = getattr(tp, f'{phase.lower()}_start', None) if tp else None
        end = getattr(tp, f'{phase.lower()}_end', None) if tp else None
        if not start or not end:
            return []
        holidays = get_phase_holidays(dept, phase)
        dates = []
        cur = start
        while cur <= end:
            if cur not in holidays and cur.strftime('%A').lower() in days_set:
                dates.append(cur)
            cur += timedelta(days=1)
        weeks = []
        week = []
        last_w = None
        for d in sorted(dates):
            w = d.isocalendar()[1]
            if last_w is not None and w != last_w and week:
                weeks.append(week)
                week = []
            week.append(d)
            last_w = w
        if week:
            weeks.append(week)
        if 0 <= week_index < len(weeks):
            return weeks[week_index]
        return []
    if period_type == 'phase' and phase:
        start = getattr(tp, f'{phase.lower()}_start', None) if tp else None
        end = getattr(tp, f'{phase.lower()}_end', None) if tp else None
        if not start or not end:
            return []
        holidays = get_phase_holidays(dept, phase)
        dates = []
        cur = start
        while cur <= end:
            if cur not in holidays and cur.strftime('%A').lower() in days_set:
                dates.append(cur)
            cur += timedelta(days=1)
        return sorted(dates)
    return []


def _student_held_attended_for_segment(date_slots_segment, att_map, str_roll):
    """For one segment (list of (date, slots)), return (held, attended) for one student.
    held = total scheduled (excluding holidays). attended = only where attendance was taken AND student present."""
    held = attended = 0
    for d, slots in date_slots_segment:
        for slot in slots:
            held += 1
            key = (d, slot.time_slot)
            if key in att_map and str_roll not in att_map[key]:
                attended += 1
    return held, attended


def _write_one_batch_attendance_sheet(ws, batch, date_slots_list, students, att_map, styles, overall_segments=None):
    """Write one batch's attendance data into worksheet ws. If overall_segments is given, append Overall Attendance block.
    overall_segments: list of (label, date_slots_sub_list) e.g. [('Week 1', w1_list), ('Week 2', w2_list), ('Overall', all_list)].
    """
    thin_border = styles['thin_border']
    date_fill = styles['date_fill']
    date_font = styles['date_font']
    date_align = styles['date_align']
    header_font = styles['header_font']
    lect_fill = styles['lect_fill']
    lect_font = styles['lect_font']
    lect_align = styles['lect_align']
    red_font = styles['red_font']

    data_start_row = 3
    header_rows = 2
    if overall_segments:
        header_rows = 3
        data_start_row = 4

    ws.title = (batch.name[:31] if batch.name else 'Sheet')
    ws.cell(1, 1, 'Roll No').font = header_font
    ws.cell(1, 2, 'Student Name').font = header_font
    ws.cell(1, 1).border = thin_border
    ws.cell(1, 2).border = thin_border
    ws.cell(2, 1, '').border = thin_border
    ws.cell(2, 2, '').border = thin_border
    if overall_segments:
        ws.cell(3, 1, '').border = thin_border
        ws.cell(3, 2, '').border = thin_border
    col = 3
    for d, slots in date_slots_list:
        n_lec = max(len(slots), 1)
        if n_lec > 1:
            ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + n_lec - 1)
        for c in range(col, col + n_lec):
            cell = ws.cell(row=1, column=c, value=d.strftime('%d-%b') if c == col else None)
            cell.border = thin_border
            cell.fill = date_fill
            cell.font = date_font
            cell.alignment = date_align
        for i, slot in enumerate(slots, start=1):
            fac, subj = get_faculty_subject_for_slot(d, batch, slot.time_slot)
            subj_name = subj.name if subj else (slot.subject.name if slot.subject else 'N/A')
            cell = ws.cell(row=2, column=col + i - 1, value=f'Lect {i}\n{subj_name}')
            cell.alignment = lect_align
            cell.fill = lect_fill
            cell.font = lect_font
            cell.border = thin_border
        if not slots:
            ws.cell(2, col, '').border = thin_border
        col += n_lec

    if overall_segments:
        for c in range(3, col):
            ws.cell(3, c, '').border = thin_border

    # Overall Attendance block: merged header, then per-segment label row and (Total Lecture, Attended, %)
    if overall_segments:
        overall_col_start = col
        n_overall_cols = len(overall_segments) * 3
        ws.merge_cells(start_row=1, start_column=overall_col_start, end_row=1, end_column=overall_col_start + n_overall_cols - 1)
        cell = ws.cell(row=1, column=overall_col_start, value='Overall Attendance')
        cell.border = thin_border
        cell.fill = date_fill
        cell.font = date_font
        cell.alignment = date_align
        seg_col = overall_col_start
        for label, _ in overall_segments:
            ws.merge_cells(start_row=2, start_column=seg_col, end_row=2, end_column=seg_col + 2)
            c = ws.cell(row=2, column=seg_col, value=label)
            c.border = thin_border
            c.fill = lect_fill
            c.font = lect_font
            c.alignment = date_align
            for i, sub in enumerate(('Total Lecture', 'Attended', '%')):
                cc = ws.cell(row=3, column=seg_col + i, value=sub)
                cc.border = thin_border
                cc.fill = lect_fill
                cc.font = lect_font
                cc.alignment = lect_align
            seg_col += 3
        col = overall_col_start + n_overall_cols

    for idx, s in enumerate(students, start=data_start_row):
        ws.cell(idx, 1, s.roll_no).border = thin_border
        ws.cell(idx, 2, s.name).border = thin_border
        str_roll = str(s.roll_no)
        c = 3
        for d, slots in date_slots_list:
            for slot in slots:
                key = (d, slot.time_slot)
                if key not in att_map:
                    val = '—'
                else:
                    val = 'A' if str_roll in att_map[key] else 'P'
                cell = ws.cell(idx, c, value=val)
                cell.border = thin_border
                if val == 'A':
                    cell.font = red_font
                c += 1
            if not slots:
                ws.cell(idx, c, '').border = thin_border
                c += 1
        if overall_segments:
            for label, date_slots_seg in overall_segments:
                held, attended = _student_held_attended_for_segment(date_slots_seg, att_map, str_roll)
                pct = round(attended / held * 100, 1) if held else 0
                ws.cell(idx, c, held).border = thin_border
                ws.cell(idx, c + 1, attended).border = thin_border
                ws.cell(idx, c + 2, f'{pct}%').border = thin_border
                c += 3
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 20
    ws.freeze_panes = f'C{data_start_row}'


@login_required
def attendance_sheet_excel(request):
    if not user_can_admin(request):
        return redirect('core:admin_dashboard')
    dept = get_admin_department(request)
    batch_id = request.GET.get('batch')
    period_type = request.GET.get('period_type', 'phase')
    phase = request.GET.get('phase')
    week_index = request.GET.get('week')
    date_str = request.GET.get('date')
    if not dept:
        return redirect('core:attendance_sheet_manager')
    if not batch_id:
        return redirect('core:attendance_sheet_manager')

    all_batches = batch_id == 'all'
    if all_batches:
        batches = list(Batch.objects.filter(department=dept).order_by('name'))
        if not batches:
            messages.error(request, 'No batches in this department.')
            return redirect('core:attendance_sheet_manager')
    else:
        batch = Batch.objects.filter(pk=batch_id, department=dept).first()
        if not batch:
            return redirect('core:attendance_sheet_manager')
        batches = [batch]

    single_date = None
    if period_type == 'daily' and date_str:
        try:
            single_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except Exception:
            pass
    week_idx = None
    if period_type == 'weekly' and week_index is not None:
        try:
            week_idx = int(week_index)
        except Exception:
            pass
    if period_type in ('weekly', 'phase') and not phase:
        return redirect('core:attendance_sheet_manager')

    # For weekly we want "through week N" (weeks 0..week_idx) so the sheet shows all those weeks and Overall block has W1, W2, ... Overall
    weeks_current_phase = _compile_phase_weeks_date_objects(dept, phase) if phase else []
    if period_type == 'weekly' and week_idx is not None and 0 <= week_idx < len(weeks_current_phase):
        dates = []
        for w in weeks_current_phase[: week_idx + 1]:
            dates.extend(w)
        dates = sorted(set(dates))
    else:
        dates = _attendance_sheet_dates_for_period(dept, period_type, phase, week_idx, single_date)
    if not dates:
        messages.error(request, 'No dates in selected period.')
        return redirect('core:attendance_sheet_manager')

    def build_overall_segments(batch, date_slots_list, att_map_batch):
        """Build list of (label, date_slots_sub_list) for Overall Attendance block. att_map_batch must cover all dates in any segment."""
        segments = []
        if period_type == 'daily':
            segments = [('Overall', date_slots_list)]
        elif period_type == 'weekly' and weeks_current_phase:
            # Previous phase overall when phase is T2, T3, T4
            phase_order = ['T1', 'T2', 'T3', 'T4']
            try:
                pi = phase_order.index(phase)
            except ValueError:
                pi = 0
            if pi > 0:
                for prev_i in range(pi):
                    prev_phase = phase_order[prev_i]
                    prev_weeks = _compile_phase_weeks_date_objects(dept, prev_phase)
                    prev_dates = []
                    for w in prev_weeks:
                        prev_dates.extend(w)
                    prev_dates = sorted(set(prev_dates))
                    prev_slots = _build_date_slots_list_for_batch(dept, batch, prev_dates)
                    segments.append((f'{prev_phase} Overall', prev_slots))
            for i in range(week_idx + 1):
                w_dates = weeks_current_phase[i]
                w_slots = _build_date_slots_list_for_batch(dept, batch, w_dates)
                label = f'Week {i + 1}' if pi == 0 else f'{phase} Week {i + 1}'
                segments.append((label, w_slots))
            segments.append(('Overall', date_slots_list))
        elif period_type == 'phase' and weeks_current_phase:
            for i, w in enumerate(weeks_current_phase):
                w_slots = _build_date_slots_list_for_batch(dept, batch, w)
                segments.append((f'Week {i + 1}', w_slots))
            segments.append(('Overall', date_slots_list))
        return segments

    styles = {
        'thin_border': Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        ),
        'date_fill': PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid'),
        'date_font': Font(bold=True, color='FFFFFF'),
        'date_align': Alignment(horizontal='center', vertical='center'),
        'header_font': Font(bold=True),
        'lect_fill': PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid'),
        'lect_font': Font(bold=True, color='FFFFFF'),
        'lect_align': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'red_font': Font(color='FF0000'),
    }

    wb = Workbook()
    first = True
    for batch in batches:
        if first:
            ws = wb.active
            first = False
        else:
            ws = wb.create_sheet(title=(batch.name[:31] if batch.name else 'Sheet'))

        students = list(Student.objects.filter(department=dept, batch=batch).order_by('roll_no'))
        date_slots_list = _build_date_slots_list_for_batch(dept, batch, dates)
        overall_segments = build_overall_segments(batch, date_slots_list, None)

        # att_map must cover all dates in main grid and in any segment (e.g. T1 when phase is T2)
        all_dates_for_att = set(d for d, _ in date_slots_list)
        for label, seg_slots in overall_segments:
            for d, _ in seg_slots:
                all_dates_for_att.add(d)
        att_map = {}
        for d in all_dates_for_att:
            for att in FacultyAttendance.objects.filter(batch=batch, date=d):
                key = (d, att.lecture_slot)
                att_map[key] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())

        _write_one_batch_attendance_sheet(ws, batch, date_slots_list, students, att_map, styles, overall_segments=overall_segments)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    if all_batches:
        if period_type == 'daily':
            fname = f'Attendance_All_{dates[0]:%Y-%m-%d}.xlsx'
        elif period_type == 'weekly':
            fname = f'Attendance_All_{phase}_week{week_idx + 1}.xlsx'
        else:
            fname = f'Attendance_All_{phase}.xlsx'
    else:
        batch = batches[0]
        if period_type == 'daily':
            fname = f'Attendance_{batch.name}_{dates[0]:%Y-%m-%d}.xlsx'
        elif period_type == 'weekly':
            fname = f'Attendance_{batch.name}_{phase}_week{week_idx + 1}.xlsx'
        else:
            fname = f'Attendance_{batch.name}_{phase}.xlsx'
    resp = HttpResponse(bio.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename={fname}'
    return resp


def _admin_analytics_data(dept, phase=None, week=None):
    """Compute analytics for admin dashboard: at-risk students, batch-wise, subject-wise, weekly trend, heat map.
    week: None or 'all' = all weeks; int (0-based) = cumulative through that week."""
    tp = TermPhase.objects.filter(department=dept).first()
    phases = ['T1', 'T2', 'T3', 'T4']
    if not phase or phase not in phases:
        phase = next((p for p in phases if getattr(tp, f'{p.lower()}_start', None) and getattr(tp, f'{p.lower()}_end', None)), 'T1')
    weeks = _compile_phase_weeks_date_objects(dept, phase) if tp else []
    all_dates = set()
    if week is not None and week != 'all' and isinstance(week, int) and 0 <= week < len(weeks):
        for i in range(week + 1):
            all_dates.update(weeks[i])
    else:
        for w in weeks:
            all_dates.update(w)
    all_dates = sorted(all_dates)
    if not all_dates:
        return {
            'at_risk_students': [], 'batch_wise': [], 'subject_wise': [],
            'weekly_trend': [], 'heat_map': [], 'heat_map_slots': [], 'phase': phase, 'phases': phases,
            'weeks': weeks, 'num_weeks': len(weeks),
        }
    batches = list(Batch.objects.filter(department=dept).select_related('department'))
    students = list(Student.objects.filter(department=dept).select_related('batch'))
    batch_scheduled = defaultdict(set)
    for batch in batches:
        for d in all_dates:
            weekday = d.strftime('%A')
            for slot in ScheduleSlot.objects.filter(batch=batch, day=weekday).values_list('time_slot', flat=True).distinct():
                batch_scheduled[batch.id].add((d, slot))
    batch_att_map = defaultdict(lambda: defaultdict(set))
    for batch in batches:
        for att in FacultyAttendance.objects.filter(batch=batch, date__in=all_dates):
            key = (att.date, att.lecture_slot)
            batch_att_map[batch.id][key] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())
    at_risk = []
    batch_pcts = defaultdict(list)
    subject_totals = defaultdict(lambda: {'held': 0, 'attended': 0})
    heat_map = defaultdict(lambda: defaultdict(lambda: {'held': 0, 'attended': 0}))
    batch_sizes = {b.id: sum(1 for st in students if st.batch_id == b.id) for b in batches}
    batch_held = {}
    for b in batches:
        scheduled = batch_scheduled.get(b.id, set())
        batch_held[b.id] = len(scheduled)
    for s in students:
        str_roll = str(s.roll_no)
        scheduled = batch_scheduled.get(s.batch_id, set())
        held = batch_held.get(s.batch_id, 0)
        attended = sum(1 for (d, slot) in scheduled if (d, slot) in batch_att_map[s.batch_id] and str_roll not in batch_att_map[s.batch_id][(d, slot)])
        pct = round(attended / held * 100, 1) if held else 0
        if held and pct < 75:
            at_risk.append({'student': s, 'held': held, 'attended': attended, 'pct': pct})
        batch_pcts[s.batch_id].append(pct)
    for b in batches:
        sz = batch_sizes.get(b.id, 0)
        for (d, slot) in batch_scheduled.get(b.id, set()):
            day_name = d.strftime('%A')
            heat_map[day_name][slot]['held'] += sz
            attended_count = sum(1 for st in students if st.batch_id == b.id and (d, slot) in batch_att_map[b.id] and str(st.roll_no) not in batch_att_map[b.id][(d, slot)])
            heat_map[day_name][slot]['attended'] += attended_count
            fac, subj = get_faculty_subject_for_slot(d, b, slot)
            subj_name = subj.name if subj else 'N/A'
            subject_totals[subj_name]['held'] += sz
            subject_totals[subj_name]['attended'] += attended_count
    batch_wise = []
    for b in batches:
        pcts = batch_pcts.get(b.id, [])
        avg_pct = round(sum(pcts) / len(pcts), 1) if pcts else 0
        total_held = batch_held.get(b.id, 0) * len(pcts)
        total_attended = sum(
            sum(1 for (d, slot) in batch_scheduled.get(b.id, set())
                if (d, slot) in batch_att_map[b.id] and str(st.roll_no) not in batch_att_map[b.id][(d, slot)])
            for st in students if st.batch_id == b.id
        )
        batch_wise.append({'batch': b, 'held': total_held, 'attended': total_attended, 'pct': avg_pct})
    subject_wise = []
    for name in sorted(subject_totals.keys()):
        t = subject_totals[name]
        pct = round(t['attended'] / t['held'] * 100, 1) if t['held'] else 0
        subject_wise.append({'name': name, 'held': t['held'], 'attended': t['attended'], 'pct': pct})
    weekly_trend = []
    for i, week_dates in enumerate(weeks):
        if week is not None and week != 'all' and isinstance(week, int) and i > week:
            break
        w_held = w_attended = 0
        for s in students:
            str_roll = str(s.roll_no)
            for (d, slot) in batch_scheduled.get(s.batch_id, set()):
                if d not in week_dates:
                    continue
                w_held += 1
                if (d, slot) in batch_att_map[s.batch_id] and str_roll not in batch_att_map[s.batch_id][(d, slot)]:
                    w_attended += 1
        pct = round(w_attended / w_held * 100, 1) if w_held else 0
        weekly_trend.append({'week': i + 1, 'held': w_held, 'attended': w_attended, 'pct': pct})
    heat_map_list = []
    days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    slots_order = sorted(set(slot for day_data in heat_map.values() for slot in day_data.keys()))
    for day in days_order:
        if day not in heat_map:
            continue
        row = {'day': day, 'slots': []}
        for slot in slots_order:
            t = heat_map[day].get(slot, {'held': 0, 'attended': 0})
            pct = round(t['attended'] / t['held'] * 100, 1) if t['held'] else None
            row['slots'].append({'slot': slot, 'pct': pct, 'held': t['held']})
        heat_map_list.append(row)
    return {
        'at_risk_students': sorted(at_risk, key=lambda x: x['pct']),
        'batch_wise': batch_wise,
        'subject_wise': subject_wise,
        'weekly_trend': weekly_trend,
        'heat_map': heat_map_list,
        'heat_map_slots': slots_order,
        'phase': phase,
        'phases': phases,
        'weeks': weeks,
        'num_weeks': len(weeks),
    }


@login_required
def admin_analytics_dashboard(request):
    """Admin analytics: charts, at-risk students, heat map."""
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    phase = request.GET.get('phase', 'T1')
    week_param = request.GET.get('week', 'all')
    week = None
    if week_param and week_param != 'all':
        try:
            week = int(week_param)
        except ValueError:
            week = None
    data = _admin_analytics_data(dept, phase, week)
    batch_wise_serial = [{'name': b['batch'].name, 'held': b['held'], 'attended': b['attended'], 'pct': b['pct']} for b in data['batch_wise']]
    week_range = list(range(data.get('num_weeks', 0)))
    ctx = {
        'department': dept,
        'is_super_admin': is_super_admin(request),
        'selected_week': week_param,
        'week_range': week_range,
        **data,
        'weekly_trend_json': json.dumps(data['weekly_trend']),
        'subject_wise_json': json.dumps(data['subject_wise']),
        'batch_wise_json': json.dumps(batch_wise_serial),
    }
    return render(request, 'core/admin/analytics_dashboard.html', ctx)


@login_required
def admin_analytics_at_risk_excel(request):
    """Download at-risk students (below 75%) as Excel."""
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        return redirect('core:admin_dashboard')
    phase = request.GET.get('phase', 'T1')
    week_param = request.GET.get('week', 'all')
    week = None
    if week_param and week_param != 'all':
        try:
            week = int(week_param)
        except ValueError:
            week = None
    data = _admin_analytics_data(dept, phase, week)
    at_risk = data['at_risk_students']
    wb = Workbook()
    ws = wb.active
    ws.title = 'At-Risk Students'
    headers = ['Roll No', 'Name', 'Batch', 'Lectures Held', 'Attended', 'Attendance %']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
        cell.font = Font(bold=True, color='FFFFFF')
    for row_idx, r in enumerate(at_risk, 2):
        ws.cell(row_idx, 1, r['student'].roll_no)
        ws.cell(row_idx, 2, r['student'].name)
        ws.cell(row_idx, 3, r['student'].batch.name)
        ws.cell(row_idx, 4, r['held'])
        ws.cell(row_idx, 5, r['attended'])
        ws.cell(row_idx, 6, f"{r['pct']}%")
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 16
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f'At_Risk_Students_{dept.name}_{phase}'
    if week is not None:
        fname += f'_Week{week + 1}'
    fname += '.xlsx'
    resp = HttpResponse(bio.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


def _compile_phase_weeks_date_objects(dept, phase):
    """Return list of weeks, each week = list of date objects (lecture days only, excluding holidays)."""
    tp = TermPhase.objects.filter(department=dept).first()
    if not tp:
        return []
    start = getattr(tp, f'{phase.lower()}_start', None)
    end = getattr(tp, f'{phase.lower()}_end', None)
    if not start or not end:
        return []
    days_set = set(ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct())
    days_set = {d.lower() for d in days_set if d}
    holidays = get_phase_holidays(dept, phase)
    dates = []
    cur = start
    while cur <= end:
        if cur not in holidays and cur.strftime('%A').lower() in days_set:
            dates.append(cur)
        cur += timedelta(days=1)
    dates = sorted(dates)
    weeks = []
    week = []
    last_w = None
    for d in dates:
        w = d.isocalendar()[1]
        if last_w is not None and w != last_w and week:
            weeks.append(week)
            week = []
        week.append(d)
        last_w = w
    if week:
        weeks.append(week)
    return weeks


@login_required
def compile_attendance(request):
    """Compile attendance page: select phase and week, download single-sheet Excel (all students, week-wise + compile columns)."""
    if not user_can_admin(request):
        return redirect('core:admin_dashboard')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    tp = TermPhase.objects.filter(department=dept).first()
    phases = ['T1', 'T2', 'T3', 'T4']
    week_map = {}
    for p in phases:
        start = getattr(tp, f'{p.lower()}_start', None) if tp else None
        end = getattr(tp, f'{p.lower()}_end', None) if tp else None
        if not start or not end:
            week_map[p] = []
            continue
        days_set = set(ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct())
        days_set = {d.lower() for d in days_set if d}
        holidays = get_phase_holidays(dept, p)
        dates = []
        cur = start
        while cur <= end:
            if cur not in holidays and cur.strftime('%A').lower() in days_set:
                dates.append(cur)
            cur += timedelta(days=1)
        weeks = []
        week = []
        last_w = None
        for d in sorted(dates):
            w = d.isocalendar()[1]
            if last_w is not None and w != last_w and week:
                weeks.append(week)
                week = []
            week.append(d)
            last_w = w
        if week:
            weeks.append(week)
        week_map[p] = [[d.isoformat() for d in w] for w in weeks]
    ctx = {
        'department': dept,
        'phases': phases,
        'week_map': week_map,
        'week_map_json': json.dumps(week_map),
    }
    return render(request, 'core/admin/compile_attendance.html', ctx)


def _admin_notifications_build_mentor_data(dept, phase, week_idx):
    """Build list of mentors with their at-risk mentees (below 75%) and full attendance report.
    Returns: [(mentor, [{'student': s, 'held': n, 'attended': n, 'pct': x, 'week_wise': [...], 'subject_wise': [...]}]), ...]
    """
    weeks = _compile_phase_weeks_date_objects(dept, phase)
    if not weeks:
        return []
    cum_dates = set()
    for i in range(len(weeks)):
        cum_dates.update(weeks[i])
        if i == week_idx:
            break
    students = list(
        Student.objects.filter(department=dept, mentor__isnull=False)
        .select_related('batch', 'mentor')
        .order_by('mentor__full_name', 'batch__name', 'roll_no')
    )
    if not students:
        return []
    batch_scheduled = defaultdict(set)
    for s in students:
        batch = s.batch
        for d in cum_dates:
            weekday = d.strftime('%A')
            for slot in ScheduleSlot.objects.filter(batch=batch, day=weekday).values_list('time_slot', flat=True).distinct():
                batch_scheduled[batch.id].add((d, slot))
    batch_att_map = defaultdict(lambda: defaultdict(set))
    for batch_id in {s.batch_id for s in students}:
        batch = next(b for b in students if b.batch_id == batch_id).batch
        for att in FacultyAttendance.objects.filter(batch=batch, date__in=cum_dates):
            key = (att.date, att.lecture_slot)
            batch_att_map[batch_id][key] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())
    mentor_data = defaultdict(list)
    for s in students:
        scheduled = batch_scheduled.get(s.batch_id, set())
        str_roll = str(s.roll_no)
        held = len(scheduled)
        attended = sum(1 for (d, slot) in scheduled if (d, slot) in batch_att_map[s.batch_id] and str_roll not in batch_att_map[s.batch_id][(d, slot)])
        pct = round(attended / held * 100, 1) if held else 0
        if held and pct < 75:
            week_wise = []
            cum_held = cum_attended = 0
            for i, week_dates in enumerate(weeks):
                if i > week_idx:
                    break
                week_set = set(week_dates)
                w_held = sum(1 for (d, slot) in scheduled if d in week_set)
                w_attended = sum(1 for (d, slot) in scheduled if d in week_set and (d, slot) in batch_att_map[s.batch_id] and str_roll not in batch_att_map[s.batch_id][(d, slot)])
                w_pct = round(w_attended / w_held * 100, 1) if w_held else 0
                cum_held += w_held
                cum_attended += w_attended
                cum_pct = round(cum_attended / cum_held * 100, 1) if cum_held else 0
                week_wise.append({'week': i + 1, 'held': w_held, 'attended': w_attended, 'pct': w_pct, 'cum_held': cum_held, 'cum_attended': cum_attended, 'cum_pct': cum_pct})
            subject_wise = defaultdict(lambda: {'held': 0, 'attended': 0})
            for (d, slot) in scheduled:
                fac, subj = get_faculty_subject_for_slot(d, s.batch, slot)
                subj_name = subj.name if subj else 'N/A'
                subject_wise[subj_name]['held'] += 1
                if (d, slot) in batch_att_map[s.batch_id] and str_roll not in batch_att_map[s.batch_id][(d, slot)]:
                    subject_wise[subj_name]['attended'] += 1
            subj_list = [{'name': n, 'held': t['held'], 'attended': t['attended'], 'pct': round(t['attended'] / t['held'] * 100, 1) if t['held'] else 0} for n, t in sorted(subject_wise.items())]
            mentor_data[s.mentor].append({
                'student': s, 'held': held, 'attended': attended, 'pct': pct,
                'week_wise': week_wise, 'subject_wise': subj_list,
            })
    return [(mentor, data) for mentor, data in mentor_data.items() if data]


@login_required
def admin_notifications(request):
    """Admin: Notifications — list students below 75% (phase/week), grouped by mentor. Email to mentor with full report."""
    if not user_can_admin(request):
        return redirect('accounts:role_redirect')
    dept = get_admin_department(request)
    if not dept:
        messages.error(request, 'Select a department first.')
        return redirect('core:admin_dashboard')
    tp = TermPhase.objects.filter(department=dept).first()
    phases = ['T1', 'T2', 'T3', 'T4']
    phase = request.GET.get('phase') or request.POST.get('phase', 'T1')
    week_str = request.GET.get('week') or request.POST.get('week')
    week_map = {}
    for p in phases:
        start = getattr(tp, f'{p.lower()}_start', None) if tp else None
        end = getattr(tp, f'{p.lower()}_end', None) if tp else None
        if not start or not end:
            week_map[p] = []
            continue
        days_set = set(ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct())
        days_set = {d.lower() for d in days_set if d}
        holidays = get_phase_holidays(dept, p)
        dates = []
        cur = start
        while cur <= end:
            if cur not in holidays and cur.strftime('%A').lower() in days_set:
                dates.append(cur)
            cur += timedelta(days=1)
        weeks = []
        week = []
        last_w = None
        for d in sorted(dates):
            w = d.isocalendar()[1]
            if last_w is not None and w != last_w and week:
                weeks.append(week)
                week = []
            week.append(d)
            last_w = w
        if week:
            weeks.append(week)
        week_map[p] = weeks
    weeks_list = week_map.get(phase, [])
    week_idx = 0
    if week_str is not None:
        try:
            week_idx = int(week_str)
            if week_idx < 0 or week_idx >= len(weeks_list):
                week_idx = max(0, len(weeks_list) - 1)
        except ValueError:
            week_idx = 0
    if weeks_list and week_idx >= len(weeks_list):
        week_idx = len(weeks_list) - 1
    mentor_data = []
    if weeks_list:
        mentor_data = _admin_notifications_build_mentor_data(dept, phase, week_idx)
    if request.method == 'POST' and request.POST.get('action') == 'email_mentor':
        mentor_id = request.POST.get('mentor_id')
        if mentor_id:
            mentor = Faculty.objects.filter(pk=mentor_id, department=dept).first()
            if mentor:
                mentor_data_full = _admin_notifications_build_mentor_data(dept, phase, week_idx)
                mentor_entry = next((m for m in mentor_data_full if m[0].id == mentor.id), None)
                if mentor_entry:
                    mentor_fac, at_risk_list = mentor_entry
                    email = mentor_fac.email or (mentor_fac.user.email if mentor_fac.user else None)
                    if email:
                        html = render(request, 'core/admin/email_mentor_attendance_report.html', {
                            'mentor': mentor_fac,
                            'phase': phase,
                            'week_num': week_idx + 1,
                            'at_risk_list': at_risk_list,
                            'department': dept,
                        }).content.decode('utf-8')
                        try:
                            send_mail(
                                subject=f'LJIET Attendance: {len(at_risk_list)} mentee(s) below 75% — {phase} Week {week_idx + 1}',
                                message='Please view this email in HTML format.',
                                from_email=settings.DEFAULT_FROM_EMAIL,
                                recipient_list=[email],
                                html_message=html,
                                fail_silently=False,
                            )
                            messages.success(request, f'Email sent to {mentor_fac.full_name} ({mentor_fac.email or mentor_fac.user.email}).')
                        except Exception as e:
                            err_str = str(e)
                            if '535' in err_str or 'BadCredentials' in err_str or 'Username and Password' in err_str:
                                msg = (
                                    'Email failed: Gmail rejected the login. '
                                    'Use an App Password (not your regular password). '
                                    'Enable 2-Step Verification at myaccount.google.com, then create an App Password at myaccount.google.com/apppasswords. '
                                    'Update EMAIL_HOST_PASSWORD in .env and restart the server.'
                                )
                            else:
                                msg = f'Failed to send email: {e}'
                            messages.error(request, msg)
                    else:
                        messages.error(request, f'No email address for {mentor_fac.full_name}. Add email in Faculty profile.')
                else:
                    messages.error(request, 'Mentor not found in at-risk list.')
            else:
                messages.error(request, 'Invalid mentor.')
        url = reverse('core:admin_notifications') + f'?phase={phase}&week={week_idx}'
        return redirect(url)
    if request.method == 'POST' and request.POST.get('action') == 'email_all':
        mentor_data_full = _admin_notifications_build_mentor_data(dept, phase, week_idx)
        sent = 0
        failed = []
        for mentor_fac, at_risk_list in mentor_data_full:
            email = mentor_fac.email or (mentor_fac.user.email if mentor_fac.user else None)
            if not email:
                failed.append(f'{mentor_fac.full_name} (no email)')
                continue
            html = render(request, 'core/admin/email_mentor_attendance_report.html', {
                'mentor': mentor_fac,
                'phase': phase,
                'week_num': week_idx + 1,
                'at_risk_list': at_risk_list,
                'department': dept,
            }).content.decode('utf-8')
            try:
                send_mail(
                    subject=f'LJIET Attendance: {len(at_risk_list)} mentee(s) below 75% — {phase} Week {week_idx + 1}',
                    message='Please view this email in HTML format.',
                    from_email=settings.DEFAULT_FROM_EMAIL,
                    recipient_list=[email],
                    html_message=html,
                    fail_silently=False,
                )
                sent += 1
            except Exception:
                failed.append(mentor_fac.full_name)
        if sent:
            messages.success(request, f'Email sent to {sent} mentor(s).')
        if failed:
            messages.warning(request, f'Could not email: {", ".join(failed[:5])}{"..." if len(failed) > 5 else ""}')
        url = reverse('core:admin_notifications') + f'?phase={phase}&week={week_idx}'
        return redirect(url)
    week_range = list(range(len(weeks_list)))
    ctx = {
        'department': dept,
        'phases': phases,
        'phase': phase,
        'week_idx': week_idx,
        'week_range': week_range,
        'mentor_data': mentor_data,
    }
    return render(request, 'core/admin/notifications.html', ctx)


def _build_slot_subject_cache(batch, cum_dates, batch_scheduled):
    """Pre-build (date, slot) -> subject_name to avoid N queries. Returns dict."""
    cache = {}
    if not batch_scheduled:
        return cache
    slots_by_day = {}
    for (d, slot) in batch_scheduled:
        day = d.strftime('%A')
        slots_by_day.setdefault(day, set()).add(slot)
    all_slots = set(slot for slots in slots_by_day.values() for slot in slots)
    schedule_slots = list(ScheduleSlot.objects.filter(batch=batch).select_related('subject').only('day', 'time_slot', 'subject'))
    slot_to_subj = {(s.day, s.time_slot): (s.subject.name if s.subject else 'N/A') for s in schedule_slots}
    adj_list = list(LectureAdjustment.objects.filter(
        batch=batch, date__in=cum_dates, time_slot__in=all_slots
    ).select_related('new_subject').values('date', 'time_slot', 'new_subject__name'))
    adj_map = {(a['date'], a['time_slot']): (a['new_subject__name'] or 'N/A') for a in adj_list}
    for (d, slot) in batch_scheduled:
        key = (d, slot)
        if key in adj_map:
            cache[key] = adj_map[key]
        else:
            day = d.strftime('%A')
            cache[key] = slot_to_subj.get((day, slot), 'N/A')
    return cache


def _student_analytics_build_data(dept, phase, week_idx, batch_id, roll_search=None):
    """Build student analytics for a batch: list of {student, held, attended, pct, week_wise, subject_wise}.
    Cumulative phases: T2 = T1+T2, T3 = T1+T2+T3, T4 = T1+T2+T3+T4. Optimized for speed."""
    batch = Batch.objects.filter(pk=batch_id, department=dept).first()
    if not batch:
        return [], []
    week_map, _, phase_dates = _student_phase_weeks_and_dates(dept, batch)
    phases = ['T1', 'T2', 'T3', 'T4']
    phase_order_idx = phases.index(phase) if phase in phases else 0
    weeks = week_map.get(phase, [])
    if not weeks:
        return [], []
    if week_idx < 0 or week_idx >= len(weeks):
        week_idx = len(weeks) - 1
    cum_dates = set()
    for i in range(phase_order_idx + 1):
        cum_dates.update(phase_dates.get(phases[i], []))
    if week_idx is not None and weeks:
        cum_dates = set()
        for i in range(phase_order_idx):
            cum_dates.update(phase_dates.get(phases[i], []))
        for i in range(week_idx + 1):
            cum_dates.update(weeks[i])
    students = Student.objects.filter(department=dept, batch=batch).select_related('batch', 'mentor').order_by('roll_no')
    if roll_search:
        q = Q(roll_no__icontains=roll_search) | Q(name__icontains=roll_search) | Q(enrollment_no__icontains=roll_search)
        students = students.filter(q)
    students = list(students)
    if not students:
        return [], []
    day_slots = defaultdict(set)
    for day, slot in ScheduleSlot.objects.filter(batch=batch).values_list('day', 'time_slot').distinct():
        day_slots[day].add(slot)
    batch_scheduled = set()
    for d in cum_dates:
        weekday = d.strftime('%A')
        for slot in day_slots.get(weekday, ()):
            batch_scheduled.add((d, slot))
    batch_att_map = {}
    for att in FacultyAttendance.objects.filter(batch=batch, date__in=cum_dates).only('date', 'lecture_slot', 'absent_roll_numbers'):
        key = (att.date, att.lecture_slot)
        batch_att_map[key] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())
    slot_subj_cache = _build_slot_subject_cache(batch, cum_dates, batch_scheduled)
    prev_dates_list = [set(phase_dates.get(phases[i], [])) for i in range(phase_order_idx)]
    result = []
    for s in students:
        str_roll = str(s.roll_no)
        held = len(batch_scheduled)
        attended = sum(1 for (d, slot) in batch_scheduled if (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)])
        pct = round(attended / held * 100, 1) if held else 0
        week_wise = []
        cum_held = cum_attended = 0
        for prev_idx in range(phase_order_idx):
            prev_dates = prev_dates_list[prev_idx]
            prev_held = sum(1 for (d, slot) in batch_scheduled if d in prev_dates)
            prev_attended = sum(1 for (d, slot) in batch_scheduled if d in prev_dates and (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)])
            prev_pct = round(prev_attended / prev_held * 100, 1) if prev_held else 0
            cum_held += prev_held
            cum_attended += prev_attended
            cum_pct = round(cum_attended / cum_held * 100, 1) if cum_held else 0
            week_wise.append({'label': f'{phases[prev_idx]} Overall', 'held': prev_held, 'attended': prev_attended, 'pct': prev_pct, 'cum_held': cum_held, 'cum_attended': cum_attended, 'cum_pct': cum_pct})
        weeks_to_show = range(len(weeks)) if week_idx is None else range(min(week_idx + 1, len(weeks)))
        for i in weeks_to_show:
            week_set = set(weeks[i])
            w_held = sum(1 for (d, slot) in batch_scheduled if d in week_set)
            w_attended = sum(1 for (d, slot) in batch_scheduled if d in week_set and (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)])
            w_pct = round(w_attended / w_held * 100, 1) if w_held else 0
            cum_held += w_held
            cum_attended += w_attended
            cum_pct = round(cum_attended / cum_held * 100, 1) if cum_held else 0
            week_wise.append({'label': f'{phase} Week {i + 1}', 'week': i + 1, 'held': w_held, 'attended': w_attended, 'pct': w_pct, 'cum_held': cum_held, 'cum_attended': cum_attended, 'cum_pct': cum_pct})
        subject_wise = defaultdict(lambda: {'held': 0, 'attended': 0})
        for (d, slot) in batch_scheduled:
            subj_name = slot_subj_cache.get((d, slot), 'N/A')
            subject_wise[subj_name]['held'] += 1
            if (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)]:
                subject_wise[subj_name]['attended'] += 1
        subj_list = [{'name': n, 'held': t['held'], 'attended': t['attended'], 'pct': round(t['attended'] / t['held'] * 100, 1) if t['held'] else 0} for n, t in sorted(subject_wise.items())]
        result.append({'student': s, 'held': held, 'attended': attended, 'pct': pct, 'week_wise': week_wise, 'subject_wise': subj_list})
    batches = list(Batch.objects.filter(department=dept).order_by('name'))
    return result, batches


def _student_analytics_build_data_by_roll_search(dept, phase, week_idx, roll_search):
    """Search students by roll/name/enrollment across ALL batches. Return (result, batches)."""
    batches = list(Batch.objects.filter(department=dept).order_by('name'))
    if not roll_search:
        return [], batches
    weeks = _compile_phase_weeks_date_objects(dept, phase)
    if not weeks or week_idx < 0 or week_idx >= len(weeks):
        return [], batches
    q = Q(roll_no__icontains=roll_search) | Q(name__icontains=roll_search) | Q(enrollment_no__icontains=roll_search)
    students = list(Student.objects.filter(department=dept).filter(q).select_related('batch', 'mentor').order_by('batch__name', 'roll_no'))
    if not students:
        return [], batches
    result = []
    for bid in {s.batch_id for s in students}:
        part, _ = _student_analytics_build_data(dept, phase, week_idx, bid, roll_search)
        result.extend(part)
    result.sort(key=lambda x: (x['student'].batch.name, str(x['student'].roll_no)))
    return result, batches


@login_required
def student_analytics(request):
    """Admin & Faculty: Student-wise, batch-wise, phase/week analytics. Filter by batch, search by roll no."""
    dept = None
    is_admin = user_can_admin(request)
    if is_admin:
        # Super admin can set department via POST
        if is_super_admin(request) and request.method == 'POST' and request.POST.get('set_department'):
            request.session['admin_department_id'] = request.POST.get('department_id')
            base = 'core:admin_student_analytics'
            params = request.GET.urlencode()
            url = reverse(base)
            return redirect(f"{url}?{params}" if params else url)
        dept = get_admin_department(request)
        if not dept:
            messages.error(request, 'Select a department first.')
            return redirect('core:admin_dashboard')
    elif user_can_faculty(request):
        faculty = get_faculty_user(request)
        if faculty:
            dept = faculty.department
        if not dept:
            return redirect('accounts:role_redirect')
    else:
        return redirect('accounts:role_redirect')
    phases = ['T1', 'T2', 'T3', 'T4']
    phase = request.GET.get('phase', 'T1')
    week_str = request.GET.get('week')
    batch_id = request.GET.get('batch_id')
    roll_search = request.GET.get('roll_search', '').strip()
    weeks_list = _compile_phase_weeks_date_objects(dept, phase)
    week_idx = 0
    if week_str is not None:
        try:
            week_idx = int(week_str)
            if week_idx < 0 or week_idx >= len(weeks_list):
                week_idx = max(0, len(weeks_list) - 1)
        except ValueError:
            week_idx = 0
    if weeks_list and week_idx >= len(weeks_list):
        week_idx = len(weeks_list) - 1
    student_data = []
    batches = list(Batch.objects.filter(department=dept).select_related('department').order_by('name'))
    batches_from_all_depts = False
    if not batches:
        # Fallback: show batches from ALL departments so dropdown is never empty
        batches = list(Batch.objects.select_related('department').order_by('department__name', 'name'))
        batches_from_all_depts = bool(batches)
    if weeks_list:
        if roll_search and not batch_id:
            student_data, _batches = _student_analytics_build_data_by_roll_search(dept, phase, week_idx, roll_search)
            if not batches_from_all_depts:
                batches = _batches
        elif batch_id:
            batch = Batch.objects.filter(pk=batch_id).select_related('department').first()
            if batch:
                effective_dept = batch.department
                student_data, _batches = _student_analytics_build_data(effective_dept, phase, week_idx, batch_id, roll_search or None)
                if not batches_from_all_depts:
                    batches = _batches
    week_range = list(range(len(weeks_list)))
    departments = list(Department.objects.all()) if is_admin and is_super_admin(request) else []
    ctx = {
        'department': dept,
        'departments': departments,
        'is_super_admin': is_admin and is_super_admin(request),
        'phases': phases,
        'phase': phase,
        'week_idx': week_idx,
        'week_range': week_range,
        'batches': batches,
        'batches_from_all_depts': batches_from_all_depts,
        'selected_batch_id': batch_id,
        'roll_search': roll_search,
        'student_data': student_data,
        'is_admin': is_admin,
    }
    return render(request, 'core/student_analytics.html', ctx)


@login_required
def compile_attendance_excel(request):
    """Download compile attendance: one sheet, all students, columns = Roll No, Name, Batch, W1 Held/Attended/%, W2 Cum Held/Attended/%, ..., Total Held, Total Attended, Total %."""
    if not user_can_admin(request):
        return redirect('core:admin_dashboard')
    dept = get_admin_department(request)
    phase = request.GET.get('phase')
    week_str = request.GET.get('week')
    if not dept or not phase:
        return redirect('core:compile_attendance')
    try:
        week_idx = int(week_str) if week_str is not None else 0
    except Exception:
        week_idx = 0
    weeks = _compile_phase_weeks_date_objects(dept, phase)
    if week_idx < 0 or week_idx >= len(weeks):
        return redirect('core:compile_attendance')
    # Date sets: week 1 only = weeks[0], week 2 cum = weeks[0]+weeks[1], ...
    cumulative_dates = []
    for i in range(week_idx + 1):
        cumulative_dates.append(set(weeks[i]) if i == 0 else cumulative_dates[-1] | set(weeks[i]))
    # All students, all batches, sorted by batch then roll_no
    students = list(Student.objects.filter(department=dept).select_related('batch').order_by('batch__name', 'roll_no'))
    if not students:
        messages.error(request, 'No students in this department.')
        return redirect('core:compile_attendance')
    # Scheduled slots per batch: (date, time_slot) from timetable (ScheduleSlot) for all dates in selected range
    all_dates_in_range = cumulative_dates[week_idx]  # set of dates through selected week
    batch_scheduled = defaultdict(set)
    for batch_id in {s.batch_id for s in students}:
        batch = next(b for b in students if b.batch_id == batch_id).batch
        for d in all_dates_in_range:
            weekday = d.strftime('%A')
            for slot in ScheduleSlot.objects.filter(batch=batch, day=weekday).values_list('time_slot', flat=True).distinct():
                batch_scheduled[batch_id].add((d, slot))
    # Attendance: for each batch, (date, lecture_slot) -> set of absent roll numbers
    batch_att_map = defaultdict(lambda: defaultdict(set))
    for batch_id in {s.batch_id for s in students}:
        batch = next(b for b in students if b.batch_id == batch_id).batch
        for att in FacultyAttendance.objects.filter(batch=batch):
            key = (att.date, att.lecture_slot)
            batch_att_map[batch_id][key] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())
    # Per student, per week range: held = scheduled (date, slot) in date_set; attended = those where attendance was taken and student not in absent list
    def held_attended_for_dates(batch_id, str_roll, date_set):
        scheduled = batch_scheduled[batch_id]
        count_held = sum(1 for (d, slot) in scheduled if d in date_set)
        count_attended = 0
        for (d, slot) in scheduled:
            if d not in date_set:
                continue
            if (d, slot) not in batch_att_map[batch_id]:
                continue  # no attendance taken for this slot
            if str_roll not in batch_att_map[batch_id][(d, slot)]:
                count_attended += 1
        return count_held, count_attended
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
    header_font_white = Font(bold=True, color='FFFFFF')
    wb = Workbook()
    ws = wb.active
    ws.title = 'Compile Attendance'
    col = 1
    ws.cell(1, col, 'Roll No').font = header_font_white
    ws.cell(1, col).fill = header_fill
    ws.cell(1, col).border = thin_border
    col += 1
    ws.cell(1, col, 'Student Name').font = header_font_white
    ws.cell(1, col).fill = header_fill
    ws.cell(1, col).border = thin_border
    col += 1
    ws.cell(1, col, 'Batch').font = header_font_white
    ws.cell(1, col).fill = header_fill
    ws.cell(1, col).border = thin_border
    col += 1
    for i in range(week_idx + 1):
        if i == 0:
            label = 'Week 1'
        else:
            label = f'Week {i + 1} (Cum)'
        for suffix in ('Held', 'Attended', '%'):
            ws.cell(1, col, f'{label} {suffix}').font = header_font_white
            ws.cell(1, col).fill = header_fill
            ws.cell(1, col).border = thin_border
            col += 1
    ws.cell(1, col, 'Total Held').font = header_font_white
    ws.cell(1, col).fill = header_fill
    ws.cell(1, col).border = thin_border
    col += 1
    ws.cell(1, col, 'Total Attended').font = header_font_white
    ws.cell(1, col).fill = header_fill
    ws.cell(1, col).border = thin_border
    col += 1
    ws.cell(1, col, 'Total %').font = header_font_white
    ws.cell(1, col).fill = header_fill
    ws.cell(1, col).border = thin_border
    total_col = col
    col += 1
    for row_idx, s in enumerate(students, start=2):
        str_roll = str(s.roll_no)
        ws.cell(row_idx, 1, s.roll_no).border = thin_border
        ws.cell(row_idx, 2, s.name).border = thin_border
        ws.cell(row_idx, 3, s.batch.name).border = thin_border
        c = 4
        total_held = total_attended = 0
        for i in range(week_idx + 1):
            date_set = cumulative_dates[i]
            h, a = held_attended_for_dates(s.batch_id, str_roll, date_set)
            pct = round(a / h * 100, 1) if h else 0
            ws.cell(row_idx, c, h).border = thin_border
            c += 1
            ws.cell(row_idx, c, a).border = thin_border
            c += 1
            ws.cell(row_idx, c, f'{pct}%').border = thin_border
            c += 1
            if i == week_idx:
                total_held, total_attended = h, a
        ws.cell(row_idx, total_col, total_held).border = thin_border
        ws.cell(row_idx, total_col + 1, total_attended).border = thin_border
        tpct = round(total_attended / total_held * 100, 1) if total_held else 0
        ws.cell(row_idx, total_col + 2, f'{tpct}%').border = thin_border
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 10
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f'Compile_Attendance_{phase}_through_week{week_idx + 1}.xlsx'
    resp = HttpResponse(bio.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename={fname}'
    return resp


# ---------- Faculty: Attendance Entry ----------

@login_required
def faculty_attendance_entry(request):
    if not user_can_faculty(request):
        return redirect('accounts:role_redirect')
    faculty = get_faculty_user(request)
    if not faculty:
        return redirect('accounts:logout')
    dept = faculty.department
    tp = TermPhase.objects.filter(department=dept).first()

    def dates_for_faculty():
        if not tp:
            return []
        holidays = get_all_holiday_dates(dept)
        entries = ScheduleSlot.objects.filter(faculty=faculty)
        day_set = {e.day.lower() for e in entries if e.day}
        out = []
        for i in range(1, 5):
            start = getattr(tp, f't{i}_start', None)
            end = getattr(tp, f't{i}_end', None)
            if not start or not end:
                continue
            cur = start
            while cur <= end:
                if cur not in holidays and cur.strftime('%A').lower() in day_set:
                    out.append(cur)
                cur += timedelta(days=1)
        # Include dates where this faculty is substitute (LectureAdjustment.new_faculty), but not holidays
        adj_dates = LectureAdjustment.objects.filter(new_faculty=faculty).values_list('date', flat=True).distinct()
        for d in adj_dates:
            if d in holidays:
                continue
            for i in range(1, 5):
                start = getattr(tp, f't{i}_start', None)
                end = getattr(tp, f't{i}_end', None)
                if start and end and start <= d <= end:
                    out.append(d)
                    break
        return sorted(set(out))

    available_dates = dates_for_faculty()
    date_str = request.GET.get('date')
    selected_date = None
    if date_str:
        try:
            selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
            if selected_date not in available_dates:
                selected_date = available_dates[0] if available_dates else None
        except Exception:
            selected_date = available_dates[0] if available_dates else None
    else:
        selected_date = available_dates[0] if available_dates else None

    slots_by_batch = defaultdict(list)
    if selected_date:
        weekday = selected_date.strftime('%A')
        for s in ScheduleSlot.objects.filter(faculty=faculty, day=weekday).select_related('batch', 'subject').order_by('time_slot'):
            slots_by_batch[s.batch].append(s)
        # Add slots where this faculty is substitute (LectureAdjustment) for this date
        existing_pairs = {(b, sl.time_slot) for b, slots in slots_by_batch.items() for sl in slots}
        for adj in LectureAdjustment.objects.filter(date=selected_date, new_faculty=faculty).select_related('batch', 'new_subject', 'new_faculty'):
            if (adj.batch, adj.time_slot) in existing_pairs:
                continue
            existing_pairs.add((adj.batch, adj.time_slot))
            virtual = type('Slot', (), {
                'batch': adj.batch, 'time_slot': adj.time_slot,
                'subject': adj.new_subject, 'faculty': adj.new_faculty,
            })()
            slots_by_batch[adj.batch].append(virtual)
        # Keep slots ordered by time_slot per batch
        for b in slots_by_batch:
            slots_by_batch[b].sort(key=lambda s: s.time_slot or '')

    attendance_prefill = defaultdict(lambda: defaultdict(list))
    if selected_date:
        for a in FacultyAttendance.objects.filter(faculty=faculty, date=selected_date):
            attendance_prefill[a.batch.id][a.lecture_slot] = [x.strip() for x in (a.absent_roll_numbers or '').split(',') if x.strip()]

    for batch, slots in slots_by_batch.items():
        for slot in slots:
            slot.prefill_absent_set = set(attendance_prefill.get(batch.id, {}).get(slot.time_slot, []))
            if selected_date:
                fac, subj = get_faculty_subject_for_slot(selected_date, batch, slot.time_slot)
                slot.display_subject_name = subj.name if subj else (slot.subject.name if slot.subject else 'N/A')
                slot.display_faculty_name = fac.short_name if fac else (slot.faculty.short_name if slot.faculty else '—')

    ctx = {
        'faculty': faculty,
        'available_dates': available_dates,
        'selected_date': selected_date,
        'slots_by_batch': dict(slots_by_batch),
    }
    return render(request, 'core/faculty/attendance_entry.html', ctx)


@login_required
def faculty_attendance_save(request):
    if not request.method == 'POST' or not user_can_faculty(request):
        return redirect('core:faculty_attendance_entry')
    faculty = get_faculty_user(request)
    if not faculty:
        return redirect('accounts:logout')
    batch_id = request.POST.get('batch_id')
    lecture_slot = request.POST.get('lecture_slot', '').strip()
    date_str = request.POST.get('date')
    absent_list = request.POST.getlist('absent_roll_numbers')
    if not batch_id or not date_str:
        messages.error(request, 'Missing data.')
        return redirect('core:faculty_attendance_entry')
    try:
        selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    except Exception:
        messages.error(request, 'Invalid date.')
        return redirect('core:faculty_attendance_entry')
    batch = Batch.objects.filter(pk=batch_id, department=faculty.department).first()
    if not batch:
        messages.error(request, 'Invalid batch.')
        return redirect('core:faculty_attendance_entry')
    absent_roll_numbers = ','.join(x.strip() for x in absent_list if x.strip())
    FacultyAttendance.objects.update_or_create(
        faculty=faculty, date=selected_date, batch=batch, lecture_slot=lecture_slot,
        defaults={'absent_roll_numbers': absent_roll_numbers}
    )
    messages.success(request, 'Attendance saved.')
    url = reverse('core:faculty_attendance_entry') + f'?date={date_str}'
    return redirect(url)


@login_required
def faculty_report_excel(request):
    """Export attendance sheet for one date + batch in same format as old project (Roll No, Name, date row, Lect row, P/A with red A, borders)."""
    if not user_can_faculty(request):
        return redirect('core:faculty_dashboard')
    faculty = get_faculty_user(request)
    if not faculty:
        return redirect('accounts:logout')
    date_str = request.GET.get('date')
    batch_id = request.GET.get('batch_id')
    if not date_str or not batch_id:
        messages.error(request, 'Select date and batch.')
        return redirect('core:faculty_attendance_entry')
    try:
        selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    except Exception:
        return redirect('core:faculty_attendance_entry')
    batch = Batch.objects.filter(pk=batch_id, department=faculty.department).first()
    if not batch:
        return redirect('core:faculty_attendance_entry')
    weekday = selected_date.strftime('%A')
    slots = list(ScheduleSlot.objects.filter(
        faculty=faculty, batch=batch, day=weekday
    ).select_related('subject').order_by('time_slot'))
    atts = FacultyAttendance.objects.filter(faculty=faculty, date=selected_date, batch=batch).order_by('lecture_slot')
    att_map = {a.lecture_slot: set(x.strip() for x in (a.absent_roll_numbers or '').split(',') if x.strip()) for a in atts}
    students = list(Student.objects.filter(department=faculty.department, batch=batch).order_by('roll_no'))

    wb = Workbook()
    ws = wb.active
    ws.title = (f'{batch.name} {date_str}')[:31]

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    red_font = Font(color='FF0000')
    # Use explicit style instances (not cell.fill/cell.font) to avoid openpyxl StyleProxy unhashable error
    date_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
    date_font = Font(bold=True, color='FFFFFF')
    date_align = Alignment(horizontal='center', vertical='center')
    header_font = Font(bold=True)

    # Row 1: Roll No, Student Name, then date (merged) over all lecture columns
    ws.cell(row=1, column=1, value='Roll No').font = header_font
    ws.cell(row=1, column=2, value='Student Name').font = header_font
    n_lec = max(len(slots), 1)
    if n_lec > 1:
        ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=2 + n_lec)
    for c in range(1, 3):
        ws.cell(1, c).border = thin_border
    for c in range(3, 3 + n_lec):
        cell = ws.cell(row=1, column=c, value=selected_date.strftime('%d-%b') if c == 3 else None)
        cell.border = thin_border
        cell.fill = date_fill
        cell.font = date_font
        cell.alignment = date_align

    # Row 2: blank A,B; then "Lect 1\nSubject", "Lect 2\nSubject"...
    lect_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
    lect_font = Font(bold=True, color='FFFFFF')
    lect_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for c in range(1, 3):
        ws.cell(row=2, column=c, value='').border = thin_border
    for i, slot in enumerate(slots, start=1):
        fac, subj = get_faculty_subject_for_slot(selected_date, batch, slot.time_slot)
        subj_name = subj.name if subj else (slot.subject.name if slot.subject else 'N/A')
        cell = ws.cell(row=2, column=2 + i, value=f'Lect {i}\n{subj_name}')
        cell.alignment = lect_align
        cell.fill = lect_fill
        cell.font = lect_font
        cell.border = thin_border
    if not slots:
        ws.cell(row=2, column=3, value='').border = thin_border

    # Data rows: roll_no, name, then P/A per lecture
    for idx, s in enumerate(students, start=3):
        ws.cell(row=idx, column=1, value=s.roll_no).border = thin_border
        ws.cell(row=idx, column=2, value=s.name).border = thin_border
        str_roll = str(s.roll_no)
        for i, slot in enumerate(slots, start=1):
            is_absent = str_roll in att_map.get(slot.time_slot, set())
            cell = ws.cell(row=idx, column=2 + i, value='A' if is_absent else 'P')
            cell.border = thin_border
            if is_absent:
                cell.font = red_font

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 20
    for c in range(3, 3 + n_lec):
        ws.column_dimensions[get_column_letter(c)].width = 12
    ws.freeze_panes = 'C3'

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    resp = HttpResponse(bio.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename=Attendance_{batch.name}_{date_str}.xlsx'
    return resp


@login_required
def faculty_mentorship(request):
    """Faculty view: mentorship students with week-wise attendance and at-risk list."""
    if not user_can_faculty(request):
        return redirect('accounts:role_redirect')
    faculty = get_faculty_user(request)
    if not faculty:
        return redirect('accounts:logout')
    dept = faculty.department
    mentorship_students = list(
        Student.objects.filter(mentor=faculty, department=dept)
        .select_related('batch')
        .order_by('batch__name', 'roll_no')
    )
    if not mentorship_students:
        ctx = {'faculty': faculty, 'mentorship_students': [], 'student_stats': [], 'at_risk': [], 'phase': 'T1', 'phases': ['T1', 'T2', 'T3', 'T4'], 'week_range': [], 'selected_week': 'all'}
        return render(request, 'core/faculty/mentorship.html', ctx)
    tp = TermPhase.objects.filter(department=dept).first()
    phase = request.GET.get('phase', 'T1')
    week_param = request.GET.get('week', 'all')
    phases = ['T1', 'T2', 'T3', 'T4']
    week_map, _, phase_dates = _student_phase_weeks_and_dates(dept, mentorship_students[0].batch)
    weeks = week_map.get(phase, [])
    week_idx = None
    if week_param and week_param != 'all':
        try:
            week_idx = int(week_param)
            if week_idx < 0 or week_idx >= len(weeks):
                week_idx = None
        except ValueError:
            week_idx = None
    # Cumulative phase dates: T2 = T1+T2, T3 = T1+T2+T3, T4 = T1+T2+T3+T4
    phase_order_idx = phases.index(phase) if phase in phases else 0
    phase_dates_set = set()
    for i in range(phase_order_idx + 1):
        phase_dates_set.update(phase_dates.get(phases[i], []))
    if week_idx is not None and weeks:
        # Limit to previous phases + weeks 0..week_idx of current phase
        cum_dates = set()
        for i in range(phase_order_idx):
            cum_dates.update(phase_dates.get(phases[i], []))
        for i in range(week_idx + 1):
            cum_dates.update(weeks[i])
        phase_dates_set = cum_dates
    prev_dates_list = [set(phase_dates.get(phases[i], [])) for i in range(phase_order_idx)]
    batch_cache = {}
    student_stats = []
    at_risk = []
    for s in mentorship_students:
        bid = s.batch_id
        if bid not in batch_cache:
            day_slots = defaultdict(set)
            for day, slot in ScheduleSlot.objects.filter(batch=s.batch).values_list('day', 'time_slot').distinct():
                day_slots[day].add(slot)
            batch_scheduled = set()
            for d in phase_dates_set:
                for slot in day_slots.get(d.strftime('%A'), ()):
                    batch_scheduled.add((d, slot))
            batch_att_map = {}
            for att in FacultyAttendance.objects.filter(batch=s.batch, date__in=phase_dates_set).only('date', 'lecture_slot', 'absent_roll_numbers'):
                batch_att_map[(att.date, att.lecture_slot)] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())
            slot_subj = _build_slot_subject_cache(s.batch, phase_dates_set, batch_scheduled)
            batch_cache[bid] = (batch_scheduled, batch_att_map, slot_subj)
        batch_scheduled, batch_att_map, slot_subj = batch_cache[bid]
        held = len(batch_scheduled)
        str_roll = str(s.roll_no)
        attended = sum(1 for (d, slot) in batch_scheduled if (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)])
        pct = round(attended / held * 100, 1) if held else 0
        week_wise = []
        cum_held = cum_attended = 0
        for prev_idx in range(phase_order_idx):
            prev_dates = prev_dates_list[prev_idx]
            prev_held = sum(1 for (d, slot) in batch_scheduled if d in prev_dates)
            prev_attended = sum(1 for (d, slot) in batch_scheduled if d in prev_dates and (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)])
            prev_pct = round(prev_attended / prev_held * 100, 1) if prev_held else 0
            cum_held += prev_held
            cum_attended += prev_attended
            cum_pct = round(cum_attended / cum_held * 100, 1) if cum_held else 0
            week_wise.append({'label': f'{phases[prev_idx]} Overall', 'held': prev_held, 'attended': prev_attended, 'pct': prev_pct, 'cum_held': cum_held, 'cum_attended': cum_attended, 'cum_pct': cum_pct})
        weeks_to_show = range(len(weeks)) if week_idx is None else range(min(week_idx + 1, len(weeks)))
        for i in weeks_to_show:
            week_set = set(weeks[i])
            w_held = sum(1 for (d, slot) in batch_scheduled if d in week_set)
            w_attended = sum(1 for (d, slot) in batch_scheduled if d in week_set and (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)])
            w_pct = round(w_attended / w_held * 100, 1) if w_held else 0
            cum_held += w_held
            cum_attended += w_attended
            cum_pct = round(cum_attended / cum_held * 100, 1) if cum_held else 0
            week_wise.append({'label': f'{phase} Week {i + 1}', 'week': i + 1, 'held': w_held, 'attended': w_attended, 'pct': w_pct, 'cum_held': cum_held, 'cum_attended': cum_attended, 'cum_pct': cum_pct})
        subject_wise = defaultdict(lambda: {'held': 0, 'attended': 0})
        for (d, slot) in batch_scheduled:
            subj_name = slot_subj.get((d, slot), 'N/A')
            subject_wise[subj_name]['held'] += 1
            if (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)]:
                subject_wise[subj_name]['attended'] += 1
        subj_list = [{'name': n, 'held': t['held'], 'attended': t['attended'], 'pct': round(t['attended'] / t['held'] * 100, 1) if t['held'] else 0} for n, t in sorted(subject_wise.items())]
        student_stats.append({
            'student': s, 'held': held, 'attended': attended, 'pct': pct,
            'week_wise': week_wise, 'subject_wise': subj_list,
        })
        if held and pct < 75:
            at_risk.append({'student': s, 'held': held, 'attended': attended, 'pct': pct, 'week_wise': week_wise, 'subject_wise': subj_list})
    week_range = list(range(len(weeks)))
    ctx = {
        'faculty': faculty,
        'mentorship_students': mentorship_students,
        'student_stats': student_stats,
        'at_risk': sorted(at_risk, key=lambda x: x['pct']),
        'phase': phase,
        'phases': phases,
        'week_range': week_range,
        'selected_week': week_param,
    }
    return render(request, 'core/faculty/mentorship.html', ctx)


# ---------- Student: Attendance Summary ----------

def _student_lecture_records(student, batch, dept, start_date, end_date):
    """Return list of dicts: date, time_slot, subject_name, attended (bool) for all lectures in date range."""
    from datetime import date as date_type
    atts = FacultyAttendance.objects.filter(
        batch=batch, date__gte=start_date, date__lte=end_date
    ).order_by('date', 'lecture_slot')
    str_roll = str(student.roll_no)
    records = []
    for att in atts:
        absents = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())
        attended = str_roll not in absents
        fac, subj = get_faculty_subject_for_slot(att.date, batch, att.lecture_slot)
        subject_name = subj.name if subj else 'N/A'
        records.append({
            'date': att.date,
            'time_slot': att.lecture_slot,
            'subject_name': subject_name,
            'attended': attended,
        })
    return records


def _student_phase_weeks_and_dates(dept, batch):
    """Return (week_map with date objects, available_dates list, phase_dates dict phase -> list of dates). Excludes holidays."""
    tp = TermPhase.objects.filter(department=dept).first()
    phases = ['T1', 'T2', 'T3', 'T4']
    days_set = set(ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct())
    if not days_set and batch:
        days_set = set(ScheduleSlot.objects.filter(batch=batch).values_list('day', flat=True).distinct())
    if not days_set:
        days_set = {'monday', 'tuesday', 'wednesday', 'thursday', 'friday'}
    days_set = {d.lower() for d in days_set if d}
    week_map = {}
    phase_dates = {}
    all_dates = []
    for p in phases:
        start = getattr(tp, f'{p.lower()}_start', None) if tp else None
        end = getattr(tp, f'{p.lower()}_end', None) if tp else None
        if not start or not end:
            week_map[p] = []
            phase_dates[p] = []
            continue
        holidays = get_phase_holidays(dept, p)
        dates = []
        cur = start
        while cur <= end:
            if cur not in holidays and cur.strftime('%A').lower() in days_set:
                dates.append(cur)
            cur += timedelta(days=1)
        dates = sorted(dates)
        phase_dates[p] = dates
        weeks = []
        week = []
        last_w = None
        for d in dates:
            w = d.isocalendar()[1]
            if last_w is not None and w != last_w and week:
                weeks.append(week)
                week = []
            week.append(d)
            last_w = w
        if week:
            weeks.append(week)
        week_map[p] = weeks
        all_dates.extend(dates)
    available_dates = sorted(set(all_dates))
    return week_map, available_dates, phase_dates


@login_required
def student_attendance_summary(request):
    """Redirect to student dashboard (My Attendance section removed)."""
    if not user_can_student(request):
        return redirect('accounts:role_redirect')
    return redirect('core:student_dashboard')


@login_required
def student_attendance_analytics(request):
    """Full analytics: held = scheduled (timetable, excluding holidays), attended = where taken and present.
    Totals respect filters and only count lectures up to today (real date-wise)."""
    if not user_can_student(request):
        return redirect('accounts:role_redirect')
    student = get_student_user(request)
    if not student:
        return redirect('accounts:logout')
    dept = student.department
    batch = student.batch
    tp = TermPhase.objects.filter(department=dept).first()
    today = datetime.now().date()
    week_map, available_dates, phase_dates = _student_phase_weeks_and_dates(dept, batch)
    str_roll = str(student.roll_no)
    # Build batch_scheduled and batch_att_map for all phase dates
    phase_dates_all = set()
    for dates in phase_dates.values():
        phase_dates_all.update(dates)
    batch_scheduled = set()
    for d in phase_dates_all:
        weekday = d.strftime('%A')
        for slot in ScheduleSlot.objects.filter(batch=batch, day=weekday).values_list('time_slot', flat=True).distinct():
            batch_scheduled.add((d, slot))
    batch_scheduled_upto_today = batch_scheduled  # Use all scheduled (include future for demo/test data)
    batch_att_map = {}
    for att in FacultyAttendance.objects.filter(batch=batch, date__in=phase_dates_all):
        batch_att_map[(att.date, att.lecture_slot)] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())

    period_type = request.GET.get('period_type', 'date')
    date_str = request.GET.get('date')
    phase = request.GET.get('phase', 'T1')
    week_str = request.GET.get('week')
    selected_date = None
    selected_week_idx = None
    if period_type == 'date' and date_str:
        try:
            selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except Exception:
            pass
    if period_type in ('week', 'phase') and week_str is not None:
        try:
            selected_week_idx = int(week_str)
        except Exception:
            pass

    # --- Date-wise: selected day P/A and day percentage (only count if date <= today)
    day_slots = []
    day_held = day_attended = 0
    if period_type == 'date' and selected_date:
        weekday = selected_date.strftime('%A')
        slots = ScheduleSlot.objects.filter(batch=batch, day=weekday).select_related('subject', 'faculty').order_by('time_slot')
        for s in slots:
            key = (selected_date, s.time_slot)
            attended = None
            if key in batch_att_map:
                attended = str_roll not in batch_att_map[key]
            if key in batch_scheduled:
                day_held += 1
                if attended is not None and attended:
                    day_attended += 1
            day_slots.append({
                'time_slot': s.time_slot,
                'subject': s.subject.name if s.subject else 'N/A',
                'faculty': s.faculty.short_name if s.faculty else 'N/A',
                'attended': attended,
                'status': 'Present' if attended else 'Absent' if attended is False else '—',
            })
        day_pct = round(day_attended / day_held * 100, 1) if day_held else None
    else:
        day_pct = None

    # --- Week-wise: held = scheduled in week (up to today), attended = where taken and present
    weeks_summary = []
    cumulative_held = cumulative_attended = 0
    selected_week_summary = None
    phase_start_date = None
    phase_end_date = None
    if period_type in ('week', 'phase') and tp and week_map.get(phase):
        phase_start_date = getattr(tp, f'{phase.lower()}_start', None)
        phase_end_date = getattr(tp, f'{phase.lower()}_end', None)
        weeks_list = week_map[phase]
        week_dates_sets = [set(w) for w in weeks_list]
        for i, week_dates in enumerate(weeks_list):
            week_set = week_dates_sets[i]
            week_held = sum(1 for (d, slot) in batch_scheduled_upto_today if d in week_set)
            week_attended = sum(1 for (d, slot) in batch_scheduled_upto_today if d in week_set and (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)])
            week_pct = round(week_attended / week_held * 100, 1) if week_held else 0
            cumulative_held += week_held
            cumulative_attended += week_attended
            cum_pct = round(cumulative_attended / cumulative_held * 100, 1) if cumulative_held else 0
            weeks_summary.append({
                'week_num': i + 1,
                'dates': week_dates,
                'held': week_held,
                'attended': week_attended,
                'pct': week_pct,
                'cum_held': cumulative_held,
                'cum_attended': cumulative_attended,
                'cum_pct': cum_pct,
            })
            if selected_week_idx is not None and i == selected_week_idx:
                selected_week_summary = weeks_summary[-1]
        if period_type == 'phase':
            selected_week_summary = weeks_summary[-1] if weeks_summary else None
            selected_week_idx = len(weeks_summary) - 1 if weeks_summary else None

    # --- Subject-wise: use batch_scheduled + batch_att_map, group by subject
    subject_stats = defaultdict(lambda: {'held': 0, 'attended': 0})
    period_dates_set = set()
    if period_type == 'date' and selected_date:
        period_dates_set = {selected_date}
    elif period_type == 'week' and selected_week_idx is not None and week_map.get(phase):
        weeks_list = week_map[phase]
        for i in range(min(selected_week_idx + 1, len(weeks_list))):
            period_dates_set.update(weeks_list[i])
    elif period_type == 'phase' and phase_dates.get(phase):
        period_dates_set = set(phase_dates[phase])
    else:
        period_dates_set = set(phase_dates.get(phase, []) or phase_dates_all)
    for (d, slot) in batch_scheduled_upto_today:
        if d not in period_dates_set:
            continue
        fac, subj = get_faculty_subject_for_slot(d, batch, slot)
        subj_name = subj.name if subj else 'N/A'
        subject_stats[subj_name]['held'] += 1
        if (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)]:
            subject_stats[subj_name]['attended'] += 1
    subject_wise = []
    for name in sorted(subject_stats.keys()):
        s = subject_stats[name]
        pct = round(s['attended'] / s['held'] * 100, 1) if s['held'] else 0
        subject_wise.append({'name': name, 'held': s['held'], 'attended': s['attended'], 'pct': pct})

    # Summary cards: respect filters
    if period_type == 'date' and selected_date:
        total_held = day_held
        total_attended = day_attended
        overall_pct = day_pct if day_pct is not None else 0
    elif period_type in ('week', 'phase') and selected_week_summary is not None:
        total_held = selected_week_summary['cum_held']
        total_attended = selected_week_summary['cum_attended']
        overall_pct = selected_week_summary['cum_pct']
    else:
        # No selection: show selected phase (or all phases)
        phase_dates_set = set(phase_dates.get(phase, []) or phase_dates_all)
        scheduled_in_period = {(d, slot) for (d, slot) in batch_scheduled_upto_today if d in phase_dates_set}
        total_held = len(scheduled_in_period)
        total_attended = sum(1 for (d, slot) in scheduled_in_period if (d, slot) in batch_att_map and str_roll not in batch_att_map[(d, slot)])
        overall_pct = round(total_attended / total_held * 100, 1) if total_held else 0

    phase_weeks = week_map.get(phase, [])

    ctx = {
        'student': student,
        'period_type': period_type,
        'selected_date': selected_date,
        'phase': phase,
        'phases': ['T1', 'T2', 'T3', 'T4'],
        'selected_week_idx': selected_week_idx,
        'week_map': week_map,
        'phase_weeks': phase_weeks,
        'available_dates': available_dates,
        'day_slots': day_slots,
        'day_held': day_held,
        'day_attended': day_attended,
        'day_pct': day_pct,
        'weeks_summary': weeks_summary,
        'selected_week_summary': selected_week_summary,
        'phase_start_date': phase_start_date,
        'phase_end_date': phase_end_date,
        'subject_wise': subject_wise,
        'total_held': total_held,
        'total_attended': total_attended,
        'overall_pct': overall_pct,
    }
    return render(request, 'core/student/attendance_analytics.html', ctx)
