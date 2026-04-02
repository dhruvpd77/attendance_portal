from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout, get_user_model
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from .models import UserRole

User = get_user_model()


def login_view(request):
    if request.user.is_authenticated:
        return redirect('accounts:role_redirect')
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        user = authenticate(request, username=username, password=password)
        if user is not None:
            login(request, user)
            return redirect('accounts:role_redirect')
        messages.error(request, 'Invalid username or password.')
    return render(request, 'accounts/login.html')


@login_required
def logout_view(request):
    logout(request)
    messages.success(request, 'You have been logged out.')
    return redirect('accounts:login')


@login_required
def role_redirect(request):
    """Redirect to the correct dashboard based on user role."""
    try:
        role_profile = request.user.role_profile
    except UserRole.DoesNotExist:
        # Superuser or staff -> admin dashboard
        if request.user.is_superuser or request.user.is_staff:
            return redirect('core:admin_dashboard')
        messages.warning(request, 'No role assigned. Contact administrator.')
        return redirect('accounts:login')
    if role_profile.role in ('admin', 'hod'):
        return redirect('core:admin_dashboard')
    if role_profile.role == 'exam_admin':
        return redirect('core:exam_admin_dashboard')
    if role_profile.role == 'exam_section':
        return redirect('core:exam_section_dashboard')
    if role_profile.role in ('dept_exam_parent', 'dept_exam_child'):
        return redirect('core:dept_exam_dashboard')
    if role_profile.role == 'faculty':
        return redirect('core:faculty_dashboard')
    if role_profile.role == 'student':
        return redirect('core:student_dashboard')
    return redirect('core:home')


def change_password_view(request):
    """Change password: supports both logged-in and anonymous users (via username)."""
    if request.method != 'POST':
        return render(request, 'accounts/change_password.html', {'show_username': not request.user.is_authenticated})

    # Determine user
    if request.user.is_authenticated:
        user = request.user
        username = user.username
    else:
        username = request.POST.get('username', '').strip()
        if not username:
            messages.error(request, 'Username is required.')
            return render(request, 'accounts/change_password.html', {'show_username': True})
        try:
            user = User.objects.get(username=username)
        except User.DoesNotExist:
            messages.error(request, 'No account found with this username.')
            return render(request, 'accounts/change_password.html', {'show_username': True})

    old_password = request.POST.get('old_password', '')
    new_password = request.POST.get('new_password', '')
    confirm_password = request.POST.get('confirm_password', '')

    if not old_password:
        messages.error(request, 'Please enter your current password.')
        return render(request, 'accounts/change_password.html', {'show_username': not request.user.is_authenticated})

    if not user.check_password(old_password):
        messages.error(request, 'Current password is incorrect.')
        return render(request, 'accounts/change_password.html', {'show_username': not request.user.is_authenticated})

    if not new_password:
        messages.error(request, 'Please enter a new password.')
        return render(request, 'accounts/change_password.html', {'show_username': not request.user.is_authenticated})

    if new_password != confirm_password:
        messages.error(request, 'New password and confirmation do not match.')
        return render(request, 'accounts/change_password.html', {'show_username': not request.user.is_authenticated})

    user.set_password(new_password)
    user.save()
    messages.success(request, 'Your password has been changed successfully. Please sign in with your new password.')
    return redirect('accounts:login')
