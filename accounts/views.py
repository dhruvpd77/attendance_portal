from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from .models import UserRole


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
    if role_profile.role == 'admin':
        return redirect('core:admin_dashboard')
    if role_profile.role == 'faculty':
        return redirect('core:faculty_dashboard')
    if role_profile.role == 'student':
        return redirect('core:student_dashboard')
    return redirect('core:home')
