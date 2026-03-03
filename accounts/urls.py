from django.urls import path
from . import views

app_name = 'accounts'

urlpatterns = [
    path('', views.role_redirect, name='role_redirect'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('change-password/', views.change_password_view, name='change_password'),
]
