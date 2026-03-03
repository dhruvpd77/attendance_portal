from django.contrib import admin
from django.http import JsonResponse
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static


def health_check(request):
    """Health check endpoint for load balancers and monitoring."""
    return JsonResponse({'status': 'ok', 'app': 'LJIET_Attendance'})


urlpatterns = [
    path('health/', health_check),
    path('admin/', admin.site.urls),
    path('portal/', include('core.urls')),
    path('', include('accounts.urls')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

handler404 = 'core.views.handler404'
handler500 = 'core.views.handler500'
