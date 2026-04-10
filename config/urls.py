from django.contrib import admin
from django.urls import path
from django.conf import settings
from django.conf.urls.static import static
from core import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.index, name='index'),
    path('upload/', views.upload, name='upload'),
    path('process/<int:session_id>/', views.process_session, name='process'),
    path('review/<int:session_id>/', views.review, name='review'),
    path('download/<int:session_id>/<str:file_type>/', views.download_file, name='download'),
    path('reupload/<int:session_id>/', views.reupload, name='reupload'),
    path('generate/<int:session_id>/', views.generate_report, name='generate_report'),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
