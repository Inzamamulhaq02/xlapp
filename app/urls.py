# urls.py
from django.urls import path
from .views import *
from django.conf import settings
from django.conf.urls.static import static

if settings.DEBUG:
    urlpatterns = [
    path('', upload_and_process_excel, name='upload_excel'),
]

    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)