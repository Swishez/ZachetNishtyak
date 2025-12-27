from django.urls import path
from LetterGen import views
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('create/', views.create_letter, name='create_letter'),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
