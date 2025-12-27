from django.contrib import admin
from django.urls import path, include
from LetterGen import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('LetterGen.urls')),  # подключение маршрутов LetterGen
]
