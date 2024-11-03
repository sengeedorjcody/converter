from django.urls import path
from . import views

urlpatterns = [
    path('upload/', views.upload_file, name='upload_file'),
    path('name_request/', views.name_request, name='name_request'),
    path('', views.home, name='home'),
]
