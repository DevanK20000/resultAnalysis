from cgitb import handler
from os import stat
from django.urls import path
from . import views
from django.conf.urls.static import static
from django.conf import settings

urlpatterns = [
    path('login/', views.loginPage, name='login'),
    path('logout/', views.logoutUser, name='logout'),
    path('table/', views.table, name='table'),
    path('download/<str:pk>', views.download, name='download'),
    path('', views.home, name='home'),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

handler404 = "base.views.page_not_found_view"
