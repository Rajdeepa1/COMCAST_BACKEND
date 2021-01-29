from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path(r'api/', include('comcast_run_script.urls'))
]