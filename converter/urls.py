from django.urls import path
from converter import views

urlpatterns = [
    path("", views.index, name="index"),
    path("convert/", views.convert_files, name="convert"),
    path("load-demo/", views.load_demo, name="load_demo"),
    path("download/<str:file_id>/", views.download_excel, name="download"),
]
