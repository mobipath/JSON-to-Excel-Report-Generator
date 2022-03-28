from django.urls import path
from django.urls import re_path

from . import views
from . import api

urlpatterns = [
    # path('', views.index, name='index'),
    path('', views.ExcelView.as_view()),
    re_path(r'^api/excel_export$', api.ExcelExport.as_view(), name='ExcelExportAPI'),
    re_path(r'^api/excel_export2$', api.ExcelExport2.as_view(), name='ExcelExportAPIpost2'),
    re_path(r'^api/excel_export/get$', api.ExcelExport.as_view(), name='ExcelExportAPI2'),
]