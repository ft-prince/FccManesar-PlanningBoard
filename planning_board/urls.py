# urls.py
from django.urls import path
from . import views

app_name = 'planning_board'

urlpatterns = [
    # Dashboard
    path('', views.planning_board_dashboard, name='dashboard'),
    
    # Planning Board CRUD
    path('boards/', views.planning_board_list, name='list'),
    path('boards/create/', views.planning_board_create, name='create'),
    path('boards/<int:pk>/', views.planning_board_detail, name='detail'),
    path('boards/<int:pk>/edit/', views.planning_board_edit, name='edit'),
    path('boards/<int:pk>/delete/', views.planning_board_delete, name='delete'),
    
    # Excel operations
    path('upload/', views.excel_upload, name='excel_upload'),
    path('boards/<int:pk>/export/', views.export_to_excel, name='export_excel'),
    
    # AJAX endpoints
    path('ajax/add-production-line/', views.ajax_add_production_line, name='ajax_add_production_line'),
]

# For main project urls.py, add this:
"""
# In your main project's urls.py file:

from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    path('planning/', include('planning_board.urls')),
    # Add other app URLs here
]

# Serve media files during development
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
"""