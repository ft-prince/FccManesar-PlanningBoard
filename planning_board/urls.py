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
    path('<int:pk>/inline-update/', views.inline_update_board, name='inline_update'),

    # AJAX endpoints
    path('ajax/add-production-line/', views.ajax_add_production_line, name='ajax_add_production_line'),
    
    
    
    
    
    
    #  new 
    
    path('live-view/', views.live_view_page, name='live_view'),
    path('api/boards/', views.get_user_planning_boards, name='api_boards'),
    path('api/board/<int:board_id>/sections/', views.get_board_sections_summary, name='api_board_sections'),
    path('api/board/<int:board_id>/section/<str:section>/', views.get_section_data, name='api_section_data'),
    path('api/board/<int:board_id>/section/<str:section>/stream/', views.live_stream_section, name='api_live_stream'),
    path('api/board/<int:board_id>/trigger-update/', views.trigger_board_update, name='api_trigger_update'),


    # Fullscreen Display - NEW ADDITION
    path('display/<int:board_id>/<str:section>/', views.fullscreen_display, name='fullscreen_display'),
    path('api/display/<int:board_id>/<str:section>/stream/', views.fullscreen_data_stream, name='fullscreen_stream'),



#  new update
    path('monitor/display/', views.monitor_display, name='monitor_display'),
    path('monitor/controller/', views.monitor_controller, name='monitor_controller'),
    
    # Monitor API endpoints
    path('api/monitor/control/', views.monitor_control_api, name='monitor_control_api'),
    path('api/monitor/status/', views.monitor_status_api, name='monitor_status_api'),
    path('api/monitor/<int:board_id>/<str:section>/stream/', views.monitor_data_stream, name='monitor_data_stream'),

]

