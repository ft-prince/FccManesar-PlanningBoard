from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from planning_board import views
urlpatterns = [
    path('admin/', admin.site.urls),
    path('planning/', include('planning_board.urls')),
    path('', views.planning_board_dashboard, name='dashboard'),
    # Add other app URLs here
]
