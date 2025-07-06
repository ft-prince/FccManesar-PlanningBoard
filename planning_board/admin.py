# admin.py
from django.contrib import admin
from .models import (
    PlanningBoard, ProductionLine, TomorrowPlan, NextDayPlan,
    CriticalPartStatus, AFMPlan, SPDPlan, OtherInformation, ExcelUpload
)

class ProductionLineInline(admin.TabularInline):
    model = ProductionLine
    extra = 1
    fields = [
        'line_number',
        'a_shift_model', 'a_shift_plan', 'a_shift_actual',
        'b_shift_model', 'b_shift_plan', 'b_shift_actual',
        'c_shift_model', 'c_shift_plan', 'c_shift_actual'
    ]

class TomorrowPlanInline(admin.TabularInline):
    model = TomorrowPlan
    extra = 1

class NextDayPlanInline(admin.TabularInline):
    model = NextDayPlan
    extra = 1

class CriticalPartStatusInline(admin.TabularInline):
    model = CriticalPartStatus
    extra = 1

@admin.register(PlanningBoard)
class PlanningBoardAdmin(admin.ModelAdmin):
    list_display = [
        'title', 'today_date', 'meeting_time', 
        'created_by', 'created_at', 'updated_at'
    ]
    list_filter = ['today_date', 'created_by', 'created_at']
    search_fields = ['title', 'created_by__username']
    readonly_fields = ['created_at', 'updated_at']
    
    inlines = [
        ProductionLineInline,
        TomorrowPlanInline,
        NextDayPlanInline,
        CriticalPartStatusInline,
    ]
    
    fieldsets = (
        ('Basic Information', {
            'fields': ('title', 'meeting_time')
        }),
        ('Dates', {
            'fields': ('today_date', 'tomorrow_date', 'next_day_date')
        }),
        ('Metadata', {
            'fields': ('created_by', 'created_at', 'updated_at'),
            'classes': ('collapse',)
        }),
    )
    
    def save_model(self, request, obj, form, change):
        if not change:  # If creating new object
            obj.created_by = request.user
        super().save_model(request, obj, form, change)

@admin.register(ProductionLine)
class ProductionLineAdmin(admin.ModelAdmin):
    list_display = [
        'line_number', 'planning_board', 
        'a_shift_model', 'a_shift_plan', 'a_shift_actual',
        'b_shift_model', 'b_shift_plan', 'b_shift_actual',
        'c_shift_model', 'c_shift_plan', 'c_shift_actual'
    ]
    list_filter = ['planning_board__today_date', 'line_number']
    search_fields = ['line_number', 'a_shift_model', 'b_shift_model', 'c_shift_model']
    
    fieldsets = (
        ('Basic Information', {
            'fields': ('planning_board', 'line_number')
        }),
        ('A Shift', {
            'fields': (
                'a_shift_model', 'a_shift_plan', 'a_shift_actual', 
                'a_shift_plan_change', 'a_shift_time', 'a_shift_remarks'
            )
        }),
        ('B Shift', {
            'fields': (
                'b_shift_model', 'b_shift_plan', 'b_shift_actual', 
                'b_shift_plan_change', 'b_shift_time', 'b_shift_remarks'
            )
        }),
        ('C Shift', {
            'fields': (
                'c_shift_model', 'c_shift_plan', 'c_shift_actual', 
                'c_shift_plan_change', 'c_shift_remarks'
            )
        }),
    )

@admin.register(TomorrowPlan)
class TomorrowPlanAdmin(admin.ModelAdmin):
    list_display = ['model', 'planning_board', 'a_shift', 'b_shift', 'c_shift']
    list_filter = ['planning_board__tomorrow_date']
    search_fields = ['model']

@admin.register(NextDayPlan)
class NextDayPlanAdmin(admin.ModelAdmin):
    list_display = ['model', 'planning_board', 'a_shift', 'b_shift', 'c_shift']
    list_filter = ['planning_board__next_day_date']
    search_fields = ['model']

@admin.register(CriticalPartStatus)
class CriticalPartStatusAdmin(admin.ModelAdmin):
    list_display = ['part_name', 'supplier', 'plan_qty', 'receiving_time', 'planning_board']
    list_filter = ['supplier', 'planning_board__today_date']
    search_fields = ['part_name', 'supplier']
    date_hierarchy = 'receiving_time'

@admin.register(AFMPlan)
class AFMPlanAdmin(admin.ModelAdmin):
    list_display = ['part_name', 'plan_type', 'part_number', 'plan_qty', 'planning_board']
    list_filter = ['plan_type', 'planning_board__today_date']
    search_fields = ['part_name', 'part_number']

@admin.register(SPDPlan)
class SPDPlanAdmin(admin.ModelAdmin):
    list_display = ['part_name', 'customer', 'part_number', 'plan_qty', 'planning_board']
    list_filter = ['customer', 'planning_board__today_date']
    search_fields = ['part_name', 'part_number']

@admin.register(OtherInformation)
class OtherInformationAdmin(admin.ModelAdmin):
    list_display = ['part_name', 'qty', 'target_date', 'planning_board']
    list_filter = ['target_date', 'planning_board__today_date']
    search_fields = ['part_name']
    date_hierarchy = 'target_date'

@admin.register(ExcelUpload)
class ExcelUploadAdmin(admin.ModelAdmin):
    list_display = ['file', 'planning_board', 'uploaded_by', 'uploaded_at', 'processed']
    list_filter = ['processed', 'uploaded_at', 'uploaded_by']
    readonly_fields = ['uploaded_at']
    
    def has_change_permission(self, request, obj=None):
        # Prevent editing processed uploads
        if obj and obj.processed:
            return False
        return super().has_change_permission(request, obj)