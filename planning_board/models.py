# models.py
from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone

class PlanningBoard(models.Model):
    """Main planning board model"""
    title = models.CharField(max_length=200, default="PRODUCTION PLANNING CONTROL DISPLAY BOARD")
    meeting_time = models.TimeField(null=True, blank=True)
    today_date = models.DateField(default=timezone.now)
    tomorrow_date = models.DateField()
    next_day_date = models.DateField()
    created_by = models.ForeignKey(User, on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        ordering = ['-created_at']
    
    def __str__(self):
        return f"Planning Board - {self.today_date}"

class ProductionLine(models.Model):
    """Production lines like CLUTCH ASSY LINE-1, PULLEY ASSY LINE-1, etc."""
    SHIFT_CHOICES = [
        ('A', 'A Shift'),
        ('B', 'B Shift'),
        ('C', 'C Shift'),
    ]
    
    planning_board = models.ForeignKey(PlanningBoard, on_delete=models.CASCADE, related_name='production_lines')
    line_number = models.CharField(max_length=100)  # e.g., "CLUTCH ASSY LINE-1"
    
    # A Shift Data
    a_shift_model = models.CharField(max_length=100, blank=True)
    a_shift_plan = models.IntegerField(null=True, blank=True)
    a_shift_actual = models.IntegerField(null=True, blank=True)
    a_shift_plan_change = models.IntegerField(null=True, blank=True)
    a_shift_time = models.TimeField(null=True, blank=True)
    a_shift_remarks = models.TextField(blank=True)
    
    # B Shift Data
    b_shift_model = models.CharField(max_length=100, blank=True)
    b_shift_plan = models.IntegerField(null=True, blank=True)
    b_shift_actual = models.IntegerField(null=True, blank=True)
    b_shift_plan_change = models.IntegerField(null=True, blank=True)
    b_shift_time = models.TimeField(null=True, blank=True)
    b_shift_remarks = models.TextField(blank=True)
    
    # C Shift Data
    c_shift_model = models.CharField(max_length=100, blank=True)
    c_shift_plan = models.IntegerField(null=True, blank=True)
    c_shift_actual = models.IntegerField(null=True, blank=True)
    c_shift_plan_change = models.IntegerField(null=True, blank=True)
    c_shift_remarks = models.TextField(blank=True)
    
    class Meta:
        ordering = []
    
    def __str__(self):
        return f"{self.line_number} - {self.planning_board.today_date}"

class TomorrowPlan(models.Model):
    """Tomorrow assembly plan"""
    SHIFT_CHOICES = [
        ('A', 'A Shift'),
        ('B', 'B Shift'),
        ('C', 'C Shift'),
    ]
    
    planning_board = models.ForeignKey(PlanningBoard, on_delete=models.CASCADE, related_name='tomorrow_plans')
    model = models.CharField(max_length=100)
    a_shift = models.IntegerField(null=True, blank=True)
    b_shift = models.IntegerField(null=True, blank=True)
    c_shift = models.IntegerField(null=True, blank=True)
    remarks = models.TextField(blank=True)
    
    def __str__(self):
        return f"Tomorrow Plan - {self.model}"

class NextDayPlan(models.Model):
    """Next day assembly plan"""
    planning_board = models.ForeignKey(PlanningBoard, on_delete=models.CASCADE, related_name='next_day_plans')
    model = models.CharField(max_length=100)
    a_shift = models.IntegerField(null=True, blank=True)
    b_shift = models.IntegerField(null=True, blank=True)
    c_shift = models.IntegerField(null=True, blank=True)
    remarks = models.TextField(blank=True)
    
    def __str__(self):
        return f"Next Day Plan - {self.model}"

class CriticalPartStatus(models.Model):
    """Critical part status tracking"""
    planning_board = models.ForeignKey(PlanningBoard, on_delete=models.CASCADE, related_name='critical_parts')
    part_name = models.CharField(max_length=100)
    supplier = models.CharField(max_length=100)
    plan_qty = models.IntegerField()
    receiving_time = models.DateTimeField(null=True, blank=True)
    remarks = models.TextField(blank=True)
    
    def __str__(self):
        return f"{self.part_name} - {self.supplier}"

class AFMPlan(models.Model):
    """AFM Plan tracking"""
    PLAN_TYPE_CHOICES = [
        ('FCIN', 'FCIN (MNS)'),
        ('IU', 'I/U'),
    ]
    
    planning_board = models.ForeignKey(PlanningBoard, on_delete=models.CASCADE, related_name='afm_plans')
    plan_type = models.CharField(max_length=10, choices=PLAN_TYPE_CHOICES)
    part_name = models.CharField(max_length=100)
    part_number = models.CharField(max_length=50, blank=True)
    plan_qty = models.IntegerField()
    remarks = models.TextField(blank=True)
    
    def __str__(self):
        return f"AFM {self.plan_type} - {self.part_name}"

class SPDPlan(models.Model):
    """SPD Plan customer-wise"""
    CUSTOMER_CHOICES = [
        ('MSIL', 'MSIL'),
        ('HMSI', 'HMSI'),
        ('IYM', 'IYM/PIAGGIO'),
        ('HMCL', 'HMCL'),
    ]
    
    planning_board = models.ForeignKey(PlanningBoard, on_delete=models.CASCADE, related_name='spd_plans')
    customer = models.CharField(max_length=20, choices=CUSTOMER_CHOICES)
    part_name = models.CharField(max_length=100)
    part_number = models.CharField(max_length=50, blank=True)
    plan_qty = models.IntegerField()
    remarks = models.TextField(blank=True)
    
    def __str__(self):
        return f"SPD {self.customer} - {self.part_name}"

class OtherInformation(models.Model):
    """Other information section"""
    planning_board = models.ForeignKey(PlanningBoard, on_delete=models.CASCADE, related_name='other_info')
    part_name = models.CharField(max_length=100)
    qty = models.IntegerField()
    target_date = models.DateField()
    remarks = models.TextField(blank=True)
    
    def __str__(self):
        return f"Other Info - {self.part_name}"

class ExcelUpload(models.Model):
    """Track Excel file uploads"""
    file = models.FileField(upload_to='uploads/excel/')
    planning_board = models.ForeignKey(PlanningBoard, on_delete=models.CASCADE, related_name='excel_uploads')
    uploaded_by = models.ForeignKey(User, on_delete=models.CASCADE)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    processed = models.BooleanField(default=False)
    
    def __str__(self):
        return f"Excel Upload - {self.uploaded_at}"