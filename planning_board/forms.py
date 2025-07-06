# forms.py
from django import forms
from django.forms import inlineformset_factory
from .models import (
    PlanningBoard, ProductionLine, TomorrowPlan, NextDayPlan,
    CriticalPartStatus, AFMPlan, SPDPlan, OtherInformation, ExcelUpload
)

class PlanningBoardForm(forms.ModelForm):
    class Meta:
        model = PlanningBoard
        fields = ['title', 'meeting_time', 'today_date', 'tomorrow_date', 'next_day_date']
        widgets = {
            'title': forms.TextInput(attrs={'class': 'form-control'}),
            'meeting_time': forms.TimeInput(attrs={'class': 'form-control', 'type': 'time'}),
            'today_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'tomorrow_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'next_day_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
        }

class ProductionLineForm(forms.ModelForm):
    class Meta:
        model = ProductionLine
        fields = [
            'line_number',
            'a_shift_model', 'a_shift_plan', 'a_shift_actual', 'a_shift_plan_change', 
            'a_shift_time', 'a_shift_remarks',
            'b_shift_model', 'b_shift_plan', 'b_shift_actual', 'b_shift_plan_change', 
            'b_shift_time', 'b_shift_remarks',
            'c_shift_model', 'c_shift_plan', 'c_shift_actual', 'c_shift_plan_change', 
            'c_shift_remarks'
        ]
        widgets = {
            'line_number': forms.TextInput(attrs={'class': 'form-control'}),
            'a_shift_model': forms.TextInput(attrs={'class': 'form-control'}),
            'a_shift_plan': forms.NumberInput(attrs={'class': 'form-control'}),
            'a_shift_actual': forms.NumberInput(attrs={'class': 'form-control'}),
            'a_shift_plan_change': forms.NumberInput(attrs={'class': 'form-control'}),
            'a_shift_time': forms.TimeInput(attrs={'class': 'form-control', 'type': 'time'}),
            'a_shift_remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'b_shift_model': forms.TextInput(attrs={'class': 'form-control'}),
            'b_shift_plan': forms.NumberInput(attrs={'class': 'form-control'}),
            'b_shift_actual': forms.NumberInput(attrs={'class': 'form-control'}),
            'b_shift_plan_change': forms.NumberInput(attrs={'class': 'form-control'}),
            'b_shift_time': forms.TimeInput(attrs={'class': 'form-control', 'type': 'time'}),
            'b_shift_remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'c_shift_model': forms.TextInput(attrs={'class': 'form-control'}),
            'c_shift_plan': forms.NumberInput(attrs={'class': 'form-control'}),
            'c_shift_actual': forms.NumberInput(attrs={'class': 'form-control'}),
            'c_shift_plan_change': forms.NumberInput(attrs={'class': 'form-control'}),
            'c_shift_remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }

class TomorrowPlanForm(forms.ModelForm):
    class Meta:
        model = TomorrowPlan
        fields = ['model', 'a_shift', 'b_shift', 'c_shift', 'remarks']
        widgets = {
            'model': forms.TextInput(attrs={'class': 'form-control'}),
            'a_shift': forms.NumberInput(attrs={'class': 'form-control'}),
            'b_shift': forms.NumberInput(attrs={'class': 'form-control'}),
            'c_shift': forms.NumberInput(attrs={'class': 'form-control'}),
            'remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }

class NextDayPlanForm(forms.ModelForm):
    class Meta:
        model = NextDayPlan
        fields = ['model', 'a_shift', 'b_shift', 'c_shift', 'remarks']
        widgets = {
            'model': forms.TextInput(attrs={'class': 'form-control'}),
            'a_shift': forms.NumberInput(attrs={'class': 'form-control'}),
            'b_shift': forms.NumberInput(attrs={'class': 'form-control'}),
            'c_shift': forms.NumberInput(attrs={'class': 'form-control'}),
            'remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }

class CriticalPartStatusForm(forms.ModelForm):
    class Meta:
        model = CriticalPartStatus
        fields = ['part_name', 'supplier', 'plan_qty', 'receiving_time', 'remarks']
        widgets = {
            'part_name': forms.TextInput(attrs={'class': 'form-control'}),
            'supplier': forms.TextInput(attrs={'class': 'form-control'}),
            'plan_qty': forms.NumberInput(attrs={'class': 'form-control'}),
            'receiving_time': forms.DateTimeInput(attrs={'class': 'form-control', 'type': 'datetime-local'}),
            'remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }

class AFMPlanForm(forms.ModelForm):
    class Meta:
        model = AFMPlan
        fields = ['plan_type', 'part_name', 'part_number', 'plan_qty', 'remarks']
        widgets = {
            'plan_type': forms.Select(attrs={'class': 'form-control'}),
            'part_name': forms.TextInput(attrs={'class': 'form-control'}),
            'part_number': forms.TextInput(attrs={'class': 'form-control'}),
            'plan_qty': forms.NumberInput(attrs={'class': 'form-control'}),
            'remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }

class SPDPlanForm(forms.ModelForm):
    class Meta:
        model = SPDPlan
        fields = ['customer', 'part_name', 'part_number', 'plan_qty', 'remarks']
        widgets = {
            'customer': forms.Select(attrs={'class': 'form-control'}),
            'part_name': forms.TextInput(attrs={'class': 'form-control'}),
            'part_number': forms.TextInput(attrs={'class': 'form-control'}),
            'plan_qty': forms.NumberInput(attrs={'class': 'form-control'}),
            'remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }

class OtherInformationForm(forms.ModelForm):
    class Meta:
        model = OtherInformation
        fields = ['part_name', 'qty', 'target_date', 'remarks']
        widgets = {
            'part_name': forms.TextInput(attrs={'class': 'form-control'}),
            'qty': forms.NumberInput(attrs={'class': 'form-control'}),
            'target_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'remarks': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
        }

class ExcelUploadForm(forms.ModelForm):
    class Meta:
        model = ExcelUpload
        fields = ['file']
        widgets = {
            'file': forms.FileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls'}),
        }
    
    def clean_file(self):
        file = self.cleaned_data.get('file')
        if file:
            if not file.name.endswith(('.xlsx', '.xls')):
                raise forms.ValidationError("Please upload a valid Excel file (.xlsx or .xls)")
            if file.size > 10 * 1024 * 1024:  # 10MB limit
                raise forms.ValidationError("File size must be less than 10MB")
        return file

# Inline formsets for handling multiple related objects
ProductionLineFormSet = inlineformset_factory(
    PlanningBoard, ProductionLine, 
    form=ProductionLineForm, 
    extra=1, 
    can_delete=True
)

TomorrowPlanFormSet = inlineformset_factory(
    PlanningBoard, TomorrowPlan, 
    form=TomorrowPlanForm, 
    extra=1, 
    can_delete=True
)

NextDayPlanFormSet = inlineformset_factory(
    PlanningBoard, NextDayPlan, 
    form=NextDayPlanForm, 
    extra=1, 
    can_delete=True
)

CriticalPartStatusFormSet = inlineformset_factory(
    PlanningBoard, CriticalPartStatus, 
    form=CriticalPartStatusForm, 
    extra=1, 
    can_delete=True
)

AFMPlanFormSet = inlineformset_factory(
    PlanningBoard, AFMPlan, 
    form=AFMPlanForm, 
    extra=1, 
    can_delete=True
)

SPDPlanFormSet = inlineformset_factory(
    PlanningBoard, SPDPlan, 
    form=SPDPlanForm, 
    extra=1, 
    can_delete=True
)

OtherInformationFormSet = inlineformset_factory(
    PlanningBoard, OtherInformation, 
    form=OtherInformationForm, 
    extra=1, 
    can_delete=True
)