# views.py - Fixed version with proper Excel processing
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import JsonResponse, HttpResponse
from django.urls import reverse
from django.utils import timezone
from datetime import datetime, timedelta
import openpyxl
import io
from .models import (
    PlanningBoard, ProductionLine, TomorrowPlan, NextDayPlan,
    CriticalPartStatus, AFMPlan, SPDPlan, OtherInformation, ExcelUpload
)
from .forms import (
    PlanningBoardForm, ExcelUploadForm, ProductionLineFormSet,
    TomorrowPlanFormSet, NextDayPlanFormSet, CriticalPartStatusFormSet,
    AFMPlanFormSet, SPDPlanFormSet, OtherInformationFormSet
)

@login_required
def planning_board_list(request):
    """List all planning boards"""
    boards = PlanningBoard.objects.filter(created_by=request.user).order_by('-created_at')
    return render(request, 'planning_board/list.html', {'boards': boards})

@login_required
def planning_board_detail(request, pk):
    """View a specific planning board"""
    board = get_object_or_404(PlanningBoard, pk=pk, created_by=request.user)
    return render(request, 'planning_board/detail.html', {'board': board})

@login_required
def planning_board_create(request):
    """Create a new planning board"""
    if request.method == 'POST':
        form = PlanningBoardForm(request.POST)
        if form.is_valid():
            board = form.save(commit=False)
            board.created_by = request.user
            board.save()
            messages.success(request, 'Planning board created successfully!')
            return redirect('planning_board:edit', pk=board.pk)
    else:
        # Set default dates
        today = timezone.now().date()
        form = PlanningBoardForm(initial={
            'today_date': today,
            'tomorrow_date': today + timedelta(days=1),
            'next_day_date': today + timedelta(days=2),
        })
    
    return render(request, 'planning_board/create.html', {'form': form})

@login_required
def planning_board_edit(request, pk):
    """Edit a planning board with all related data"""
    board = get_object_or_404(PlanningBoard, pk=pk, created_by=request.user)
    
    if request.method == 'POST':
        form = PlanningBoardForm(request.POST, instance=board)
        production_formset = ProductionLineFormSet(request.POST, instance=board)
        tomorrow_formset = TomorrowPlanFormSet(request.POST, instance=board)
        next_day_formset = NextDayPlanFormSet(request.POST, instance=board)
        critical_formset = CriticalPartStatusFormSet(request.POST, instance=board)
        afm_formset = AFMPlanFormSet(request.POST, instance=board)
        spd_formset = SPDPlanFormSet(request.POST, instance=board)
        other_formset = OtherInformationFormSet(request.POST, instance=board)
        
        if (form.is_valid() and production_formset.is_valid() and 
            tomorrow_formset.is_valid() and next_day_formset.is_valid() and
            critical_formset.is_valid() and afm_formset.is_valid() and
            spd_formset.is_valid() and other_formset.is_valid()):
            
            form.save()
            production_formset.save()
            tomorrow_formset.save()
            next_day_formset.save()
            critical_formset.save()
            afm_formset.save()
            spd_formset.save()
            other_formset.save()
            
            messages.success(request, 'Planning board updated successfully!')
            return redirect('planning_board:detail', pk=board.pk)
        else:
            messages.error(request, 'Please correct the errors below.')
    else:
        form = PlanningBoardForm(instance=board)
        production_formset = ProductionLineFormSet(instance=board)
        tomorrow_formset = TomorrowPlanFormSet(instance=board)
        next_day_formset = NextDayPlanFormSet(instance=board)
        critical_formset = CriticalPartStatusFormSet(instance=board)
        afm_formset = AFMPlanFormSet(instance=board)
        spd_formset = SPDPlanFormSet(instance=board)
        other_formset = OtherInformationFormSet(instance=board)
    
    context = {
        'form': form,
        'board': board,
        'production_formset': production_formset,
        'tomorrow_formset': tomorrow_formset,
        'next_day_formset': next_day_formset,
        'critical_formset': critical_formset,
        'afm_formset': afm_formset,
        'spd_formset': spd_formset,
        'other_formset': other_formset,
    }
    
    return render(request, 'planning_board/edit.html', context)

@login_required
def planning_board_delete(request, pk):
    """Delete a planning board"""
    board = get_object_or_404(PlanningBoard, pk=pk, created_by=request.user)
    
    if request.method == 'POST':
        board.delete()
        messages.success(request, 'Planning board deleted successfully!')
        return redirect('planning_board:list')
    
    return render(request, 'planning_board/delete.html', {'board': board})

@login_required
def excel_upload(request):
    """Upload and process Excel file"""
    if request.method == 'POST':
        form = ExcelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            # Create planning board first
            today = timezone.now().date()
            board = PlanningBoard.objects.create(
                created_by=request.user,
                today_date=today,
                tomorrow_date=today + timedelta(days=1),
                next_day_date=today + timedelta(days=2),
            )
            
            # Save the upload record
            upload = form.save(commit=False)
            upload.planning_board = board
            upload.uploaded_by = request.user
            upload.save()
            
            # Process the Excel file
            try:
                success = process_excel_file(upload.file.path, board)
                if success:
                    upload.processed = True
                    upload.save()
                    messages.success(request, f'Excel file uploaded and processed successfully! Planning board created with ID: {board.id}')
                    return redirect('planning_board:detail', pk=board.pk)
                else:
                    board.delete()
                    messages.error(request, 'Error processing Excel file. Please check the file format.')
            except Exception as e:
                board.delete()  # Clean up if processing fails
                messages.error(request, f'Error processing Excel file: {str(e)}')
    else:
        form = ExcelUploadForm()
    
    return render(request, 'planning_board/excel_upload.html', {'form': form})

def process_excel_file(file_path, board):
    """Process uploaded Excel file and populate database"""
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        worksheet = workbook.active
        
        # Extract basic information
        extract_basic_info(worksheet, board)
        
        # Extract production lines data
        extract_production_lines(worksheet, board)
        
        # Extract other data sections
        extract_future_plans(worksheet, board)
        
        return True
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        return False

def extract_basic_info(worksheet, board):
    """Extract basic information from Excel"""
    try:
        # Extract meeting time from B2
        meeting_cell = worksheet['B2'].value
        if meeting_cell and 'TIME' in str(meeting_cell).upper():
            # Try to extract time
            import re
            time_match = re.search(r'(\d{1,2}):(\d{2})', str(meeting_cell))
            if time_match:
                from datetime import time
                hour, minute = int(time_match.group(1)), int(time_match.group(2))
                board.meeting_time = time(hour, minute)
        
        # Update title from C2
        title_cell = worksheet['C2'].value
        if title_cell:
            board.title = str(title_cell)
        
        board.save()
    except Exception as e:
        print(f"Error extracting basic info: {e}")

def extract_production_lines(worksheet, board):
    """Extract production line data from specific rows"""
    try:
        # Define production lines and their row positions based on your Excel structure
        line_configs = [
            {'name': 'CLUTCH ASSY LINE-1', 'row': 7},
            {'name': 'CLUTCH ASSY LINE-2', 'row': 13},
            {'name': 'PULLEY ASSY LINE-1', 'row': 19},
            {'name': 'FMD/FFD', 'row': 22},
            {'name': 'NEW BUSINESS', 'row': 26},
        ]
        
        for config in line_configs:
            try:
                # Check if line exists in Excel
                line_cell_value = worksheet.cell(row=config['row'], column=2).value
                
                # Create production line regardless (we'll populate what we can)
                production_line = ProductionLine.objects.create(
                    planning_board=board,
                    line_number=config['name']
                )
                
                # Try to extract data from multiple rows around the line
                for row_offset in range(6):  # Check up to 6 rows
                    current_row = config['row'] + row_offset
                    
                    # A Shift data (columns C-H approximately)
                    if not production_line.a_shift_model:
                        model_val = get_cell_value(worksheet, current_row, 3)
                        if model_val:
                            production_line.a_shift_model = model_val
                            production_line.a_shift_plan = get_numeric_value(worksheet, current_row, 4)
                            production_line.a_shift_actual = get_numeric_value(worksheet, current_row, 5)
                            production_line.a_shift_plan_change = get_numeric_value(worksheet, current_row, 6)
                            production_line.a_shift_remarks = get_cell_value(worksheet, current_row, 8)
                    
                    # B Shift data (columns I-N approximately)
                    if not production_line.b_shift_model:
                        model_val = get_cell_value(worksheet, current_row, 9)
                        if model_val:
                            production_line.b_shift_model = model_val
                            production_line.b_shift_plan = get_numeric_value(worksheet, current_row, 10)
                            production_line.b_shift_actual = get_numeric_value(worksheet, current_row, 11)
                            production_line.b_shift_plan_change = get_numeric_value(worksheet, current_row, 12)
                            production_line.b_shift_remarks = get_cell_value(worksheet, current_row, 14)
                    
                    # C Shift data (columns O-S approximately)
                    if not production_line.c_shift_model:
                        model_val = get_cell_value(worksheet, current_row, 15)
                        if model_val:
                            production_line.c_shift_model = model_val
                            production_line.c_shift_plan = get_numeric_value(worksheet, current_row, 16)
                            production_line.c_shift_actual = get_numeric_value(worksheet, current_row, 17)
                            production_line.c_shift_plan_change = get_numeric_value(worksheet, current_row, 18)
                            production_line.c_shift_remarks = get_cell_value(worksheet, current_row, 19)
                
                production_line.save()
                
            except Exception as e:
                print(f"Error processing line {config['name']}: {e}")
                continue
                
    except Exception as e:
        print(f"Error extracting production lines: {e}")

def extract_future_plans(worksheet, board):
    """Extract tomorrow and next day plans"""
    try:
        # Extract tomorrow plans (around columns T-W)
        for row in range(6, 20):  # Check multiple rows
            model = get_cell_value(worksheet, row, 20)  # Column T
            if model and model.strip() and model not in ['MODEL', 'SHIFT']:
                TomorrowPlan.objects.create(
                    planning_board=board,
                    model=model,
                    a_shift=get_numeric_value(worksheet, row, 21),
                    b_shift=get_numeric_value(worksheet, row, 22),
                    c_shift=get_numeric_value(worksheet, row, 23),
                    remarks=get_cell_value(worksheet, row, 24)
                )
        
        # Extract next day plans (around columns Y-AB)
        for row in range(6, 20):
            model = get_cell_value(worksheet, row, 25)  # Column Y
            if model and model.strip() and model not in ['MODEL', 'SHIFT']:
                NextDayPlan.objects.create(
                    planning_board=board,
                    model=model,
                    a_shift=get_numeric_value(worksheet, row, 26),
                    b_shift=get_numeric_value(worksheet, row, 27),
                    c_shift=get_numeric_value(worksheet, row, 28),
                    remarks=get_cell_value(worksheet, row, 29)
                )
                
    except Exception as e:
        print(f"Error extracting future plans: {e}")

def get_cell_value(worksheet, row, col):
    """Get cell value as string, handling None values"""
    try:
        cell = worksheet.cell(row=row, column=col)
        value = cell.value
        if value is None:
            return ''
        return str(value).strip()
    except:
        return ''

def get_numeric_value(worksheet, row, col):
    """Get cell value as number, handling None and non-numeric values"""
    try:
        cell = worksheet.cell(row=row, column=col)
        value = cell.value
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return int(value)
        # Try to convert string to number
        try:
            return int(float(str(value)))
        except:
            return None
    except:
        return None

@login_required
def export_to_excel(request, pk):
    """Export planning board data to Excel"""
    board = get_object_or_404(PlanningBoard, pk=pk, created_by=request.user)
    
    # Create new workbook
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Planning Board"
    
    # Set up the Excel structure similar to the original format
    worksheet['B2'] = f"MEETING TIME: {board.meeting_time if board.meeting_time else ''}"
    worksheet['C2'] = board.title
    
    # Dates
    worksheet['C3'] = f"DATE:- {board.today_date}"
    worksheet['T3'] = f"DATE:- {board.tomorrow_date}"
    worksheet['Y3'] = f"DATE:- {board.next_day_date}"
    
    # Section headers
    worksheet['C4'] = "TODAY ASSY PLAN"
    worksheet['T4'] = "TOMORROW ASSY PLAN"
    worksheet['Y4'] = "NEXT DAY ASSY PLAN"
    
    # Column headers for production lines
    headers = ['LINE NO.', 'MODEL', 'PLAN', 'ACTUAL', 'PLAN CHANGE', 'TIME', 'REMARKS']
    for i, header in enumerate(headers, start=2):
        worksheet.cell(row=6, column=i, value=header)
    
    # Production line data
    row = 7
    for line in board.production_lines.all():
        worksheet.cell(row=row, column=2, value=line.line_number)
        # A Shift
        worksheet.cell(row=row, column=3, value=line.a_shift_model)
        worksheet.cell(row=row, column=4, value=line.a_shift_plan)
        worksheet.cell(row=row, column=5, value=line.a_shift_actual)
        worksheet.cell(row=row, column=6, value=line.a_shift_plan_change)
        worksheet.cell(row=row, column=7, value=str(line.a_shift_time) if line.a_shift_time else '')
        worksheet.cell(row=row, column=8, value=line.a_shift_remarks)
        row += 1
    
    # Tomorrow plans
    row = 7
    for plan in board.tomorrow_plans.all():
        worksheet.cell(row=row, column=20, value=plan.model)
        worksheet.cell(row=row, column=21, value=plan.a_shift)
        worksheet.cell(row=row, column=22, value=plan.b_shift)
        worksheet.cell(row=row, column=23, value=plan.c_shift)
        row += 1
    
    # Create HTTP response
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="planning_board_{board.today_date}.xlsx"'
    
    # Save workbook to response
    workbook.save(response)
    return response

@login_required
def ajax_add_production_line(request):
    """AJAX view to add new production line form"""
    if request.method == 'POST':
        board_id = request.POST.get('board_id')
        board = get_object_or_404(PlanningBoard, pk=board_id, created_by=request.user)
        
        # Get the current count to set the form index
        count = board.production_lines.count()
        
        # Create empty form
        formset = ProductionLineFormSet(instance=board)
        empty_form = formset.empty_form
        
        # Replace __prefix__ with actual index
        form_html = str(empty_form).replace('__prefix__', str(count))
        
        return JsonResponse({'form_html': form_html})
    
    return JsonResponse({'error': 'Invalid request'})

@login_required
def planning_board_dashboard(request):
    """Dashboard view showing overview of all planning boards"""
    recent_boards = PlanningBoard.objects.filter(
        created_by=request.user
    ).order_by('-created_at')[:5]
    
    # Get some statistics
    total_boards = PlanningBoard.objects.filter(created_by=request.user).count()
    total_production_lines = ProductionLine.objects.filter(
        planning_board__created_by=request.user
    ).count()
    total_uploads = ExcelUpload.objects.filter(uploaded_by=request.user).count()
    
    context = {
        'recent_boards': recent_boards,
        'total_boards': total_boards,
        'total_production_lines': total_production_lines,
        'total_uploads': total_uploads,
    }
    
    return render(request, 'planning_board/dashboard.html', context)