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
    """Extract production line data from specific rows - improved version"""
    try:
        # Define production lines and their starting row positions
        line_configs = [
            {'name': 'CLUTCH ASSY LINE-1', 'start_row': 7, 'max_rows': 5},
            {'name': 'CLUTCH ASSY LINE-2', 'start_row': 13, 'max_rows': 5},
            {'name': 'PULLEY ASSY LINE-1', 'start_row': 19, 'max_rows': 2},
            {'name': 'FMD/FFD', 'start_row': 22, 'max_rows': 3},
            {'name': 'NEW BUSINESS', 'start_row': 26, 'max_rows': 5},
        ]
        
        for config in line_configs:
            extract_single_production_line(worksheet, board, config)
                
    except Exception as e:
        print(f"Error extracting production lines: {e}")


def extract_row_data(worksheet, row):
    """Extract all data from a single row"""
    try:
        return {
            # A Shift data (adjust column numbers based on your Excel layout)
            'a_shift': {
                'model': get_cell_value(worksheet, row, 3),       # Column C
                'plan': get_numeric_value(worksheet, row, 4),     # Column D
                'actual': get_numeric_value(worksheet, row, 5),   # Column E
                'plan_change': get_numeric_value(worksheet, row, 6), # Column F
                'time': get_time_value(worksheet, row, 7),        # Column G
                'remarks': get_cell_value(worksheet, row, 8),     # Column H
            },
            # B Shift data
            'b_shift': {
                'model': get_cell_value(worksheet, row, 9),       # Column I
                'plan': get_numeric_value(worksheet, row, 10),    # Column J
                'actual': get_numeric_value(worksheet, row, 11),  # Column K
                'plan_change': get_numeric_value(worksheet, row, 12), # Column L
                'time': get_time_value(worksheet, row, 13),       # Column M
                'remarks': get_cell_value(worksheet, row, 14),    # Column N
            },
            # C Shift data
            'c_shift': {
                'model': get_cell_value(worksheet, row, 15),      # Column O
                'plan': get_numeric_value(worksheet, row, 16),    # Column P
                'actual': get_numeric_value(worksheet, row, 17),  # Column Q
                'plan_change': get_numeric_value(worksheet, row, 18), # Column R
                'time': get_time_value(worksheet, row, 19),       # Column S
                'remarks': get_cell_value(worksheet, row, 20),    # Column T
            }
        }
    except Exception as e:
        print(f"Error extracting row {row}: {e}")
        return None
    
def has_meaningful_data(row_data):
    """Check if a row contains meaningful data"""
    if not row_data:
        return False
    
    # Check if any shift has a model name (indicating actual data)
    for shift in ['a_shift', 'b_shift', 'c_shift']:
        if row_data[shift]['model'] and row_data[shift]['model'].strip():
            # Exclude header rows
            model = row_data[shift]['model'].strip().upper()
            if model not in ['MODEL', 'SHIFT', 'LINE', 'NO.', '']:
                return True
    
    return False

def create_production_line_entries(board, line_name, line_entries):
    """Create production line database entries from extracted data"""
    try:
        entry_count = 0
        
        for entry_data in line_entries:
            # Determine if this should be a separate ProductionLine or combined
            # For now, we'll create separate entries for each meaningful row
            
            # Check which shifts have data
            shifts_with_data = []
            for shift_name in ['a_shift', 'b_shift', 'c_shift']:
                if entry_data[shift_name]['model'] and entry_data[shift_name]['model'].strip():
                    shifts_with_data.append(shift_name)
            
            if shifts_with_data:
                entry_count += 1
                # Create line name with entry number if multiple entries
                display_name = f"{line_name}" if entry_count == 1 else f"{line_name} - Entry {entry_count}"
                
                # Create the production line
                production_line = ProductionLine.objects.create(
                    planning_board=board,
                    line_number=display_name,
                    # A Shift
                    a_shift_model=entry_data['a_shift']['model'],
                    a_shift_plan=entry_data['a_shift']['plan'],
                    a_shift_actual=entry_data['a_shift']['actual'],
                    a_shift_plan_change=entry_data['a_shift']['plan_change'],
                    a_shift_time=entry_data['a_shift']['time'],
                    a_shift_remarks=entry_data['a_shift']['remarks'],
                    # B Shift
                    b_shift_model=entry_data['b_shift']['model'],
                    b_shift_plan=entry_data['b_shift']['plan'],
                    b_shift_actual=entry_data['b_shift']['actual'],
                    b_shift_plan_change=entry_data['b_shift']['plan_change'],
                    b_shift_time=entry_data['b_shift']['time'],
                    b_shift_remarks=entry_data['b_shift']['remarks'],
                    # C Shift
                    c_shift_model=entry_data['c_shift']['model'],
                    c_shift_plan=entry_data['c_shift']['plan'],
                    c_shift_actual=entry_data['c_shift']['actual'],
                    c_shift_plan_change=entry_data['c_shift']['plan_change'],
                    c_shift_remarks=entry_data['c_shift']['remarks'],
                )
                
                print(f"Created production line: {display_name}")
                
    except Exception as e:
        print(f"Error creating production line entries for {line_name}: {e}")

def get_time_value(worksheet, row, col):
    """Get cell value as time, handling various time formats"""
    try:
        cell = worksheet.cell(row=row, column=col)
        value = cell.value
        
        if value is None:
            return None
            
        # If it's already a time object
        if hasattr(value, 'hour'):
            return value
            
        # Try to parse string time formats
        if isinstance(value, str):
            import re
            from datetime import time
            
            # Match HH:MM format
            time_match = re.search(r'(\d{1,2}):(\d{2})', value)
            if time_match:
                hour, minute = int(time_match.group(1)), int(time_match.group(2))
                if 0 <= hour <= 23 and 0 <= minute <= 59:
                    return time(hour, minute)
        
        return None
    except:
        return None

def extract_single_production_line(worksheet, board, config):
    """Extract data for a single production line with multiple entries"""
    try:
        line_name = config['name']
        start_row = config['start_row']
        max_rows = config.get('max_rows', 5)
        
        # Track all data entries for this production line
        line_entries = []
        
        # Scan through all possible rows for this production line
        for row_offset in range(max_rows):
            current_row = start_row + row_offset
            
            # Check if this row has any meaningful data
            row_data = extract_row_data(worksheet, current_row)
            
            if has_meaningful_data(row_data):
                line_entries.append(row_data)
        
        # Create production line entries
        if line_entries:
            create_production_line_entries(board, line_name, line_entries)
        else:
            # Create empty production line if no data found
            ProductionLine.objects.create(
                planning_board=board,
                line_number=line_name
            )
            
    except Exception as e:
        print(f"Error extracting line {config['name']}: {e}")

def extract_plan_section(worksheet, board, plan_type, start_col, model_col, 
                        a_shift_col, b_shift_col, c_shift_col, remarks_col):
    """Extract a specific plan section (tomorrow or next day)"""
    try:
        # Scan through rows looking for data
        for row in range(5, 35):  # Expanded row range
            model = get_cell_value(worksheet, row, model_col)
            
            if model and model.strip():
                model = model.strip().upper()
                
                # Skip header rows
                if model in ['MODEL', 'SHIFT', 'LINE', 'NO.', 'DATE', 'ASSY', 'PLAN']:
                    continue
                
                # Get shift data
                a_shift = get_numeric_value(worksheet, row, a_shift_col)
                b_shift = get_numeric_value(worksheet, row, b_shift_col)
                c_shift = get_numeric_value(worksheet, row, c_shift_col)
                remarks = get_cell_value(worksheet, row, remarks_col)
                
                # Create plan entry if we have meaningful data
                if a_shift or b_shift or c_shift:
                    if plan_type == 'tomorrow':
                        TomorrowPlan.objects.create(
                            planning_board=board,
                            model=model,
                            a_shift=a_shift,
                            b_shift=b_shift,
                            c_shift=c_shift,
                            remarks=remarks
                        )
                    else:  # next_day
                        NextDayPlan.objects.create(
                            planning_board=board,
                            model=model,
                            a_shift=a_shift,
                            b_shift=b_shift,
                            c_shift=c_shift,
                            remarks=remarks
                        )
                    
                    print(f"Created {plan_type} plan for model: {model}")
                
    except Exception as e:
        print(f"Error extracting {plan_type} plans: {e}")
def extract_additional_sections(worksheet, board):
    """Extract additional sections like Critical Parts, AFM, SPD, etc."""
    try:
        # Find and extract Critical Parts section
        extract_critical_parts(worksheet, board)
        
        # Find and extract AFM Plans section
        extract_afm_plans(worksheet, board)
        
        # Find and extract SPD Plans section
        extract_spd_plans(worksheet, board)
        
        # Find and extract Other Information section
        extract_other_information(worksheet, board)
        
    except Exception as e:
        print(f"Error extracting additional sections: {e}")
        
def find_section_header(worksheet, keyword):
    """Find the row containing a specific section header keyword"""
    try:
        for row in range(1, 100):  # Search in reasonable range
            for col in range(1, 30):
                cell_value = get_cell_value(worksheet, row, col)
                if keyword.upper() in cell_value.upper():
                    return row
        return None
    except:
        return None

# Helper function improvements
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
            return int(value) if value == int(value) else value
        # Try to convert string to number
        try:
            num_val = float(str(value).replace(',', ''))
            return int(num_val) if num_val == int(num_val) else num_val
        except:
            return None
    except:
        return None        
def extract_critical_parts(worksheet, board):
    """Extract Critical Part Status section"""
    try:
        # Look for "CRITICAL" keyword to find the section
        critical_section_row = find_section_header(worksheet, "CRITICAL")
        
        if critical_section_row:
            # Extract data starting from a few rows after the header
            for row in range(critical_section_row + 2, critical_section_row + 15):
                part_name = get_cell_value(worksheet, row, 2)  # Adjust column as needed
                supplier = get_cell_value(worksheet, row, 3)
                plan_qty = get_numeric_value(worksheet, row, 4)
                
                if part_name and part_name.strip():
                    CriticalPartStatus.objects.create(
                        planning_board=board,
                        part_name=part_name,
                        supplier=supplier or '',
                        plan_qty=plan_qty or 0,
                        remarks=get_cell_value(worksheet, row, 6)
                    )
                    
    except Exception as e:
        print(f"Error extracting critical parts: {e}")

def extract_future_plans(worksheet, board):
    """Extract tomorrow and next day plans - improved version"""
    try:
        # Tomorrow plans - scan more comprehensively
        extract_plan_section(worksheet, board, 'tomorrow', 
                            start_col=21, model_col=21, 
                            a_shift_col=22, b_shift_col=23, c_shift_col=24, 
                            remarks_col=25)
        
        # Next day plans
        extract_plan_section(worksheet, board, 'next_day', 
                            start_col=26, model_col=26, 
                            a_shift_col=27, b_shift_col=28, c_shift_col=29, 
                            remarks_col=30)
                
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