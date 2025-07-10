# views.py - Complete fixed version with proper Excel processing
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
    """Process uploaded Excel file and populate database - enhanced version"""
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        worksheet = workbook.active
        
        print(f"Processing Excel file with {worksheet.max_row} rows and {worksheet.max_column} columns")
        
        # Debug the Excel structure first
        debug_excel_structure(worksheet)
        
        # Extract basic information
        extract_basic_info(worksheet, board)
        
        # Extract production lines data
        extract_production_lines(worksheet, board)
        
        # Extract future planning data with fixed approach
        extract_future_plans_fixed(worksheet, board)
        
        # Extract additional sections using enhanced approach
        extract_additional_sections(worksheet, board)
        
        return True
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False

def debug_excel_structure(worksheet):
    """Debug function to understand Excel layout"""
    try:
        print("=== DEBUGGING EXCEL STRUCTURE ===")
        
        # Look for plan headers specifically
        for row in range(1, 15):
            row_content = []
            for col in range(1, 35):
                cell_value = get_cell_value(worksheet, row, col)
                if cell_value and any(keyword in cell_value.upper() for keyword in ['TOMORROW', 'NEXT', 'DAY', 'PLAN', 'MODEL', 'SHIFT']):
                    col_letter = chr(64 + col) if col <= 26 else f"A{chr(64 + col - 26)}"
                    row_content.append(f"Col {col}({col_letter}): '{cell_value}'")
            
            if row_content:
                print(f"Row {row}: {' | '.join(row_content)}")
        
        print("=== END DEBUG ===")
    except Exception as e:
        print(f"Error in debug: {e}")

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
            {'name': 'CLUTCH ASSY LINE-3', 'start_row': 16, 'max_rows': 5},
            {'name': 'CLUTCH ASSY LINE-2', 'start_row': 10, 'max_rows': 5},
            {'name': 'PULLEY ASSY LINE-1', 'start_row': 7, 'max_rows': 2},
            {'name': 'FMD/FFD', 'start_row': 22, 'max_rows': 3},
            {'name': 'NEW BUSINESS', 'start_row': 26, 'max_rows': 5},
        ]
        
        for config in line_configs:
            extract_single_production_line(worksheet, board, config)
                
    except Exception as e:
        print(f"Error extracting production lines: {e}")

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

def extract_row_data(worksheet, row):
    """Extract all data from a single row"""
    try:
        return {
            # A Shift data
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
                'remarks': get_cell_value(worksheet, row, 19),    # Column S
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

def extract_future_plans_fixed(worksheet, board):
    """Extract tomorrow and next day plans with manual column specification - FIXED"""
    try:
        print("=== EXTRACTING FUTURE PLANS FIXED ===")
        
        # MANUAL COLUMN SPECIFICATION - Adjust these based on your Excel layout
        
        # TOMORROW PLAN COLUMNS (typically starts around column T/U)
        tomorrow_config = {
            'model_col': 20,     # Column T - TOMORROW MODEL
            'a_shift_col': 21,   # Column U - TOMORROW A SHIFT  
            'b_shift_col': 22,   # Column V - TOMORROW B SHIFT
            'c_shift_col': 23,   # Column W - TOMORROW C SHIFT
            'remarks_col': 24,   # Column X - TOMORROW REMARKS
            'start_row': 7,      # Data starts at row 7
            'end_row': 30        # Data ends at row 30
        }
        
        # NEXT DAY PLAN COLUMNS (typically starts around column Y/Z)
        next_day_config = {
            'model_col': 25,     # Column Y - NEXT DAY MODEL
            'a_shift_col': 26,   # Column Z - NEXT DAY A SHIFT
            'b_shift_col': 27,   # Column AA - NEXT DAY B SHIFT
            'c_shift_col': 28,   # Column AB - NEXT DAY C SHIFT
            'remarks_col': 29,   # Column AC - NEXT DAY REMARKS
            'start_row': 7,      # Data starts at row 7
            'end_row': 30        # Data ends at row 30
        }
        
        print("Extracting TOMORROW plans with config:", tomorrow_config)
        extract_plan_with_config_fixed(worksheet, board, 'tomorrow', tomorrow_config)
        
        print("Extracting NEXT DAY plans with config:", next_day_config)
        extract_plan_with_config_fixed(worksheet, board, 'next_day', next_day_config)
        
    except Exception as e:
        print(f"Error extracting future plans: {e}")

def extract_plan_with_config_fixed(worksheet, board, plan_type, config):
    """Extract plan data using manual configuration - FIXED"""
    try:
        print(f"\n=== EXTRACTING {plan_type.upper()} PLANS ===")
        print(f"Config: {config}")
        
        created_count = 0
        
        for row in range(config['start_row'], config['end_row']):
            model = get_cell_value(worksheet, row, config['model_col'])
            
            if model and model.strip() and len(model.strip()) > 1:
                model = model.strip()
                
                # Skip headers and common non-data entries
                if model.upper() in ['MODEL', 'SHIFT', 'REMARKS', 'A', 'B', 'C', 'PLAN', 'ASSY', 'DAY', 'TOMORROW', 'NEXT']:
                    print(f"  Skipping header row {row}: '{model}'")
                    continue
                
                # Get shift data
                a_shift = get_numeric_value(worksheet, row, config['a_shift_col'])
                b_shift = get_numeric_value(worksheet, row, config['b_shift_col'])
                c_shift = get_numeric_value(worksheet, row, config['c_shift_col'])
                remarks = get_cell_value(worksheet, row, config['remarks_col'])
                
                # Debug print
                print(f"  {plan_type} Row {row}: Model='{model}', A={a_shift}, B={b_shift}, C={c_shift}, Remarks='{remarks}'")
                
                # Create database entry if we have model name (quantity can be 0)
                if plan_type == 'tomorrow':
                    TomorrowPlan.objects.create(
                        planning_board=board,
                        model=model,
                        a_shift=a_shift or 0,
                        b_shift=b_shift or 0,
                        c_shift=c_shift or 0,
                        remarks=remarks or ''
                    )
                else:  # next_day
                    NextDayPlan.objects.create(
                        planning_board=board,
                        model=model,
                        a_shift=a_shift or 0,
                        b_shift=b_shift or 0,
                        c_shift=c_shift or 0,
                        remarks=remarks or ''
                    )
                
                created_count += 1
                print(f"  ✓ Created {plan_type} plan: {model}")
        
        print(f"=== COMPLETED {plan_type.upper()}: Created {created_count} entries ===\n")
                
    except Exception as e:
        print(f"Error extracting {plan_type} with config: {e}")
        import traceback
        traceback.print_exc()

def extract_additional_sections(worksheet, board):
    """Extract additional sections like Critical Parts, AFM, SPD, etc. - ENHANCED"""
    try:
        print("Starting to extract additional sections...")
        
        # Look for section keywords and their positions
        critical_row = find_section_header(worksheet, "CRITICAL")
        fcin_row = find_section_header(worksheet, "FCIN")
        iu_row = find_section_header(worksheet, "I/U")
        msil_row = find_section_header(worksheet, "MSIL")
        hmsi_row = find_section_header(worksheet, "HMSI")
        iymp_row = find_section_header(worksheet, "IYM")
        hmcl_row = find_section_header(worksheet, "HMCL")
        other_row = find_section_header(worksheet, "OTHER")
        
        # Extract sections with proper data filtering
        if critical_row:
            extract_critical_parts_fixed(worksheet, board, critical_row)
        
        if fcin_row:
            extract_afm_plans_fixed(worksheet, board, fcin_row, "FCIN")
            
        if iu_row:
            extract_afm_plans_fixed(worksheet, board, iu_row, "IU")
        
        if msil_row:
            extract_spd_plans_fixed(worksheet, board, msil_row, "MSIL")
            
        if hmsi_row:
            extract_spd_plans_fixed(worksheet, board, hmsi_row, "HMSI")
            
        if iymp_row:
            extract_spd_plans_fixed(worksheet, board, iymp_row, "IYM")
            
        if hmcl_row:
            extract_spd_plans_fixed(worksheet, board, hmcl_row, "HMCL")
        
        if other_row:
            extract_other_information_fixed(worksheet, board, other_row)
        
    except Exception as e:
        print(f"Error extracting additional sections: {e}")
        import traceback
        traceback.print_exc()

def extract_critical_parts_fixed(worksheet, board, start_row):
    """Extract critical parts - ENHANCED to scan more rows"""
    try:
        print(f"Extracting critical parts from row {start_row}")
        
        # Start from start_row + 2 to skip header row, scan more rows
        for row in range(start_row + 2, start_row + 25):
            part_name = get_cell_value(worksheet, row, 2)    # Column B
            supplier = get_cell_value(worksheet, row, 3)     # Column C  
            plan_qty = get_numeric_value(worksheet, row, 4)  # Column D
            receiving_time_str = get_cell_value(worksheet, row, 5) # Column E
            remarks = get_cell_value(worksheet, row, 6)      # Column F
            
            # Skip if no part name or empty
            if not part_name or not part_name.strip():
                continue
                
            # Skip header rows explicitly
            if part_name.upper() in ['PART NAME', 'PART', 'NAME', 'SUPPLIER', 'QTY', 'QUANTITY']:
                print(f"Skipping critical parts header: {part_name}")
                continue
            
            # Only process real data rows
            if len(part_name.strip()) > 2:
                CriticalPartStatus.objects.create(
                    planning_board=board,
                    part_name=part_name.strip(),
                    supplier=supplier.strip() if supplier else '',
                    plan_qty=plan_qty or 0,
                    remarks=remarks.strip() if remarks else ''
                )
                print(f"Created critical part: {part_name}")
                
    except Exception as e:
        print(f"Error extracting critical parts: {e}")

def extract_afm_plans_fixed(worksheet, board, start_row, plan_type):
    """Extract AFM plans - ENHANCED to scan more rows"""
    try:
        print(f"Extracting AFM {plan_type} from row {start_row}")
        
        # Determine column positions based on plan type
        if plan_type == "FCIN":
            part_col = 7   # Column G
            num_col = 8    # Column H
            qty_col = 9    # Column I  
            rem_col = 10   # Column J
        else:  # I/U
            part_col = 11  # Column K
            num_col = 12   # Column L
            qty_col = 13   # Column M
            rem_col = 14   # Column N
        
        # Extract data starting from row after header, scan more rows
        for row in range(start_row + 3, start_row + 25):
            part_name = get_cell_value(worksheet, row, part_col)
            part_number = get_cell_value(worksheet, row, num_col)
            plan_qty = get_numeric_value(worksheet, row, qty_col)
            remarks = get_cell_value(worksheet, row, rem_col)
            
            # Skip if no part name
            if not part_name or not part_name.strip():
                continue
                
            # Skip header rows
            if part_name.upper() in ['PART NAME', 'PART', 'NAME', 'SUPPLIER']:
                print(f"Skipping AFM header: {part_name}")
                continue
            
            # Only process real data
            if len(part_name.strip()) > 2:
                AFMPlan.objects.create(
                    planning_board=board,
                    plan_type=plan_type,
                    part_name=part_name.strip(),
                    part_number=part_number.strip() if part_number else '',
                    plan_qty=plan_qty or 0,
                    remarks=remarks.strip() if remarks else ''
                )
                print(f"Created AFM {plan_type} plan: {part_name}")
                
    except Exception as e:
        print(f"Error extracting AFM {plan_type}: {e}")

def extract_spd_plans_fixed(worksheet, board, start_row, customer):
    """Extract SPD plans - ENHANCED to scan more rows"""
    try:
        print(f"Extracting SPD {customer} from row {start_row}")
        
        # Column positions for different customers
        customer_cols = {
            'MSIL': {'part': 15, 'num': 16, 'qty': 17, 'rem': 18},    # O, P, Q, R
            'HMSI': {'part': 18, 'num': 19, 'qty': 20, 'rem': 21},   # R, S, T, U  
            'IYM': {'part': 21, 'num': 22, 'qty': 23, 'rem': 24},    # U, V, W, X
            'HMCL': {'part': 24, 'num': 25, 'qty': 26, 'rem': 27}    # X, Y, Z, AA
        }
        
        cols = customer_cols.get(customer, customer_cols['MSIL'])
        
        # Extract data starting from row after header, scan more rows
        for row in range(start_row + 2, start_row + 30):
            part_name = get_cell_value(worksheet, row, cols['part'])
            part_number = get_cell_value(worksheet, row, cols['num'])
            plan_qty = get_numeric_value(worksheet, row, cols['qty'])
            remarks = get_cell_value(worksheet, row, cols['rem'])
            
            # Skip if no part name
            if not part_name or not part_name.strip():
                continue
                
            # Skip header rows
            if part_name.upper() in ['PART NAME', 'PART', 'NAME', 'SUPPLIER']:
                print(f"Skipping SPD header: {part_name}")
                continue
            
            # Only process real data
            if len(part_name.strip()) > 2:
                SPDPlan.objects.create(
                    planning_board=board,
                    customer=customer,
                    part_name=part_name.strip(),
                    part_number=part_number.strip() if part_number else '',
                    plan_qty=plan_qty or 0,
                    remarks=remarks.strip() if remarks else ''
                )
                print(f"Created SPD {customer} plan: {part_name}")
                
    except Exception as e:
        print(f"Error extracting SPD {customer}: {e}")

def extract_other_information_fixed(worksheet, board, start_row):
    """Extract other information - ENHANCED to scan more rows"""
    try:
        print(f"Extracting other information from row {start_row}")
        
        # Extract data starting from row after header, scan more rows
        for row in range(start_row + 2, start_row + 25):
            part_name = get_cell_value(worksheet, row, 27)   # Column AA
            qty = get_numeric_value(worksheet, row, 28)      # Column AB
            target_date_str = get_cell_value(worksheet, row, 29) # Column AC
            remarks = get_cell_value(worksheet, row, 30)     # Column AD
            
            # Skip if no part name
            if not part_name or not part_name.strip():
                continue
                
            # Skip header rows
            if part_name.upper() in ['PART NAME', 'PART', 'NAME']:
                print(f"Skipping other info header: {part_name}")
                continue
            
            # Only process real data
            if len(part_name.strip()) > 2:
                # Use current date as default target date
                target_date = timezone.now().date()
                
                OtherInformation.objects.create(
                    planning_board=board,
                    part_name=part_name.strip(),
                    qty=qty or 0,
                    target_date=target_date,
                    remarks=remarks.strip() if remarks else ''
                )
                print(f"Created other information: {part_name}")
                
    except Exception as e:
        print(f"Error extracting other information: {e}")

def find_section_header(worksheet, keyword):
    """Find the row containing a specific section header keyword"""
    try:
        for row in range(1, worksheet.max_row + 1):
            for col in range(1, worksheet.max_column + 1):
                cell_value = get_cell_value(worksheet, row, col)
                if cell_value and keyword.upper() in cell_value.upper():
                    print(f"Found '{keyword}' at row {row}, col {col}: '{cell_value}'")
                    return row
        return None
    except Exception as e:
        print(f"Error finding section header for '{keyword}': {e}")
        return None

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
        worksheet.cell(row=row, column=24, value=plan.remarks)
        row += 1
    
    # Next day plans
    row = 7
    for plan in board.next_day_plans.all():
        worksheet.cell(row=row, column=25, value=plan.model)
        worksheet.cell(row=row, column=26, value=plan.a_shift)
        worksheet.cell(row=row, column=27, value=plan.b_shift)
        worksheet.cell(row=row, column=28, value=plan.c_shift)
        worksheet.cell(row=row, column=29, value=plan.remarks)
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