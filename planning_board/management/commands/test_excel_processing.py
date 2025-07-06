from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from planning_board.models import PlanningBoard
from planning_board.views import process_excel_file
from datetime import date, timedelta
import os

class Command(BaseCommand):
    help = 'Test Excel file processing with a sample file'

    def add_arguments(self, parser):
        parser.add_argument('excel_file_path', type=str, help='Path to Excel file to test')

    def handle(self, *args, **options):
        excel_file_path = options['excel_file_path']
        
        if not os.path.exists(excel_file_path):
            self.stdout.write(self.style.ERROR(f"File does not exist: {excel_file_path}"))
            return
        
        # Get or create a test user
        user, created = User.objects.get_or_create(username='test_user')
        if created:
            user.set_password('test123')
            user.save()
            self.stdout.write(f"Created test user: {user.username}")
        
        # Create a test planning board
        today = date.today()
        board = PlanningBoard.objects.create(
            title="Test Excel Upload",
            created_by=user,
            today_date=today,
            tomorrow_date=today + timedelta(days=1),
            next_day_date=today + timedelta(days=2),
        )
        
        self.stdout.write(f"Created test planning board with ID: {board.id}")
        
        try:
            # Test Excel processing
            success = process_excel_file(excel_file_path, board)
            
            if success:
                self.stdout.write(self.style.SUCCESS("Excel processing completed successfully!"))
                
                # Show results
                self.stdout.write(f"\nResults:")
                self.stdout.write(f"  Production Lines: {board.production_lines.count()}")
                self.stdout.write(f"  Tomorrow Plans: {board.tomorrow_plans.count()}")
                self.stdout.write(f"  Next Day Plans: {board.next_day_plans.count()}")
                self.stdout.write(f"  Critical Parts: {board.critical_parts.count()}")
                
                # Show production line details
                for line in board.production_lines.all():
                    self.stdout.write(f"  Line: {line.line_number}")
                    if line.a_shift_model:
                        self.stdout.write(f"    A Shift: {line.a_shift_model} (Plan: {line.a_shift_plan})")
                    if line.b_shift_model:
                        self.stdout.write(f"    B Shift: {line.b_shift_model} (Plan: {line.b_shift_plan})")
                    if line.c_shift_model:
                        self.stdout.write(f"    C Shift: {line.c_shift_model} (Plan: {line.c_shift_plan})")
                
            else:
                self.stdout.write(self.style.ERROR("Excel processing failed!"))
                
        except Exception as e:
            self.stdout.write(self.style.ERROR(f"Error during processing: {str(e)}"))
            import traceback
            self.stdout.write(traceback.format_exc())
