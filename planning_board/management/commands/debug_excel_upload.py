from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from planning_board.models import PlanningBoard, ExcelUpload
import os

class Command(BaseCommand):
    help = 'Debug Excel upload issues and show uploaded data'

    def handle(self, *args, **options):
        self.stdout.write("=== Planning Board Debug Information ===\n")
        
        # Check total planning boards
        total_boards = PlanningBoard.objects.count()
        self.stdout.write(f"Total Planning Boards: {total_boards}")
        
        # Check recent planning boards
        recent_boards = PlanningBoard.objects.order_by('-created_at')[:5]
        self.stdout.write(f"\nRecent Planning Boards:")
        for board in recent_boards:
            self.stdout.write(f"  - ID: {board.id}, Title: {board.title}, Date: {board.today_date}, Created: {board.created_at}")
            self.stdout.write(f"    Production Lines: {board.production_lines.count()}")
            self.stdout.write(f"    Tomorrow Plans: {board.tomorrow_plans.count()}")
            self.stdout.write(f"    Critical Parts: {board.critical_parts.count()}")
        
        # Check Excel uploads
        total_uploads = ExcelUpload.objects.count()
        self.stdout.write(f"\nTotal Excel Uploads: {total_uploads}")
        
        recent_uploads = ExcelUpload.objects.order_by('-uploaded_at')[:5]
        self.stdout.write(f"\nRecent Excel Uploads:")
        for upload in recent_uploads:
            self.stdout.write(f"  - File: {upload.file.name}, Processed: {upload.processed}, Date: {upload.uploaded_at}")
            if upload.planning_board:
                self.stdout.write(f"    Associated Board ID: {upload.planning_board.id}")
            
            # Check if file exists
            if upload.file:
                file_path = upload.file.path
                file_exists = os.path.exists(file_path)
                self.stdout.write(f"    File exists: {file_exists}")
                if file_exists:
                    file_size = os.path.getsize(file_path)
                    self.stdout.write(f"    File size: {file_size} bytes")
        
        # Show users
        users = User.objects.all()
        self.stdout.write(f"\nUsers in system:")
        for user in users:
            user_boards = PlanningBoard.objects.filter(created_by=user).count()
            self.stdout.write(f"  - {user.username}: {user_boards} boards")
        
        self.stdout.write("\n=== End Debug Information ===")
