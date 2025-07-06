from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from planning_board.models import PlanningBoard, ProductionLine, TomorrowPlan, NextDayPlan
from datetime import date, timedelta, time

class Command(BaseCommand):
    help = 'Create sample planning board data for testing'

    def handle(self, *args, **options):
        # Get or create a user
        user, created = User.objects.get_or_create(username='admin')
        if created:
            user.set_password('admin123')
            user.is_superuser = True
            user.is_staff = True
            user.save()
            self.stdout.write(f"Created admin user")
        
        # Create sample planning board
        today = date.today()
        board = PlanningBoard.objects.create(
            title="Sample Production Planning Board",
            meeting_time=time(9, 0),  # 9:00 AM
            today_date=today,
            tomorrow_date=today + timedelta(days=1),
            next_day_date=today + timedelta(days=2),
            created_by=user
        )
        
        # Create sample production lines
        production_lines = [
            {
                'line_number': 'CLUTCH ASSY LINE-1',
                'a_shift_model': 'Model CA-100',
                'a_shift_plan': 150,
                'a_shift_actual': 145,
                'b_shift_model': 'Model CA-200',
                'b_shift_plan': 120,
                'b_shift_actual': 118,
                'c_shift_model': 'Model CA-300',
                'c_shift_plan': 100,
                'c_shift_actual': 95,
            },
            {
                'line_number': 'CLUTCH ASSY LINE-2',
                'a_shift_model': 'Model CB-100',
                'a_shift_plan': 180,
                'a_shift_actual': 175,
                'b_shift_model': 'Model CB-200',
                'b_shift_plan': 160,
                'b_shift_actual': 155,
            },
            {
                'line_number': 'PULLEY ASSY LINE-1',
                'a_shift_model': 'Model PA-100',
                'a_shift_plan': 200,
                'a_shift_actual': 190,
                'c_shift_model': 'Model PA-300',
                'c_shift_plan': 80,
                'c_shift_actual': 85,
            }
        ]
        
        for line_data in production_lines:
            ProductionLine.objects.create(
                planning_board=board,
                **line_data
            )
        
        # Create sample tomorrow plans
        tomorrow_plans = [
            {'model': 'Model X-100', 'a_shift': 100, 'b_shift': 80, 'c_shift': 60},
            {'model': 'Model Y-200', 'a_shift': 120, 'b_shift': 90, 'c_shift': 70},
        ]
        
        for plan_data in tomorrow_plans:
            TomorrowPlan.objects.create(
                planning_board=board,
                **plan_data
            )
        
        # Create sample next day plans
        next_day_plans = [
            {'model': 'Model Z-300', 'a_shift': 90, 'b_shift': 70, 'c_shift': 50},
            {'model': 'Model W-400', 'a_shift': 110, 'b_shift': 85, 'c_shift': 65},
        ]
        
        for plan_data in next_day_plans:
            NextDayPlan.objects.create(
                planning_board=board,
                **plan_data
            )
        
        self.stdout.write(
            self.style.SUCCESS(
                f"Successfully created sample data!\n"
                f"Planning Board ID: {board.id}\n"
                f"Production Lines: {board.production_lines.count()}\n"
                f"Tomorrow Plans: {board.tomorrow_plans.count()}\n"
                f"Next Day Plans: {board.next_day_plans.count()}\n"
                f"User: {user.username} (password: admin123)"
            )
        )