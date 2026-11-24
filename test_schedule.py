import calendar
import unittest
from datetime import date, timedelta

from app import WEEKDAYS, build_month_schedule


class MonthlyScheduleTests(unittest.TestCase):
    def test_each_sunday_to_saturday_week_has_two_consecutive_days_off(self):
        employees = [
            {
                "Canal": "Boutique",
                "Loja": "Loja de teste",
                "Colaborador": f"Pessoa {index:02d}",
                "Cargo": "Gerente de Loja" if index == 0 else "Vendedor",
            }
            for index in range(12)
        ]
        off_dates = {employee["Colaborador"]: set() for employee in employees}

        for month in (8, 9):
            schedule, _, _, _ = build_month_schedule(employees, 2026, month)
            for _, row in schedule.iterrows():
                employee = row["Colaborador"]
                for day in range(1, calendar.monthrange(2026, month)[1] + 1):
                    column = f"{day:02d} {WEEKDAYS[calendar.weekday(2026, month, day)]}"
                    if row[column] == "F":
                        off_dates[employee].add(date(2026, month, day))

        first_sunday = date(2026, 8, 2)
        last_saturday = date(2026, 9, 26)
        for employee, days_off in off_dates.items():
            week_start = first_sunday
            while week_start + timedelta(days=6) <= last_saturday:
                weekly_off = sorted(
                    day
                    for day in days_off
                    if week_start <= day <= week_start + timedelta(days=6)
                )
                with self.subTest(employee=employee, week=week_start):
                    self.assertEqual(len(weekly_off), 2)
                    self.assertEqual(weekly_off[1] - weekly_off[0], timedelta(days=1))
                    self.assertNotEqual(weekly_off[0].weekday(), 5)
                    self.assertNotEqual(weekly_off[1].weekday(), 5)
                week_start += timedelta(days=7)


if __name__ == "__main__":
    unittest.main()
