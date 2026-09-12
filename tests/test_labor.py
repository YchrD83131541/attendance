import os
import csv
import datetime
import tempfile
import unittest

import labor

class LaborTests(unittest.TestCase):
    def setUp(self):
        # create a temp CSV
        fd, path = tempfile.mkstemp(prefix="att_", suffix=".csv", dir=os.getcwd())
        os.close(fd)
        self.path = path
        with open(self.path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id','name','date','checkin','checkout','break_minutes','admin_minutes','lessons_main','lessons_sub'])
            # create sample rows across months
            # Staff A: 750 hours in year 2026
            # We'll create many 8-hour days to sum approx 750
            d = datetime.date(2026,1,1)
            idc = 1
            for i in range(100):
                rowdate = (d + datetime.timedelta(days=i)).isoformat()
                checkin = f"{rowdate}T09:00:00"
                checkout = f"{rowdate}T18:30:00"  # 9.5h -> minus break 1h = 8.5
                writer.writerow([str(idc), 'StaffA', rowdate, checkin, checkout, '60', '0', '0', '0'])
                idc += 1
            # Staff B: moderate
            for i in range(20):
                rowdate = (d + datetime.timedelta(days=i)).isoformat()
                checkin = f"{rowdate}T09:00:00"
                checkout = f"{rowdate}T17:00:00"
                writer.writerow([str(idc), 'StaffB', rowdate, checkin, checkout, '60', '0', '0', '0'])
                idc += 1

    def tearDown(self):
        try:
            os.remove(self.path)
        except Exception:
            pass

    def test_annual_720(self):
        am = labor.AttendanceManager(self.path)
        lm = labor.LaborAgreementManager(am)
        res = lm.compute_annual_summary(2026)
        # find StaffA
        sa = [r for r in res if r['name']=='StaffA']
        self.assertTrue(len(sa)==1)
        self.assertTrue(sa[0]['total_hours']>=720)
        self.assertTrue(sa[0]['over_720'])

    def test_multi_month_average(self):
        am = labor.AttendanceManager(self.path)
        lm = labor.LaborAgreementManager(am)
        # Use end month Jan 2026, window 3 -> months Jan, Dec 2025, Nov 2025; only data in Jan
        res = lm.compute_multi_month_average(2026,1,window=3)
        # StaffA average should be total_jan/3, still may exceed 80 depending on data; ensure structure
        sa = [r for r in res if r['name']=='StaffA']
        self.assertTrue(len(sa)==1)
        self.assertIn('average_overtime', sa[0])

if __name__ == '__main__':
    unittest.main()
