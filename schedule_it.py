'''Create a flat schedule from CSV input files'''
import csv, random, sys
from datetime import datetime, timedelta, time

pname = 'schedule_it.py'
usage = 'YEAR_NUM MONTH_NUM [availability.csv] [holidays.csv] [VERBOSE]'
# Test for correct number of arguments or print usage
if len(sys.argv) < 2:
    print(f'Usage: {pname} {usage}')
    print('Example: schedule_it.py 2025 3')
    print('   ^^^^^^ assumes files are availability.csv and holidays.csv')
    print(f'Example: {pname} 2025 3 diff_avail_file.csv my_holiday_list.csv')
    sys.exit(1)

verbose = False
available = 'availability.csv'
holidays = 'holidays.csv'
schedule_out = 'flat_sched.csv'
out_headers = ['Date', 'IsWeekday', 'IsHoliday', 'EarlyDutyInvestigator',
                'EarlyOperationsOfficer', 'LateDutyInvestigator',
                'LateOperationsOfficer']
DAYS = {0:'Monday', 1:'Tuesday', 2:'Wednesday', 3:'Thursday', 4:'Friday'}
try:
    YEAR = int(sys.argv[1])
    MONTH = int(sys.argv[2])
    if len(sys.argv) > 3:
        available = sys.argv[3]
    if len(sys.argv) > 4:
        holidays = sys.argv[4]
    if len(sys.argv) > 5:
        if sys.argv[5].upper() == 'VERBOSE':
            verbose = True
except (IndexError, ValueError):
    print('Error: Invalid input. YEAR and MONTH must be integers and the')
    print('  first two arguments')
    sys.exit(1)

morning_start = time(8, 00)
afternoon_finish = time(17, 00)

workers = dict()
skip_list = []
in_rotation = []
last_month = []
force_in = []
# line format is:
headers = {'Person':0, 'Monday':1, 'Tuesday':2, 'Wednesday':3,
           'Thursday':4, 'Friday':5, 'DayStart':6, 'DayFinish':7, 'OpsOK':8,
           'SkipList':9, 'InRotation':10, 'WentLastMonth':11,
           'ForceInFirst':12}

# We are not using the blank column and InRotation will be calculated and
#  should match what shows in the sheet
with open(available) as f:
    reader = csv.reader(f, delimiter=',')
    for i, row in enumerate(reader):
        if i == 0:
            continue
        if verbose == True:
            print(f'processing row: {row}')
        name = row[headers['Person']].strip()
        if name == '':
            continue
        mon = row[headers['Monday']] == 'TRUE'
        tue = row[headers['Tuesday']] == 'TRUE'
        wed = row[headers['Wednesday']] == 'TRUE'
        thu = row[headers['Thursday']] == 'TRUE'
        fri = row[headers['Friday']] == 'TRUE'
        hour, minute = map(int, row[headers['DayStart']].split(':'))
        d_start = time(hour, minute)
        hour, minute = map(int, row[headers['DayFinish']].split(':'))
        d_end = time(hour, minute)
        ops_ok = row[headers['OpsOK']] == 'TRUE'
        if row[headers['SkipList']] != '':
            skip_list.append(row[headers['SkipList']].strip())
        if row[headers['WentLastMonth']] != '':
            last_month.append(row[headers['WentLastMonth']].strip())
        if row[headers['ForceInFirst']] != '':
            force_in.append(row[headers['ForceInFirst']].strip())
        if name in workers:
            print(f'Error: did worker {name} get listed twice? line {i+2}')
        workers[name] = {'Monday':mon, 'Tuesday':tue, 'Wednesday':wed,
                        'Thursday':thu, 'Friday':fri, 'Start':d_start,
                        'End':d_end, 'OpsOK':ops_ok}

# get the holidays
holiday_list = []
try:
    with open(holidays) as f:
        ls = f.read().strip().split('\n')
    for l in ls[1:]:
        day, name = l.split(',')
        month, day, year = map(int, day.split('/'))
        holiday = datetime(year, month, day)
        holiday_list.append(holiday)
except:
    print(f'Error: could not read {holidays} file or it is not formatted')
    print('  correctly. Holidays should be in MM/DD/YYYY format')
    sys.exit(1)

days = dict()
current_day = datetime(YEAR, MONTH, 1)
while current_day.month == MONTH:
    is_weekday = current_day.weekday() < 5  # Mon to Fri are weekdays
    is_holiday = current_day in holiday_list
    slots = ['NONE' for x in range(4)]
    if is_holiday:
        slots = ['HOLIDAY' for x in range(4)]
    days[current_day] = {'IsWeekday':is_weekday, 'IsHoliday':is_holiday,
                         'Slots':slots}
    current_day += timedelta(days=1)

def findFit(name, days, p_info):
    '''Look for a scheduling fit and put the person in the first fitting
    slot. Return True if a fit was found and False otherwise'''
    for day in sorted(days.keys()):
        if not days[day]['IsWeekday'] or days[day]['IsHoliday']:
            continue
        if 'NONE' not in days[day]['Slots']:
            continue
        day_of_week = DAYS[day.weekday()]
        # this part dependant on num slots being filled, type and timing
        if days[day]['Slots'][0] == 'NONE' and not p_info[name]['OpsOK']:
            if p_info[name]['Start'] <= morning_start:
                if p_info[name][day_of_week]:
                    days[day]['Slots'][0] = name
                    return(True)
        if days[day]['Slots'][1] == 'NONE' and p_info[name]['OpsOK']:
            if p_info[name]['Start'] <= morning_start:
                if p_info[name][day_of_week]:
                    days[day]['Slots'][1] = name
                    return(True)
        if days[day]['Slots'][2] == 'NONE' and not p_info[name]['OpsOK']:
            if p_info[name]['End'] >= afternoon_finish:
                if p_info[name][day_of_week]:
                    days[day]['Slots'][2] = name
                    return(True)
        if days[day]['Slots'][3] == 'NONE' and p_info[name]['OpsOK']:
            if p_info[name]['End'] >= afternoon_finish:
                if p_info[name][day_of_week]:
                    days[day]['Slots'][3] = name
                    return(True)
    return(False)

def isCalDone(days, role=3):
    '''Test if all days that need a slot filled are filled. If role is 1,
    just check for "duty" slots, if role is 2, just check for "ops" slots
    and if 3, check for both. Return True if done and False if not'''
    if role == 3:
        for day in days.keys():
            if not days[day]['IsWeekday'] or days[day]['IsHoliday']:
                continue
            if 'NONE' in days[day]['Slots']:
                return(False)
    if role == 2:
        for day in days.keys():
            if not days[day]['IsWeekday'] or days[day]['IsHoliday']:
                continue
            if days[day]['Slots'][1] == 'NONE' or \
                days[day]['Slots'][3] == 'NONE':
                return(False)
    if role == 1:
        for day in days.keys():
            if not days[day]['IsWeekday'] or days[day]['IsHoliday']:
                continue
            if days[day]['Slots'][0] == 'NONE' or \
                days[day]['Slots'][2] == 'NONE':
                return(False)
    return(True)

# start fitting from the force_in list
#  this section will apply to both "duty" and "ops"
already_in = []
while len(force_in) > 0:
    name = force_in.pop()
    fit_found = findFit(name, days, workers)
    if fit_found:
        already_in.append(name)
    else:
        print(f'Warning: could not find a fit for {name}', end='')
        print(' from the ForceIn list')

# make a list of those that didn't go last time and draw from it
#  this part will be for "duty" only
try_list = []
for person in workers.keys():
    if person not in last_month and person not in already_in and \
        person not in skip_list and not workers[person]['OpsOK']:
        try_list.append(person)
random.shuffle(try_list)
done = False
while len(try_list) > 0 and not done:
    name = try_list.pop()
    fit_found = findFit(name, days, workers)
    if fit_found:
        already_in.append(name)
    elif isCalDone(days, 1):
        done = True
        break
    elif verbose == True:
        print(f'Note: could not find a fit for {name}', end='')
        print(' from the not-in-last-month duty list')

# draw from the remaining officers randomly, just for "duty" role
try_list = []
for person in workers.keys():
    if person not in already_in and not workers[person]['OpsOK'] and \
        person not in skip_list:
        try_list.append(person)
random.shuffle(try_list)
while len(try_list) > 0 and not done:
    name = try_list.pop()
    fit_found = findFit(name, days, workers)
    if fit_found:
        already_in.append(name)
    elif isCalDone(days, 1):
        done = True
        if verbose == True:
            print(f'Finished duty schedule with {len(try_list)}', end='')
            print(' people still in the list')
        break
    elif verbose == True:
        print(f'Note: could not find a fit for {name}', end='')
        print(' from the full list for duty schedule')
# we should be done with the "duty" list
if not done:
    print("Error: Was not able to fill the duty schedule")

# now fill the "ops" schedule
done = False
loop_count = 0
while not done:
    if loop_count > 100:
        print('Error: seem unable to fill ops schedule')
        break
    try_list = []
    for person in workers.keys():
        if person not in already_in and workers[person]['OpsOK'] and \
            person not in skip_list:
            try_list.append(person)
    random.shuffle(try_list)
    while len(try_list) > 0 and not done:
        name = try_list.pop()
        fit_found = findFit(name, days, workers)
        if isCalDone(days, 2):
            done = True
            break
    loop_count += 1

# all slots should now be filled
if not isCalDone(days):
    print('Error: failed final check that calendar is complete')

# output the flat schedule for input to Excel
with open(schedule_out, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(out_headers)
    for day in sorted(days.keys()):
        row = []
        day_str = day.strftime("%m/%d/%Y")
        row.append(day_str)
        if days[day]['IsWeekday']:
            row.append(1)
        else:
            row.append(0)
        if days[day]['IsHoliday']:
            row.append(1)
        else:
            row.append(0)
        row.extend(days[day]['Slots'])
        writer.writerow(row)
