# Create a flat schedule from CSV input files

# Constants and configuration
$DEBUG = $false
$available = "availability.csv"
$holidays = "holidays.csv"
$schedule_out = "flat_sched.csv"
$out_headers = @('Date', 'IsWeekday', 'IsHoliday', 'EarlyDutyInvestigator',
                'EarlyOperationsOfficer', 'LateDutyInvestigator',
                'LateOperationsOfficer')
$DAYS = @{
    0 = 'Monday'
    1 = 'Tuesday'
    2 = 'Wednesday'
    3 = 'Thursday'
    4 = 'Friday'
}

# check our command line arguments
if ($args.Length -lt 2) {
    Write-Host "Usage: schedule_it.ps1 YEAR_NUM MONTH_NUM [availability.csv] [holidays.csv] [VERBOSE]"
    Write-Host "Example: schedule_it.ps1 2025 3"
    Write-Host "Example: schedule_it.ps1 2025 3 my_times.csv my_holidays.csv"
    exit 1
}
$arg0IsInt = [int]::TryParse($args[0], [ref]$null)
$arg1IsInt = [int]::TryParse($args[1], [ref]$null)
if (-not $arg0IsInt -or -not $arg1IsInt) {
    Write-Host "Error: YEAR_NUM and MONTH_NUM must be integers"
    exit 1
} else {
    $YEAR = [int]$args[0]
    $MONTH = [int]$args[1]
}
if ($args.Length -gt 2) {
    $available = $args[2]
}
if ($args.Length -gt 3) {
    $holidays = $args[3]
}
if ($args.Length -gt 4) {
    if ($args[4] -eq "VERBOSE") {
        $DEBUG = $true
    }
}

$morning_start = (Get-Date "08:00").TimeOfDay
$afternoon_finish = (Get-Date "17:00").TimeOfDay

# Initialize data structures
$workers = @{}
$skip_list = @()
$in_rotation = @()
$last_month = @()
$force_in = @()

# Header mappings
$headers = @{
    'Person' = 0
    'Monday' = 1
    'Tuesday' = 2
    'Wednesday' = 3
    'Thursday' = 4
    'Friday' = 5
    'DayStart' = 6
    'DayFinish' = 7
    'OpsOK' = 8
    'SkipList' = 9
    'InRotation' = 10
    'WentLastMonth' = 11
    'ForceInFirst' = 12
}

# Read availability CSV
$csv_content = Import-Csv $available
foreach ($row in $csv_content) {
    if ([string]::IsNullOrWhiteSpace($row.Person)) { continue }
    
    if ($DEBUG) {
        Write-Host "Processing row: $($row.Person)"
    }

    $name = $row.Person.Trim()
    $mon = $row.Monday -eq "TRUE"
    $tue = $row.Tuesday -eq "TRUE"
    $wed = $row.Wednesday -eq "TRUE"
    $thu = $row.Thursday -eq "TRUE"
    $fri = $row.Friday -eq "TRUE"
    
    $d_start = [TimeSpan]::Parse($row.DayStart)
    $d_end = [TimeSpan]::Parse($row.DayFinish)
    $ops_ok = $row.OpsOK -eq "TRUE"
    
    if (![string]::IsNullOrWhiteSpace($row.SkipList)) {
        $skip_list += $row.SkipList.Trim()
    }
    if (![string]::IsNullOrWhiteSpace($row.WentLastMonth)) {
        $last_month += $row.WentLastMonth.Trim()
    }
    if (![string]::IsNullOrWhiteSpace($row.ForceInFirst)) {
        $force_in += $row.ForceInFirst.Trim()
    }
    
    if ($workers.ContainsKey($name)) {
        Write-Host "Error: did worker $name get listed twice?"
    }
    
    $workers[$name] = @{
        'Monday' = $mon
        'Tuesday' = $tue
        'Wednesday' = $wed
        'Thursday' = $thu
        'Friday' = $fri
        'Start' = $d_start
        'End' = $d_end
        'OpsOK' = $ops_ok
    }
}

# Read holidays
$holiday_list = @()
$holidays_content = Get-Content $holidays | Select-Object -Skip 1
foreach ($line in $holidays_content) {
    $day, $name = $line -split ','
    $holiday = [DateTime]::ParseExact($day, "M/d/yyyy", $null)
    $holiday_list += $holiday
}

# Generate calendar days
$days_of_month = @{}
$current_day = Get-Date -Year $YEAR -Month $MONTH -Day 1
while ($current_day.Month -eq $MONTH) {
    $is_weekday = $current_day.DayOfWeek -in @([DayOfWeek]::Monday..[DayOfWeek]::Friday)
    $is_holiday = $holiday_list -contains $current_day.Date
    $slots = @('NONE', 'NONE', 'NONE', 'NONE')
    if ($is_holiday) {
        $slots = @('HOLIDAY', 'HOLIDAY', 'HOLIDAY', 'HOLIDAY')
    }
    
    $days_of_month[$current_day] = @{
        'IsWeekday' = $is_weekday
        'IsHoliday' = $is_holiday
        'Slots' = $slots
    }
    $current_day = $current_day.AddDays(1)
}

function Find-Fit {
    param(
        [string]$name,
        [hashtable]$days_of_month,
        [hashtable]$p_info
    )
    
    foreach ($day in ($days_of_month.Keys | Sort-Object)) {
        if (-not $days_of_month[$day].IsWeekday -or $days_of_month[$day].IsHoliday) {
            continue
        }
        if ($days_of_month[$day].Slots -notcontains 'NONE') {
            continue
        }
        
        $day_of_week = $day.DayOfWeek.ToString()
        $works_today = $p_info[$name][$day_of_week]
        # Check slots based on timing and role
        if ($days_of_month[$day].Slots[0] -eq 'NONE' -and -not $p_info[$name].OpsOK) {
            if ($p_info[$name].Start -le $morning_start -and $works_today) {
                $days_of_month[$day].Slots[0] = $name
                return $true
            }
        }
        if ($days_of_month[$day].Slots[1] -eq 'NONE' -and $p_info[$name].OpsOK) {
            if ($p_info[$name].Start -le $morning_start -and $works_today) {
                $days_of_month[$day].Slots[1] = $name
                return $true
            }
        }
        if ($days_of_month[$day].Slots[2] -eq 'NONE' -and -not $p_info[$name].OpsOK) {
            if ($p_info[$name].End -ge $afternoon_finish -and $works_today) {
                $days_of_month[$day].Slots[2] = $name
                return $true
            }
        }
        if ($days_of_month[$day].Slots[3] -eq 'NONE' -and $p_info[$name].OpsOK) {
            if ($p_info[$name].End -ge $afternoon_finish -and $works_today) {
                $days_of_month[$day].Slots[3] = $name
                return $true
            }
        }
    }
    return $false
}

function Test-CalendarDone {
    param(
        [hashtable]$days_of_month,
        [int]$role = 3
    )
    
    if ($role -eq 3) {
        foreach ($day in $days_of_month.Keys) {
            if (-not $days_of_month[$day].IsWeekday -or $days_of_month[$day].IsHoliday) {
                continue
            }
            if ($days_of_month[$day].Slots -contains 'NONE') {
                return $false
            }
        }
    }
    elseif ($role -eq 2) {
        foreach ($day in $days_of_month.Keys) {
            if (-not $days_of_month[$day].IsWeekday -or $days_of_month[$day].IsHoliday) {
                continue
            }
            if ($days_of_month[$day].Slots[1] -eq 'NONE' -or $days_of_month[$day].Slots[3] -eq 'NONE') {
                return $false
            }
        }
    }
    elseif ($role -eq 1) {
        foreach ($day in $days_of_month.Keys) {
            if (-not $days_of_month[$day].IsWeekday -or $days_of_month[$day].IsHoliday) {
                continue
            }
            if ($days_of_month[$day].Slots[0] -eq 'NONE' -or $days_of_month[$day].Slots[2] -eq 'NONE') {
                return $false
            }
        }
    }
    return $true
}

# Start fitting from force_in list
$already_in = @()
while ($force_in.Count -gt 0) {
    $name = $force_in[-1]
    if ($force_in.Count -eq 1) {
        $force_in = @()
    }
    else {
        $force_in = $force_in[0..($force_in.Count-2)]
    }
    $fit_found = Find-Fit -name $name -days_of_month $days_of_month -p_info $workers
    if ($fit_found) {
        $already_in += $name
    }
    else {
        Write-Host "Warning: could not find a fit for $name from the ForceIn list"
    }
}

# Process those that didn't go last time (duty only)
$try_list = @()
foreach ($person in $workers.Keys) {
    if ($person -notin $last_month -and 
        $person -notin $already_in -and 
        $person -notin $skip_list -and 
        -not $workers[$person].OpsOK) {
        $try_list += $person
    }
}
$try_list = $try_list | Sort-Object {Get-Random}
$done = $false
while ($try_list.Count -gt 0 -and -not $done) {
    $name = $try_list[-1]
    if ($try_list.Count -eq 1) {
        $try_list = @()
    }
    else {
        $try_list = $try_list[0..($try_list.Count-2)]
    }
    $fit_found = Find-Fit -name $name -days_of_month $days_of_month -p_info $workers
    if ($fit_found) {
        $already_in += $name
    }
    elseif (Test-CalendarDone -days_of_month $days_of_month -role 1) {
        $done = $true
        break
    }
    else {
        Write-Host "Note: could not find a fit for $name from the not-in-last-month duty list"
    }
}

# Fill remaining duty slots
$try_list = @()
foreach ($person in $workers.Keys) {
    if ($person -notin $already_in -and 
        -not $workers[$person].OpsOK -and 
        $person -notin $skip_list) {
        $try_list += $person
    }
}
$try_list = $try_list | Sort-Object {Get-Random}
while ($try_list.Count -gt 0 -and -not $done) {
    $name = $try_list[-1]
    if ($try_list.Count -eq 1) {
        $try_list = @()
    }
    else {
        $try_list = $try_list[0..($try_list.Count-2)]
    }
    $fit_found = Find-Fit -name $name -days_of_month $days_of_month -p_info $workers
    if ($fit_found) {
        $already_in += $name
    }
    elseif (Test-CalendarDone -days_of_month $days_of_month -role 1) {
        $done = $true
        if ($DEBUG) {
            Write-Host "Finished duty schedule with $($try_list.Count) people still in the list"
        }
        break
    }
    else {
        Write-Host "Note: could not find a fit for $name from the full list for duty schedule"
    }
}

if (-not $done) {
    Write-Host "Error: Was not able to fill the duty schedule"
}

# Fill ops schedule
$done = $false
$loop_count = 0
while (-not $done) {
    if ($loop_count -gt 100) {
        Write-Host "Error: seem not completing ops schedule"
        break
    }
    $try_list = @()
    foreach ($person in $workers.Keys) {
        if ($person -notin $already_in -and 
            $workers[$person].OpsOK -and 
            $person -notin $skip_list) {
            $try_list += $person
        }
    }
    $try_list = $try_list | Sort-Object {Get-Random}
    while ($try_list.Count -gt 0 -and -not $done) {
        $name = $try_list[-1]
        if ($try_list.Count -eq 1) {
            $try_list = @()
        }
        else {
            $try_list = $try_list[0..($try_list.Count-2)]
        }
        $fit_found = Find-Fit -name $name -days_of_month $days_of_month -p_info $workers
        if (Test-CalendarDone -days_of_month $days_of_month -role 2) {
            $done = $true
            break
        }
    }
    $loop_count++
}

# Final check and output
if (-not (Test-CalendarDone -days_of_month $days_of_month)) {
    Write-Host "Error: failed final check that calendar is complete"
}

# Output the flat schedule
$output = @()
foreach ($day in ($days_of_month.Keys | Sort-Object)) {
    $row = [PSCustomObject]@{
        'Date' = $day.ToString("MM/dd/yyyy")
        'IsWeekday' = [int]$days_of_month[$day].IsWeekday
        'IsHoliday' = [int]$days_of_month[$day].IsHoliday
        'EarlyDutyInvestigator' = $days_of_month[$day].Slots[0]
        'EarlyOperationsOfficer' = $days_of_month[$day].Slots[1]
        'LateDutyInvestigator' = $days_of_month[$day].Slots[2]
        'LateOperationsOfficer' = $days_of_month[$day].Slots[3]
    }
    $output += $row
}

$output | Export-Csv -Path $schedule_out -NoTypeInformation
