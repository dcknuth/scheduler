# Schedule It
Take some csv files from Excel and create a schedule using it. Output the result to a csv to be read into Excel and put in a nice calendar format.

## Overall usage instructions
1. Get the [sample Office 365 Excel file](https://1drv.ms/x/c/9f09d57232dd8748/ESoklATP_r9GkHnsAf8lpF0BRrcTuGRoCxhuPNuaFxI85Q?e=MbN3Nk)
2. Change cell B1 on the Input sheet to the start of the month you would like to schedule. Also change B2-B5 in that sheet for some of the automatic checking features to work
3. Check that the holiday list in the HolidayList sheet is correct. You should only have to do this once a year or if a holiday changes on you
4. **Most of your work** will be to update the OfficerList sheet to have the people you are filling the lists with and their schedules (columns A-I). This part you only need to do once or as changes are needed
5. Update the **SkipList** column to whoever you need to skip this month. Use **exactly** what you put for that person in column A (still in the OfficerList sheet)
6. (Optional) Check the **InRotation** column to see if there is no match with either of the slots that need to be scheduled (still in the OfficerList sheet)
7. Update the **WentLastMonth** column to whoever had a shift last month (you will need to pull this out of the previous months schedule yourself). The script will then start with those that were not in the schedule last month. You can also just leave this blank if you don't care about an end-of-month person getting a beginning-of-month slot or a general push toward fairness for large lists of people
8. Update the **ForceInFirst** column if you want to try to fit a few people in first. Useful if the schedule was wrong and a person has not been picked for a long time, starting a new person or other fairness concerns. This column can be blank and should be reasonably short to get the intended effect
9. Export the **HolidayList** sheet to a CSV file named `holidays.csv` and put it in the same location as the script. You only need to do this once a year if there are no changes to holidays
10. Export the **OfficerList** sheet to a CSV file named `availability.csv` and put it in the same location as the script. You only need to do this a first time and then each time your team's schedule has a change
11. Run the script as noted in the next section. It will be something like:  
```python schedule_it.py 2025 3```  
OR  
```powershell schedule_it.ps1 2025 3```  
12. Open the output CSV `flat_sched.csv` and copy and paste it into the **FlatSched** sheet in Excel
13. The **Schedule** sheet should now have a nicely formatted schedule for the month you wanted. I recommend you copy this and paste values into a shared sheet for your team with edit history turned on. Then you can tell them to do trades on their own and look into the edit history to resolve any disputes. Best of all, they can do this without you and not mess up your scheduling sheet
14. The **SpecialCases** sheet is only to keep notes about what you did and why, so when someone asks you remember the answer

### If there is a problem:
* Check that the month in B1 of the **Input** sheet is the same as you passed into the script. If it is not, the schedule will look weird for the month
* Follow up on any errors the script generated
* Look at the InRotation column of the **OfficerList** sheet. If the people you expect to get scheduled are not, look at their work week setup
* Run the script with more information output to the command line with the VERBOSE flag. That would look like this:  
```python schedule_it.py 2025 3 availability.csv holidays.csv VERBOSE```  
Yes, you need to also pass the file names to use the VERBOSE flag
* Last, look at the code. For instance, the morning shift's start time and the afternoon shift's end time are hard coded and might need to be updated for your case (they do not currently come from the input sheet). I also may have a logic error or other nice items needed. Feel free to make a suggestion or pull request.

## Usage of the command line tool
**schedule_it.py YEAR_NUM MONTH_NUM [availability.csv] [holidays.csv] [VERBOSE]**

Arguments have to be in order with at least a year (as four digits) and a month (as an integer). The files are assumed as listed and must be specified if listing VERBOSE at the end to get additional output

There is a matching Excel sheet that has inputs that need to match those passed in on the command line (the year and month) or the output CSV will not work with the Excel formulas

## Assumptions
* Expecting the list to "force in" to the rotation to be fairly short
* Expecting two morning slots that need to be filled and two afternoon slots that need to be filled
* The two types of slots are filled with the **OpsOK** column and people cannot serve in both slots, just one or the other

## Requirements
* Standard Python only - This is going into an environment that **SHOULD** be locked down. The Windows App store has a Python 3.11 that should be available to even a non-admin user. I don't know if pip will run, so will assume it is fire-walled off and non-functional, so no Pandas or non-included modules
* It ends up that Python could not be installed, so ported to PowerShell also
* Support a flexible work week where any day could be a working or non-working day for any person in the rotation
* Early and late shifts - Support a morning and an afternoon shift.
* Support flexible work hours - Days can start after the morning shift and end before the afternoon shift ends. We only want to schedule convenient shifts
* Support two roles that need to be scheduled, from the **OpsOK** column
* Support a list of people to skip this month
* Support a list to be scheduled first with very high probability of getting scheduled. A "force in" list via the **ForceInFirst** column
* Support a list of who was scheduled last time and put those at the end of the scheduling process for lower probability of getting scheduled (except if they are in the "force in" list)

## Out of Scope
* No support of different hours each day. If needed, use the latest start and earliest end times or pick a single time bracket and mark as not working for other days
* No fancy calendar output from Python(or PowerShell), we will let Excel handle that part
* Running fast. It just has to book out a single month with lists that are under 200 long. Go for easy to understand vs. highly optimized

## For Later
* Move to a configuration CSV with more flexability that sets things like how many slots need to be filled and when they start/end
* Move the fit tests into another file and import in
* Add more input testing and warnings/errors including mis-matches between the lists of people (there is some)
* Include more about final status and what was tried for outside debugging or Excel sheet updates (again, there is a bit currently)
* Add to use names a second time if needed for the non-ops role. For the moment, it should finish using each name just once and the ops role will refill if needed
