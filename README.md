# Confirmation Sheets with Jupiter Attendance

Use this program to create confirmation cover sheets per teacher which includes student jupiter Attendance

## Files Needed
1. Folder of RDSCs for week of interest
2. Jupiter attendance export 

## Directions
1. Save the RDSC files as `.xlsx` The RDSCs files can keep their original naming 
2. Add RDCSs to a folder inside of `data`. Name the folder `Week_of_YYYY_MM_DD`
3. Download Jupiter Attendance for the week of interest. The columns should be 
    - Data -> Header
    - Student ID -> StudentID
    - Date (YYYY-MM-DD) -> Date
    - Attendance Mark -> Attendance
    - Course# -> Course
    - Section# -> Section
    - Period -> Period
4. Put this file inside the RDSC folder `attendance.csv`
5. Run `main.py` and when promoted for the week of, type `YYYY_MM_DD`