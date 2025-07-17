Excel 2007: Macro-Based Data Entry Form (No VBA)

This project contains a form-based system for entering employee data in Excel 2007 using recorded macros only (no manual VBA coding). 
It includes dropdowns, radio buttons, input validation, and automatic storage into a database sheet.

Features : 
Form UI with formatting

Drop-down list for Department selection

Radio buttons for Gender selection

Auto-paste data to a database sheet

Button-triggered macro (recorded, not coded)

Clear and consistent input structure

Steps to Create This System: 

1. Create Form Layout:
   
Merge cells for heading and apply bold formatting.
Create labels like Name, DOB, Department, Gender, Location.
Apply border formatting to the input cells.

2. Apply Formatting:
   
For DOB → Format as Short Date .
For numeric fields (e.g., Salary) → Format Cells → Number.
Text fields → Format as General or Text as needed.

3. Add Department Drop-Down:
   
Select the input cell for Department.
Go to: Data → Data Validation.
Choose “List”.
Enter the list values (comma-separated), for example:
HR, Finance, IT, Marketing, Admin .

4. Add Gender Radio Buttons:
   
Go to: Developer → Insert → Option Button (Form Control).
Insert two buttons: "Male" and "Female" .
Right-click each → Format Control → Set both to same Cell Link (e.g., J3).
Use a helper formula (e.g., in K3):
=IF(J3=1, "Male", IF(J3=2, "Female", "")) 

5. Link Inputs to Helper Range:
   
Use formulas in a helper column to map form fields:

=B3      → Name
=B4      → DOB
=B5      → Department
=B6      → Location
=K3      → Gender (from formula above)

This creates a clean range to copy when saving to the database.

6. Prepare Database Sheet:
   
Add a second sheet named "Database".
Insert headers: Name, DOB, Department, Gender, Location, etc.
This will store all submitted entries row by row.

7. Record Macro to Save Data:
   
Go to Developer → Record Macro .
Name it SaveToDatabase .
While recording:

Copy the helper cell range.

Go to the Database sheet.

Find the next empty row.

Use Paste Special → Values (you can also use Transpose if vertical).

Return to the Form sheet.

Stop recording.

8. Add a Submit Button:
   
Go to: Developer → Insert → Button (Form Control).
Assign the recorded macro SaveToDatabase to it.
Rename the button (e.g., "Submit", "Add Entry").
Final Usage Flow.
Fill the form with all details.
Click the Submit button.

Data is saved to the next available row in the Database sheet.

Form remains ready for the next input.

