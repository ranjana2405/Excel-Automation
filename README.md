# Excel-Automation



Load Workbook and Select Sheet:

Load the existing Excel workbook sample.xlsx.

Select the active sheet within the workbook.

Read Cell Values:

Iterate through rows 1 to 5 and columns 1 to 3.

Print the values of each row to the console.

Write Data to Sheet:

Define a data list containing rows of data (Name, Age, City and specific values).

Append each row of data from the data list to the end of the sheet.

Apply Formatting to Header Row:

Define a bold white font (Font(bold=True, color="FFFFFF") and a solid black fill (PatternFill(start_color="000000", end_color="000000", fill_type="solid")).

Apply these formatting styles to all cells in the first row (header).

Calculate Average Age and Write to Sheet:

Extract ages from rows 2 to 4, column 2.

Compute the average of these ages.

Write "Average Age" to cell E1 and the calculated average age to cell E2.

Save Workbook:

Save the modified workbook back to sample.xlsx.

Print a success message indicating that data has been written and formatted successfully.

