# Append Excel Table To SQL Already Existed table Data-pipeline using Python – 
The Usecase of Task – 

If our Sales-Person fill data in Excel sheet. It should dump Excel data to SQL Database for already existed table – 
Give input of Excel file Location and SQL Table Name.

Conditions to dump Excel Table Data to SQL Table –
1.	Columns Length should match. 
2.	Columns Name by Order should Match.
3.	Columns Name Data type should Match.

If all above 3 conditions matched it should append data in SQL.

Otherwise Give error message from below – 
1.	Table Columns Length not Matched.
2.	Column Name by Order Not Matched.
3.	Column DataType by Order Not Matched.

