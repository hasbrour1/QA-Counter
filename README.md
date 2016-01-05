# QA-Counter
Counts Analytes from each month and puts it in one excel table
Written in C#

The user can select each months excel file.  The QA-Counter will add each months analyte counts up and put them into a single excel file.

The main code is in Form1.cs.  I have switched from sorting the analytes using sequential sort to binary sort to increase run speed.  The program searches the list of analytes for the current counted analyte.  If it is not in the list it will add it, and if it is the anlyte count will be added to the current count stored. 
