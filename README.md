# schneider_internship
I have completed a data analysis project while i was at schneider.

This project was done for the HR department of Schneider Electric Hyderbad in India.

Schneider specialises in digital automation and energy management.
Where I worked, it was a factory setup where employees have to punch in and punch out their RF Ids before they start to work.
Thus the system keeps track of the amount of time the employees have spent working at the factory, and stores it into an excel file with a name similar to SIM2to5Report (78).xlsx.
The people from HR take that RAW data and run excel operations on the file, to extract usefull info from it.
The end product is an excel file that shows the name and the dept of the employee and the no. of hours they have worked each week in that month.
And the cell with less than 10 hrs a week gets an orange highliht and the cell with 0 hrs gets a yellow highlight.

I have been told that it takes approx 30-40mins to complete this task manually which was the usual norm.
I did manage to automate this task in python using pandas, numpy and openpyxl libraries.
Openpyxl is a Python library that is used to read from an Excel file or write to an Excel file
I found this library to be very usefull, because once I was able to read data directly from the source 
the rest of the task to sort or anlyze the data was very easy using numpy and pandas.
