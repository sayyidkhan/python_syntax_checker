# python_syntax_checker
We use this application to mark the students work.

- checks your html criteria, if a set of element exist in a html file
- can be used as a marking scheme
- currently it checks 17 tags by default
- update the testcase variable to customise the testing specific for your marking scheme

## pre-requisite:
must install
- python 3.7 and above
- openpyxl (pip3 install openpyxl)

before even running the program!!

## how to use


1. put all the students work in the same directory ensure that html, css and js is all in one html file
   - when the student submit the work, the filename should by their name
   - update the file name in the excel spreadsheet
3. run the python checker.py
   ```
   python3 checker.py
   ```
3. open the spreadsheet to view the results of the student

## how to modify test cases

proceed to the testcases variable and add below from the last element in the list.
thats it

** test result's will be overwritten on each run, make a copy of the file should u not want it to be over-written **
** clear the student_list.xlsx after marking, delete all values in all cell. do not delete the file. the program will not work **
