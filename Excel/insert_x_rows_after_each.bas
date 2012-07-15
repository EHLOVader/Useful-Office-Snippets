'This script will ask you how many rows you want to insert, then go through the entire spreadsheet and after each row that contains data it will add X rows.
'For example:  You have a document with 200 rows, you then insert 6 into the dialoge that pops up, it will then go through the spreadsheet and add 6 rows after the first 199 rows.
Attribute VB_Name = "Module1"
Option Explicit
Sub InsertRows()
Dim rw, LastRow, inRows, numRows As Double
'Get Number of Rows to insert from User
 numRows = Application.InputBox("Enter The Number Of Rows To Insert")
'Find last row with data in Column A
 LastRow = Cells(65536, 1).End(xlUp).Row
'Loop through data, inserting rows
  For rw = LastRow To 2 Step -1
   For inRows = 1 To numRows
    Rows(rw).Insert Shift:=xlDown
   Next
  Next rw
End Sub
