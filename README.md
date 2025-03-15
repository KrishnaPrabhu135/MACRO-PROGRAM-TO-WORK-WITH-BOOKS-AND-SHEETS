# MACRO-PROGRAM-TO-WORK-WITH-BOOKS-AND-SHEETS
SRI KRISHNA ARTS AND SCIENCE COLLEGE
Coimbatore-641 008
RECORD NOTE
DEPARTMENT OF COMPUTER TECHNOLOGY AND DATA SCIENCE
NAME
REGISTER NUMBER
PROGRAMME :
CLASS :
COURSE :
SRI KRISHNA ARTS AND SCIENCE COLLEGE
Coimbatore-641 008
REGISTER NO:
Certified bonafide record of work done by ____________________________________________
during the year 2024 – 2025.
Staff In-charge Head of the Department
Submitted to the Sri Krishna Arts & Science College (Autonomous) end semester examination
held on____________________________________________.
Internal Examiner External Examiner
DECLARATION
I
___________________________________________________ hereby declare that this record of
observations is based on the experiments carried out and recorded by me during the laboratory
classes of “_____________________________________________ ” conducted by Sri Krishna Arts
and Science College, Coimbatore - 641 008.
Date : ___________________
Signature of the Student
Name of the Student :
Register Number :
___________________________
Countersigned by Staff
CONTENTS
S. No TITLE OF THE EXPERIMENTS Page
No
Sign
1
2
3
4
5
6
7
8
9
10
11
12
EX. NO: 01
MACRO PROGRAMS USING BUTTON AND MESSAGE BOX
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
(General)
Sub message ()
MsgBox “Welcome to SKASC”
End Sub
OUTPUT:
RESULT:
EX. NO: 02
DATE:
MACRO PROGRAM TO WORK WITH BOOKS AND SHEETS
USING LOOPS
AIM:
ALGORITHM:
SOURCE CODE:
Sub Button2_Click()
Dim book As Workbook, sheet As Worksheet, text As StringFor Each book In Workbooks
text = text & "Workbook: " & book.Name & vbNewLine & "Worksheets: "& vbNewLine
For Each sheet In book.Worksheets
text = text & sheet.Name & vbNewLineNext sheet
text = text & vbNewLine Next book
MsgBox text
End Sub
OUTPUT:
RESULT:
EX. NO: 03
MACRO PROGRAM TO FIND AREA OF SHAPES
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
(General)
Function rect (height As Double, width As Double
rect = width * height
End Function
Function sq(side As Double) As Double
sq = side * side
End Function
Function tri (side As Double, side2 As Double, side3 As Double) As Double
Dim p As Double
p = (side + side2 + side3) /2
tri = Sqr (p * (p – side1) * (p – side2) * (p – side3) )
End Function
Function cir (radius As Double) As Double
cir = 3.14159 * radius * radius
End Function
OUTPUT:
RESULT:
EX. NO: 04
PROGRAM TO PERFORM ARITHMETIC AND
LOGICAL OPERATIONS
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Private Sub CommandButton1_Click()
If (1 = 1) And (0 = 0) Then
MsgBox "AND evaluated to TRUE", vbOKOnly, "AND operator"
Else
MsgBox "AND evaluated to FALSE", vbOKOnly, "AND operator"
End If
End Sub
Private Sub CommandButton2_Click()
If (1 = 1) Or (5 = 0) Then
MsgBox "OR evaluated to TRUE", vbOKOnly, "OR operator"
Else
MsgBox "OR evaluated to FALSE", vbOKOnly, "OR operator"
End If
End Sub
Private Sub CommandButton3_Click()
If Not (0 = 0) Then
MsgBox "NOT evaluated to TRUE", vbOKOnly, "NOT operator"
Else
MsgBox "NOT evaluated to FALSE", vbOKOnly, "NOT operator"
End If
End Sub
OUTPUT:
RESULT:
EX. NO: 05
MACRO PROGRAM TO CALCULATE VARIATION AND
STANDARD DEVIATION
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Sub varandstd()
Dim dataRange As Range
Dim variation As Double
Dim stdDev As Double
Set dataRange = ActiveSheet.Range("A2:A100")
variation = WorksheetFunction.Var(dataRange)
stdDev = WorksheetFunction.StDev(dataRange)
MsgBox "Variation: " & variation & vbCrLf & "Standard Deviation: " & stdDev
End Sub
OUTPUT:
RESULT:
EX. NO: 06
MACRO PROGRAM TO PERFORM DATE AND TIME IN EXCEL VBA
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Dim exampleDate As Date
exampleDate = DateValue("Jan 19, 2020")
MsgBox Year(exampleDate)
OUTPUT:
RESULT:
EX. NO: 07
MACRO PROGRAM TO GENERATE SALES CALCULATOR
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Sub CalculateTotalSales()
Dim employee As String, total As Double, sheet As Worksheet, i AsInteger
total = 0
employee = InputBox("Enter the employee name (case sensitive)")
For Each sheet In Worksheets
For i = 2 To 13 If sheet.Cells(i, 2).Value = employee Then
total = total + sheet.Cells(i, 3).Value
End If
Next i
Next sheet
MsgBox "Total sales of " & employee & " is " & total
End Sub
OUTPUT:
RESULT:
EX. NO: 08
EXCEL MACRO PROGRAM TO PREPARE CHARTS FOR RESULT
ANALYSIS
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Sub ModifyCharts()
Dim cht As ChartObject
For Each cht In Worksheets(1).ChartObjects
cht.Chart.ChartType = xlPie
Next cht Worksheets(1).ChartObjects(1).Activate
ActiveChart.ChartTitle.Text = "Sales Report"
ActiveChart.Legend.Position = xlLegendPositionBottom
ThisWorkbook.Save
End Sub
OUTPUT:
RESULT:
EX. NO: 09
EXCEL MACRO PROGRAM TO PERFORM STRING
MANIPULATION
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Sub ExecuteCode()
Dim text1 As String, text2 As Stringtext1 = "Hi"
text2 = "Tim"
MsgBox text1 & " " & text2
Dim text As String text = "example text"
MsgBox Left(text, 4)
MsgBox Right("example text", 2)
MsgBox Mid("example text", 9, 2)
MsgBox Len("example text")
MsgBox InStr("example text", "am")
End Sub
OUTPUT:
RESULT:
EX. NO: 10
EXCEL MACRO PROGRAM TO COUNT NUMBER OF WORDS IN A
GIVEN SENTENCE
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Sub CountWordsInSelectedRange()
Dim rng As Range, cell As Range
Dim cellWords As Integer, totalWords As Integer
Dim content As String
' Set the selected range
Set rng = Selection
totalWords = 0
' Loop through each cell in the range
For Each cell In rng
If Not cell.HasFormula Then
content = Trim(cell.Value)
If content = "" Then
cellWords = 0
Else
cellWords = 1
' Count spaces to determine number of words
Do While InStr(content, " ") > 0
content = Mid(content, InStr(content, " ") + 1)
content = Trim(content)
cellWords = cellWords + 1
Loop
End If
totalWords = totalWords + cellWords
End If
Next cell
' Display the total word count
MsgBox totalWords & " words found in the selected range.", vbInformation, "Word Count"
End Sub
OUTPUT:
RESULT:
EX. NO: 11
EXCEL MACRO PROGRAM PERFORM CREDIT POLICY TO FIND
MINIMUM AND MAXIMUM VALUES
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Sub AA()
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim D As Integer
Dim result As Integer
A = InputBox("Enter number 1")
B = InputBox("Enter number 2")
C = InputBox("Enter number 3")
D = InputBox("Enter number 4")
result = WorksheetFunction.Max(A, B, C, D)
MsgBox "The maximum number is: " & result
End Sub
OUTPUT:
RESULT:
EX. NO: 12
EXCEL MACRO PROGRAM TO PERFORM CASH FLOW ESTIMATION
DATE:
AIM:
ALGORITHM:
SOURCE CODE:
Sub EstimateCashFlow()
Dim ws As Worksheet
Dim lastRow As Long
Dim lastColumn As Long
Dim startColumn As Long
Dim totalInflows As Double
Dim totalOutflows As Double
Dim netCashFlow As Double
Dim col As Integer
' Set the worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")
' Find the last row in column A
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
' Find the last column in row 1
lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
' Start column (assuming first column contains descriptions or dates)
startColumn = 2
' Initialize totals
totalInflows = 0
totalOutflows = 0
' Loop through columns
For col = startColumn To lastColumn
If ws.Cells(1, col).Value Like "Inflow*" Then
totalInflows = totalInflows + WorksheetFunction.Sum(ws.Columns(col))
ElseIf ws.Cells(1, col).Value Like "Outflow*" Then
totalOutflows = totalOutflows + WorksheetFunction.Sum(ws.Columns(col))
End If
Next col
' Calculate net cash flow
netCashFlow = totalInflows - totalOutflows
' Display message box with results
MsgBox "Total Inflows: " & totalInflows & vbCrLf & _
"Total Outflows: " & totalOutflows & vbCrLf & _
"Net Cash Flow: " & netCashFlow, vbInformation, "Cash Flow Estimation"
End Sub
OUTPUT:
RESULT:
