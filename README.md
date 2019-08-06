# VBA ─ Overview

VBA stands for `Visual Basic for Applications` an event-driven programming language from
Microsoft that is now predominantly used with Microsoft office applications such as MS-Excel,
MS-Word, and MS-Access.

In Excel window, press `ALT+F11`. 

```java
Private Sub say_helloWorld()
    MsgBox "Hi"
End Sub
```

## Comments

* Keyword `REM` and `'`.

```java
' Written by : Jeevan Lal

REM Modified by : Jeevan Lal
```

## Variables/Constants

**Syntax**

> Dim `variable_name` As `variable_type`

> Const `constant_name` As `constant_type` = `constant_value`

### Data Types

* Byte
* Integer
* Long
* Single
* Double
* Currency
* Decimal
* String
* Date
* Boolean
* Object
* Variant

Example : 

```java
Private Sub Variables()

    Dim name As String
    name = "jeevan"

    Const password As Integer = 1234

    Dim birthDay As Date
    birthDay = 30 / 10 / 2020

    MsgBox "Password is " & password & Chr(10) & "Name " & name & Chr(10) & "Birthday is " & birthDay

End Sub
```

```java
Private Sub AnotherMethod()

    Dim a as String : a = "jeevan lal"
    Dim score As Integer, result As String
    
End Sub
```

## If Then/Else Statement

```java
Dim score As Integer, result As String
score = Range("A1").Value

If score >= 60 Then result = "pass"

Range("B1").Value = result
```

```java
Dim score As Integer, result As String
score = Range("A1").Value

If score >= 60 Then
    result = "pass"
Else
    result = "fail"
End If

Range("B1").Value = result
```

## Workbook and Worksheet Object

```java
Range("A1").Value = "Hello"
```

but what we really meant was:

```java
Application.Workbooks("create-a-macro").Worksheets(1).Range("A1").Value = "Hello"
```

### Collections

* Using the worksheet name

```java
Worksheets("Sales").Range("A1").Value = "Hello"
```

*  Using the index number (1 is the first worksheet starting from the left).

```java
Worksheets(1).Range("A1").Value = "Hello"
```

* Using the CodeName.

```java
Sheet1.Range("A1").Value = "Hello"
```

### Properties and Methods

1. The Add method of the Workbooks collection creates a new workbook.

```
Workbooks.Add
```

2. The Count property of the Worksheets collection counts the number of worksheets in a workbook.

```
MsgBox Worksheets.Count
```

## Range Object

```java
Range("B3").Value = 2

Range("A1:A4").Value = 5

Range("A1:A2,B3:C4").Value = 10
```

### Named Range

[Web URL](https://www.excel-easy.com/examples/names-in-formulas.html)

Example :

```java
Range("Prices").Value = 15
```

### Cells

```java
Cells(3, 2).Value = 2
```

Excel VBA enters the value 2 into the cell at the intersection of row 3 and column 2.

```java
Range(Cells(1, 1), Cells(4, 1)).Value = 5
```

### Declare a Range Object

You can declare a Range object by using the keywords `Dim` and `Set`.

```java
Dim example As Range
Set example = Range("A1:C4")

example.Value = 8
```

### Select

An important method of the Range object is the Select method. The Select method simply selects a range.

```java
Dim example As Range
Set example = Range("A1:C4")

example.Select
```

`Note`: To select cells on a different worksheet.

```java
Worksheets(3).Activate
Worksheets(3).Range("B7").Select
```

### Rows/Columns

The Rows property gives access to a specific row of a range. The Columns property gives access to a specific column of a range.

```java
Dim example As Range
Set example = Range("A1:C4")

example.Rows(3).Select
```

```java
Dim example As Range
Set example = Range("A1:C4")

example.Columns(2).Select
```

### Copy/Paste

The Copy and Paste method are used to copy a range and to paste it somewhere else on the worksheet.

```java
Range("A1:A2").Select
Selection.Copy

Range("C3").Select
ActiveSheet.Paste
```

```java
Range("C3:C4").Value = Range("A1:A2").Value
```

### Clear

To clear the content of an Excel range, you can use the ClearContents method.

```java
Range("A1").ClearContents

Range("A1").Value = ""
```

### Count

With the Count property, you can count the number of cells, rows and columns of a range.

```java
Dim example As Range
Set example = Range("A1:C4")

MsgBox example.Count
MsgBox example.Rows.Count
```

## Loop

### Single Loop

```java
Dim i As Integer

For i = 1 To 6
    Cells(i, 1).Value = 100
Next i
```

### Double Loop

```java
Dim i As Integer, j As Integer

For i = 1 To 6
    For j = 1 To 2
        Cells(i, j).Value = 100
    Next j
Next i
```

### Triple Loop

```java
Dim c As Integer, i As Integer, j As Integer

For c = 1 To 3
    For i = 1 To 6
        For j = 1 To 2
            Worksheets(c).Cells(i, j).Value = 100
        Next j
    Next i
Next c
```

## Do While Loop

```java
Dim i As Integer
i = 1

Do While i < 6
    Cells(i, 1).Value = 20
    i = i + 1
Loop
```

```java
Dim i As Integer
i = 1

Do While Cells(i, 1).Value <> ""
    Cells(i, 2).Value = Cells(i, 1).Value + 10
    i = i + 1
Loop
```

**Explanation**: as long as `Cells(i, 1)`.Value is not empty `(<> means not equal to)`, Excel VBA enters the value into the cell at the intersection of row i and column 2, that is 10 higher than the value in the cell at the intersection of row i and column 1. Excel VBA stops when i equals 7 because Cells(7, 1).Value is empty. This is a great way to loop through any number of rows on a worksheet.


## String Manipulation

### Join Strings

```java
Dim text1 As String, text2 As String
text1 = "Hi"
text2 = "Tim"

MsgBox text1 & " " & text2
```

`Note`: to insert a space, use " "

### Left/Right/Mid/Len/Instr

`Note`: To find the position of a substring in a string, use `Instr`.

```java
MsgBox Left("example text", 4)
MsgBox Right("example text", 2)
MsgBox Mid("example text", 9, 2)
MsgBox Len("example text")

' Note: string "am" found at position 3.
MsgBox Instr("example text", "am") 
```

## Date and Time

To get the current date and time, use the Now function.

```java
MsgBox Now
```

### Year, Month, Day of a Date, Hour, Minute, Second

```java
Dim exampleDate As Date

exampleDate = DateValue("Jun 19, 2010")

MsgBox Year(exampleDate)
MsgBox Hour(Now)
```

### DateAdd

```java
Dim firstDate As Date, secondDate As Date

firstDate = DateValue("Jun 19, 2010")
secondDate = DateAdd("d", 3, firstDate)

MsgBox secondDate
```

### TimeValue

The TimeValue function converts a string to a time serial number. The time's serial number is a number between 0 and 1. For example, noon (halfway through the day) is represented as 0.5.

```java
MsgBox TimeValue("9:20:01 am")
```

## Target

```java
Target.Address
Target.Value
```

## Array

### One-dimensional Array

```java
Dim Films(1 To 5) As String

Films(1) = "Lord of the Rings"
Films(2) = "Speed"
Films(3) = "Star Wars"
Films(4) = "The Godfather"
Films(5) = "Pulp Fiction"

MsgBox Films(4)
```

### Two-dimensional Array

```java
Dim Films(1 To 5, 1 To 2) As String
Dim i As Integer, j As Integer

For i = 1 To 5
    For j = 1 To 2
        Films(i, j) = Cells(i, j).Value
    Next j
Next i

MsgBox Films(4, 2)
```

## Function and Sub

The difference between a function and a sub in Excel VBA is that a function can return a value while a sub cannot.

### Function

```java
Function Area(x As Double, y As Double) As Double

    Area = x * y

End Function
```

**Using Function**

```java
Dim z As Double

z = Area(3, 5) + 2

MsgBox z
```

### Sub

```java
Sub Area(x As Double, y As Double)

    MsgBox x * y

End Sub
```

**Using Function**

```java
Area 3, 5
```

## Application Object

The mother of all objects is Excel itself. We call it the `Application object`. The application object gives access to a lot of Excel related options.

* WorksheetFunction
* ScreenUpdating
* DisplayAlerts
* Calculation