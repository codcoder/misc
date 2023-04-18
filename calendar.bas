Attribute VB_Name = "Module1"
Sub CreateManyWeekCalendar()
    Dim numRepeats As Integer
    Dim weekNum As Integer
    Dim startDate As Date
    Dim i As Integer
    Dim j As Integer
    Dim dayOfWeek As Integer
    
    'Prompt user for number of times to repeat calendar
    numRepeats = InputBox("Enter number of times to repeat the calendar:", "Number of Repeats")
    
    'Prompt user for starting week number
    weekNum = InputBox("Enter a starting week number:", "Starting Week Number")
    
    For j = 1 To numRepeats
        'Calculate start date of the week (Monday)
        startDate = DateAdd("d", -Weekday(DateValue("1/1/" & Year(Date))) + 2 + (weekNum - 1) * 7, DateValue("1/1/" & Year(Date)))

        'Add week number to cell B1
        Range("B2").Offset(0, (j - 1) * 8).Value = "V" & weekNum
        Range("A3").Offset(0, (j - 1) * 8).Value = "Tid (start)"

        'Add time slots to column A
        For i = 0 To 11
            Range("A" & i + 4).Offset(0, (j - 1) * 8).Value = i + 7
            If i = 11 Then
                Range("A" & i + 4).Offset(0, (j - 1) * 8).Value = "18 (+)"
            End If
        Next i

        'Add days of the week to row 2
        For i = 1 To 7
            Range("A3").Offset(0, (j - 1) * 8 + i).Value = UCase(Left(Format(startDate, "ddd"), 1)) & LCase(Mid(Format(startDate, "ddd"), 2, 1))
            Range("A3").Offset(0, (j - 1) * 8 + i).HorizontalAlignment = xlCenter
            startDate = DateAdd("d", 1, startDate)
        Next i

        'Align text
        Range("A2:H2").Offset(0, (j - 1) * 8).HorizontalAlignment = xlCenter
        Range("A3:A13").Offset(0, (j - 1) * 8).HorizontalAlignment = xlRight
        Range("A14:A22").Offset(0, (j - 1) * 8).HorizontalAlignment = xlRight

        'Set column width
        Range("A1:A1").Offset(0, (j - 1) * 8).ColumnWidth = 8.86
        Range("B1:H1").Offset(0, (j - 1) * 8).ColumnWidth = 3.86

        'Set color
        Range("A1:A15").Offset(0, (j - 1) * 8).Interior.Color = RGB(242, 242, 242)
        Range("G1:H15").Offset(0, (j - 1) * 8).Interior.Color = RGB(242, 242, 242)
        
        'Increment week number
        weekNum = weekNum + 1
    Next j
End Sub
