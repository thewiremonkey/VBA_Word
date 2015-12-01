Attribute VB_Name = "Cal"
Option Explicit
Sub calendar()
Application.ScreenUpdating = False
Dim i, j, k, m, mStart, mEnd, yStart, yEnd, mCount, d, n, yy, x As Long
Dim stStart, stEnd, stCount As String
Dim r As Range
Dim tn As Column
Dim tc, tr  As Long
Dim t As Table
Dim oDoc As Document
Dim start, FINISH

start = Timer


'----------Initialize variables---------'
n = 0
Set oDoc = ActiveDocument


stStart = Format(InputBox("start mm/yyyy "), "mm/yyyy") 'get start month
stCount = CLng(InputBox("how many months? (1-6)"))



mStart = Month(DateValue(stStart))
yStart = Year(DateValue(stStart))

mEnd = mStart + stCount 'as many months as indicated

mCount = mEnd - mStart

x = mStart
'-----------------------------------'
Set t = oDoc.Tables.Add(Range:=Selection.Range, NumRows:=6, _
NumColumns:=(mCount * 7), AutoFitBehavior:=False) 'create base table

With t
.PreferredWidth = 434
.Columns.Width = InchesToPoints(0.22)
.Rows.Height = InchesToPoints(0.22)
.Rows.Alignment = wdAlignRowCenter
.Range.ParagraphFormat.SpaceAfter = 0
.Range.ParagraphFormat.SpaceBefore = 0
.Range.Font.Name = "times new roman"
.Range.Font.Size = 9
.RightPadding = 0
.LeftPadding = 0
.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth025pt
            .Color = wdColorAutomatic
        End With

End With

Selection.ExtendMode = False

    tc = t.Columns.Count
    
    tr = t.Rows.Count
    
    For k = mStart To (mStart + mCount) - 1
    n = n + 1
            If k < 13 Then
            
            yy = yStart
            mStart = k
            
            
            Else:
            yy = yStart + 1
            mStart = k - 12
           
            End If
    
    d = Weekday((mStart) & "/1/" & yy) - 1
    t.Columns((7 * n) - 6).Select
    

    
    With Selection
    .MoveRight unit:=wdCharacter, Count:=6, Extend:=wdExtend
    .Bookmarks.Add ("t" & k) 'add a bookmark, name it
    End With
    Selection.ExtendMode = False
    oDoc.Bookmarks("t" & k).Select
    With Selection
            
            
            For j = 1 To (DateSerial(yy, mStart + 1, 0) - _
                        DateSerial(yy, mStart, 0)) 'find the number of days in the month

        
        oDoc.Bookmarks("t" & k).Select
        Selection.Cells(d + j).Range.Text = j
        If Weekday(d + j) = 1 Or Weekday(d + j) = 7 Then
        Selection.Cells.Shading.BackgroundPatternColor = wdColorGray10
        End If
        
        
        Next j
        
    End With
   
    Next k
   t.Rows.Add beforerow:=t.Rows(1)
   
   For k = 1 To mCount
   For m = 1 To 7
   
   With t.Cell(1, (k * 7) - 7 + m).Range
   .Text = WeekdayName(m, True)
   .Font.Size = 7
   .ParagraphFormat.Alignment = wdAlignParagraphCenter
   .Cells.VerticalAlignment = wdCellAlignVerticalCenter
   .Cells.Shading.BackgroundPatternColor = -587137114
   
   End With
   Next m
   Next k
   t.AutoFitBehavior wdAutoFitFixed
    For k = 1 To mCount - 1
       t.Columns.Add beforecolumn:=t.Columns((k * 7) + k)
    Next k
    t.AutoFitBehavior wdAutoFitFixed
    
    For k = 1 To mCount - 1
   t.Columns(k * 8).PreferredWidth = 7
   t.Columns(k * 8).Shading.BackgroundPatternColor = wdColorWhite
   t.Columns(k * 8).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
   Next k
   
t.Rows.Add beforerow:=t.Rows(1)
   t.Cell(1, 1).Select
   
For k = x To (x + mCount) - 1
    If k = x Then
    oDoc.Bookmarks("t" & k).Range.Select
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveUp unit:=wdLine, Count:=2
    Else
    If k > x Then
    oDoc.Bookmarks("t" & k).Range.Select
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveRight unit:=wdCell, Count:=k - x
    Selection.MoveUp unit:=wdLine, Count:=2
    'Selection.HomeKey Unit:=wdColumn
    End If
    End If

   Selection.ExtendMode = True
   
   Selection.MoveRight unit:=wdCharacter, Count:=(7), Extend:=wdExtend
   Selection.Cells.Merge
   Selection.Range.Font.Size = 10
   Selection.Range.Font.Bold = True
   Selection.Range.Font.Color = -587137114
   Selection.Range.Cells.Shading.BackgroundPatternColor = wdColorWhite
      
If k >= 13 Then
Selection.Range.Text = MonthName(k - 12) & " " & yStart + 1
Else
If k < 13 Then
Selection.Range.Text = MonthName(k) & " " & yStart
End If
End If

   Selection.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
   Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
   Selection.ExtendMode = False
   
   Selection.MoveRight unit:=wdCell, Count:=3
    
   Next k
  
   
   For i = 1 To t.Rows(2).Cells.Count
   
   If t.Rows(2).Cells(i).Range.Text = Chr(13) & Chr(7) Then GoTo RESUMEK
   
   't.Rows(2).Cells(i).Shading.BackgroundPatternColor = wdColorGray90
   t.Rows(2).Cells(i).Range.Font.ColorIndex = wdWhite
   
RESUMEK:
   Next i


FINISH = Timer
Debug.Print FINISH - start
If stCount = 1 Then GoTo FORMATSINGLE Else

GoTo FINISH


FORMATSINGLE:
    'Selection.tables(1).Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).Cell(3, 1).Select
    Selection.Tables(1).Rows.Height = Selection.Tables(1).Cell(3, 1).Width
    
 
  
    
FINISH:

    Selection.EndKey unit:=wdColumn
    Selection.EndKey unit:=wdRow
    Selection.MoveDown unit:=wdLine, Count:=1
    Selection.Text = Chr(13)
    Selection.MoveDown unit:=wdLine, Count:=1

Application.ScreenUpdating = True


End Sub
Sub CalShade()
Attribute CalShade.VB_Description = "Macro recorded 11/3/2011 by Document Production User"
Attribute CalShade.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.CalShade"
'
' CalShade Macro
' Macro recorded 11/3/2011 by Richard Fitch/Document Production/x8896
'
If Selection.Information(wdWithInTable) = False Then
    MsgBox "To shade a cell, please place the cursor inside the chosen calendar cell or select a range of cells.", vbOK + vbInformation, "Calendar Shading"

GoTo Terminate
Else

    With Selection.Cells
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorGray10
        End With
    End With
End If
Terminate:

End Sub
Sub CalRemoveShade()
Attribute CalRemoveShade.VB_Description = "Macro recorded 11/3/2011 by Document Production User"
Attribute CalRemoveShade.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.CalRemoveShade"
'
' CalRemoveShade Macro
' Macro recorded 11/3/2011 by Richard Fitch/Document Production/x8896
'
If Selection.Information(wdWithInTable) = False Then
    MsgBox "To remove shading from a cell, please place the cursor inside the chosen calendar cell or select a range of cells.", vbOK + vbInformation, "Calendar Shading"

GoTo Terminate
Else
    
    
    With Selection.Cells
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
    End With
End If
Terminate:

End Sub
Sub CalMarkDate()
'
' CalMarkDate Macro
' Macro recorded 11/3/2011 by Richard Fitch/Document Production/x8896
'
    
If Selection.Information(wdWithInTable) = False Then
    MsgBox "To emphasize a date by framing it, please place the cursor inside the chosen calendar cell.", vbOK + vbInformation, "Calendar Shading"

GoTo Terminate
Else
    
    With Selection.Cells
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End If
Terminate:

End Sub
Sub CalRemoveDateMark()
'
' CalRemoveDateMark Macro
' Macro recorded 11/3/2011 by Richard Fitch/Document Production/x8896
'

If Selection.Information(wdWithInTable) = False Then
    MsgBox "To remove a mark from a cell, please place the cursor inside the marked calendar cell.", vbOK + vbInformation, "Calendar Shading"

GoTo Terminate
Else
    With Selection.Cells
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End If
Terminate:

End Sub

