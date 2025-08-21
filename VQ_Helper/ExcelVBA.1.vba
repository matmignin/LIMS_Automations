
'  --= Add_New_Sheet
Sub Add_New_Sheet()
    Dim sht                 As Worksheet
    Dim newsht              As Worksheet
    Dim VBA_BlankBidSheet   As Worksheet
    Dim newshtname          As String

    ' modify to your sheet's name
    Sheets("Template").Visible = xlSheetVisible
    Set VBA_BlankBidSheet = Sheets("Template")

    VBA_BlankBidSheet.Copy After:=Sheets("Main")
    Set newsht = ActiveSheet
    ActiveSheet.Tab.ColorIndex = 2
    Sheets("Template").Visible = False
    ' you can change it to your needs, or add an InputBox to select the Sheet's name
    newshtname = "ItemCode" ' InputBox("Item Code")

    newshtname = CustomInputBox("Enter Name for New Sheet", "new Sheet Name?")
    If StrPtr(newshtname) = 0 Then
    Application.ScreenUpdating = True
    Exit Sub
    End If
    If newshtname = NullString Then
    Application.ScreenUpdating = True
    Exit Sub
    End If
    ActiveSheet.Range("B1") = newshtname

    For Each sht In ThisWorkbook.Sheets
        If sht.Name = newshtname Then
            newshtname = newshtname & "_" & ThisWorkbook.Sheets.Count
            If InStr(UCase(Target), "ItemCode") > 0 Then
          '  newshtname = InputBox("Item Code")
            newsht.Name = newshtname
            End If
                Application.ScreenUpdating = True
            Exit Sub
              '  ActiveSheet.Range("B4") = ActiveSheet.Range("S3").Value
                End If
    Next sht
        newsht.Name = newshtname
        If InStr(UCase(Target), "ItemCode") > 0 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
End Sub




'  --= Paste_LabelCopy
 Sub Paste_LabelCopy()

    Dim myCell As Range
    Dim myPhrases() As String
    Dim i As Integer

If Application.ActiveSheet.Name = "Main" Or Application.ActiveSheet.Name = "Template" Or Application.ActiveSheet.Name = "SAMPLELOG" Then Exit Sub

On Error GoTo Exitthesub
   Application.Calculation = xlCalculationManual

ActiveCell.Select
 Range("CT7").Select  'switched from CT8 to CT7
 Selection.PasteSpecial xlPasteValues
 Application.CutCopyMode = False
Range("$CT7:$CY$100").Style = "IngredientAmounts"

 Range("$CT$7:$CY$80").UnMerge

    ' Define the range of cells to search
    Dim myRange As Range
    Set myRange = Range("CT12:CV80") 'Change this to your desired range

    ' Find the first occurrence of either "Supplement Facts" or "Nutrition Facts" in column CT and CU and CV
    Application.DisplayAlerts = False
    Dim startCell As Range
    Set startCell = Columns("CT").Find(What:="Supplement Facts", LookIn:=xlValues, LookAt:=xlPart)  'added more checks for startcell VVV
    If startCell Is Nothing Then Set startCell = Columns("CU").Find(What:="Supplement Facts", LookIn:=xlValues, LookAt:=xlPart)
        If startCell Is Nothing Then Set startCell = Columns("CU").Find(What:="Facts", LookIn:=xlValues, LookAt:=xlPart)

   ' If a startCell is found, determine the start row for myRange
    If Not startCell Is Nothing Then
        startRow = Application.Min(startCell.Row, myRange.Cells(1).Row)
        Set myRange = Range("CT" & startRow & ":CV80") 'changed CU to CT

        ' Get the last row that has a non-blank cell in column K that follows the startCell
        lastRow = startCell.Offset(2, -1).End(xlDown).Row
        Set myRange = Range(myRange.Cells(1), Cells(lastRow, "CV"))
    End If




    ' Clear the cells in columns CU-CY from lastRow+1 to the end of the worksheet
    'On Error GoTo Exitthesub
    Range("CU" & lastRow & ":CY" & Rows.Count).ClearContents

    ' Define the phrases to search for
    myPhrases = Split("per serving,per Container,Calories,Total Fat,Saturated Fat,Trans Fat,Cholesterol,Carbohydrate,Dietary Fiber,Total Sugars,Added Sugars,Daily Value,ingredients", ",")

    ' Loop through each cell in the range
    For Each myCell In myRange

        ' Loop through each phrase to search for
        For i = LBound(myPhrases) To UBound(myPhrases)

            ' If the cell contains the phrase, clear the cell and the two cells to the right
            If InStr(1, myCell.Value, myPhrases(i), vbTextCompare) > 0 Then
                ActiveSheet.Range(Cells(myCell.Row, myCell.Column), Cells(myCell.Row, myCell.Column + 5)).ClearContents
                i = LBound(myPhrases) 'Restart the loop from the beginning
                Exit For 'Exit the loop once a match is found
            End If

        Next i

    Next myCell
Call FindRange

 Application.CutCopyMode = False

    Call RemoveWhiteSpacefromB
 Call GetLabelCopyText

 Application.Calculation = xlCalculationAutomatic
     Application.DisplayAlerts = True
        Application.GoTo Range("B8"), Scroll:=True
     ActiveWindow.Zoom = 230
     ActiveSheet.Range("C8:C54").ShrinkToFit = True
     ActiveSheet.Range("C8:C54").Font.Color = RGB(255, 255, 255)
     Application.ScreenUpdating = True
Exitthesub:
'Application.CutCopyMode = False
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
 Exit Sub
End Sub


 Sub FindRange()
    Dim FirstCell As Range, LastCell As Range, CheckCell As Range

    'Find the first cell below "M8" that is not blank
    Set CheckCell = Range("CV9")
    Do Until CheckCell.Value <> ""
        Set CheckCell = CheckCell.Offset(1, 0)
    Loop
    Set FirstCell = CheckCell

    'Find the last non-blank cell in column L
    Set LastCell = Range("CU" & Rows.Count).End(xlUp)

    'Select the range between the first non-blank cell and the last non-blank cell
    Range(FirstCell, LastCell).Copy Range("B8")
End Sub


 Public Sub TrimWhiteSpacefromC()
Dim A As Range
Set A = Range("C8:C50")
  For Each cell In A
    cell.Replace What:="_ ", Replacemernt:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
Next
End Sub

 Public Sub RemoveWhiteSpacefromB()

    On Error GoTo Exitthesub
    Application.DisplayAlerts = False
  ActiveSheet.Range("B8:B50").Replace "  ", "|", xlPart
  ActiveSheet.Range("B8:B50").Replace "|", "", xlPart  '''new 07/03/24

Exitthesub:
Application.ScreenUpdating = True
Application.DisplayAlerts = True
 Exit Sub
End Sub




' --= Sub ReplaceDoubleSpaceLabelCopies()
Sub ReplaceDoubleSpaceLabelCopies()
    Dim dataRange As Range
    Dim delimiter As String

    ' Set the range where the data is located (B8:B50)
    Set dataRange = ActiveSheet.Range("B8:B50")

    ' Set the delimiter (two spaces)
    delimiter = "_ "
        dataRange.TextToColumns Destination:=dataRange, DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
            Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
            Other:=True, OtherChar:=delimiter, FieldInfo:=Array(1, 1)

            ActiveSheet.Range("C8:C50").Replace "m", " m", xlPart

End Sub


Sub SplitAndFormatTextShittyLabelCopyTexts()
    Dim ws As Worksheet
    Dim i As Long
    Dim text As String
    Dim splitIndex As Long
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
    ' Set the worksheet
    Set ws = ActiveSheet

    ' Loop through each cell in Column B from row 8 to row 54
    For i = 8 To 54
        text = ws.Cells(i, "B").Value
        If Len(text) > 0 Then
            ' Find the position of the last set of numbers in the text
            splitIndex = FindSplitIndex(text)

            ' Split the text
            If splitIndex > 0 Then
                ws.Cells(i, "B").Value = CleanUpText(Trim(Left(text, splitIndex - 1)))
                ws.Cells(i, "C").Value = CleanUpText(Trim(Mid(text, splitIndex)))
            Else
                ws.Cells(i, "B").Value = CleanUpText(text)
            End If
        End If
    Next i
    Call AddSpaceBetweenNumbersAndUnits
    Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
End Sub

Function FindSplitIndex(text As String) As Long
    Dim i As Long
    Dim lastDigitPos As Long

    ' Find the position of the last digit
    For i = Len(text) To 1 Step -1
        If IsNumeric(Mid(text, i, 1)) Then
            lastDigitPos = i
            Exit For
        End If
    Next i

    ' Find the position before the last set of digits starts
    For i = lastDigitPos To 1 Step -1
        If Not IsNumeric(Mid(text, i, 1)) And Mid(text, i, 1) <> "." Then
            FindSplitIndex = i + 1
            Exit Function
        End If
    Next i

    FindSplitIndex = 0 ' Return 0 if no suitable position found
End Function

Sub AddSpaceBetweenNumbersAndUnits()
    Dim ws As Worksheet
    Dim i As Long
    Dim text As String
    Dim result As String
    Dim pos As Long

    ' Set the worksheet
    Set ws = ActiveSheet

    ' Loop through each cell in Column C from row 8 to row 54
    For i = 8 To 54
        text = ws.Cells(i, "C").Value
        If Len(text) > 0 Then
            ' Check for the presence of "g", "mg", or "mcg" and add space before these units
            If InStr(text, "mcg") > 0 Then
                pos = InStr(text, "mcg")
            ElseIf InStr(text, "mg") > 0 Then
                pos = InStr(text, "mg")
            ElseIf InStr(text, "g") > 0 Then
                pos = InStr(text, "g")
            End If

            ' Insert the space
            If pos > 1 Then
                result = Left(text, pos - 1) & " " & Mid(text, pos)
                ws.Cells(i, "C").Value = CleanUpText(result)
            End If
        End If
    Next i
End Sub

Function CleanUpText(text As String) As String
    ' Remove trailing whitespace
    text = Trim(text)

    ' Remove double spaces
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop

    ' Remove "|"
    text = Replace(text, "|", "")

    CleanUpText = text
End Function





' --= GetLabelCopyText()


Public Sub GetLabelCopyText()
    Dim productSearchRange As Range
    Dim sizeSearchRange As Range
    Dim servingsSearchRange As Range
    Dim productCell As Range
    Dim sizeCell As Range
    Dim servingsCell As Range
    Dim productSearchString As String
    Dim productStartPosition As Long
    Dim productResult As String
    Dim sizeSearchString As String
    Dim sizeSearchStringTablet As String
    Dim sizeStartPosition As Long
    Dim sizeResult As String
    Dim servingsSearchString As String
    Dim servingsStartPosition As Long
    Dim servingsResult As String
   Application.Calculation = xlCalculationManual


    ' Set the search range to cells J8:L10
    Set productSearchRange = ActiveSheet.Range("CT8:CT10")
    Set sizeSearchRange = ActiveSheet.Range("CT10:CT14")
    Set servingsSearchRange = ActiveSheet.Range("CU15:CU25")
    On Error GoTo Exitthesub
Application.DisplayAlerts = False
    ' Loop through each cell in the search range
    For Each productCell In productSearchRange
            ' Check if the productCell starts with "Final Label Copy"
        productSearchString = "Final label copy"
        If Left(productCell.Value, Len(productSearchString)) = productSearchString Then
            ' Check if the productCell contains "000", "BL0", or "# " and extract 4 characters after the keyword
            If InStr(productCell.Value, "000") > 0 Then
                productStartPosition = InStr(productCell.Value, "000") + 3
                productResult = Mid(productCell.Value, productStartPosition, 4)


                        For Each sizeCell In sizeSearchRange

                            sizeSearchString = "Capsule size"
                            sizeSearchStringTablet = "Tablet size"
                            ' Check if the sizeCell starts with "Capsule Size"
                            If Left(sizeCell.Value, Len(sizeSearchString)) = sizeSearchString Then
                                ' Check if the sizeCell contains "000", "BL0", or "# " and extract 4 characters after the keyword
                                If InStr(sizeCell.Value, ":") > 0 Then
                                    sizeStartPosition = InStr(sizeCell.Value, ":") + 2
                                '    sizeStopPosition = ' InStr(sizeCell.Value, "clear")
                                    sizeResult = Mid(sizeCell.Value, sizeStartPosition)
                                    Exit For
                                ElseIf InStr(sizeCell.Value, " #") > 0 Then
                                    sizeStartPosition = InStr(sizeCell.Value, " #") - 1
                                    sizeResult = Mid(sizeCell.Value, sizeStartPosition, sizeStopPosition)
                                        Exit For
                                    End If

                            ' Check if the sizeCell starts with "Tablet Size"
                           ElseIf Left(sizeCell.Value, Len(sizeSearchStringTablet)) = sizeSearchStringTablet Then
                                ' Check if the sizeCell contains "000", "BL0", or "# " and extract 4 characters after the keyword
                                If InStr(sizeCell.Value, ":") > 0 Then
                                    sizeStartPosition = InStr(sizeCell.Value, ":") + 2
                                '    sizeStopPosition = ' InStr(sizeCell.Value, "clear")
                                    sizeResult = Mid(sizeCell.Value, sizeStartPosition)
                                    Exit For
                                ElseIf InStr(sizeCell.Value, " #") > 0 Then
                                    sizeStartPosition = InStr(sizeCell.Value, " #") - 1
                                    sizeResult = Mid(sizeCell.Value, sizeStartPosition, sizeStopPosition)
                                        Exit For
                                    End If
                                End If
                            Next sizeCell


                        Exit For
            ElseIf InStr(productCell.Value, " BL0") > 0 Then
                productStartPosition = InStr(productCell.Value, "BL0") + 3
                productResult = Mid(productCell.Value, productStartPosition, 4)
                Exit For
            ElseIf InStr(productCell.Value, "# ") > 0 Then
                productStartPosition = InStr(productCell.Value, "# ") + 2
                productResult = Mid(productCell.Value, productStartPosition, 4)
                Exit For
            End If
        End If
    Next productCell




    For Each servingsCell In servingsSearchRange
        ' Check if the servingsCell starts with "Final Label Copy"
        servingsSearchString = "Serving Size"
        If Left(servingsCell.Value, Len(servingsSearchString)) = servingsSearchString Then
            ' Check if the servingsCell contains "000", "BL0", or "# " and extract 4 characters after the keyword
            If InStr(servingsCell.Value, ": ") > 0 Then
                servingsStartPosition = InStr(servingsCell.Value, ": ") + 2
                servingsResult = Mid(servingsCell.Value, servingsStartPosition)
                 Exit For
            ElseIf servingsCell.Offset(0, 1).Value <> "" Then
                servingsResult = Trim(servingsCell.Offset(0, 1).Value)

                Exit For

            End If
        End If
    Next servingsCell

    If sizeResult <> "" Then ActiveSheet.Range("C5") = sizeResult


    If servingsResult <> "" Then ActiveSheet.Range("B4") = servingsResult
Exitthesub:
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
 Exit Sub
 Application.Calculation = xlCalculationAutomatic
End Sub