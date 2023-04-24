
''// Initial Test within a form - or button click
''// Try implement timer subsequently - is it a pre-requisite/requirement to keep the form open though? Confirm
''// See documentation for TransferSpreadSheet command details and arguments ::
''// https://docs.microsoft.com/en-us/office/vba/api/Access.DoCmd.TransferSpreadsheet

Private Sub SendToExcel()
    strPath = "\\srv-scd-fs-01.KHE.local\SCDData\BU3\Proj-Open\Hexion\68511-000\147-Instrument\1478-Specs-Reqs\Execute Phase\"
    strExportDate = DateValue(Now)
    If Instr(strExportDate,"/") > 0 Then
        strExportDate = RemoveSlashes(CStr(strExportDate))
    End If
    strFileName = strPath & "PipingInterface_" & strExportDate & ".xlsx"
    DoCmd.TransferSpreadsheet acExport, 10 , "_Daniel_Piping_Interface", strFileName, True''// "A1:Z400" ''// Range is optional-define range/columns in Access rather 
    Call FormatSheet(strFileName)
    ''// TODO - tweak Table/Columns/Query to be exported to align with Piping Requirements/Formatting

End Sub


Private Function RemoveSlashes(strInput)
''// Returns date fromat with dashes as opposed to with back/forward slashes
    If strInput <> "" Then
        Set regEx = CreateObject("vbscript.regexp")
        regEx.pattern = "/{1,}"
        regEx.Global = True
        regEx.IgnoreCase = True
        Set Matches = regEx.Execute(strInput)
        If Not Matches Is Nothing Then
         If Matches.Count > 0 Then
            RemoveSlashes = regEx.Replace(strInput, "-")
         End If 
        End IF 
    End If 

End Function

Private Sub FormatSheet(strFileName)
''// Formats Sheet at target location (strFileName = strPath + Date + .xlsx)
    intColor& = RGB(100, 200, 200)
    Set XlApp = CreateObject("Excel.Application")
    With XlApp
        .Visible = True
        Set oWorkBook = .WorkBooks.Open(strFileName)
    End With ''// End WITH XLApp

    Set oWorkSheet = oWorkBook.Worksheets(1)
    With oWorkSheet
        .Name = "INSTRUMENTATION"
        .DisplayPageBreaks = True
        With .Cells
            .Select
            .EntireColumn.AutoFit ''// Auto fit columns
        End With ''// End oWorkSheet.Cells WITH

        ''//Need to fix column header formatting
        For i = 1 to 14
            With .Cells(1,i) ''//Iterating through columns on top row
                .Interior.Color = RGB(255,0,0)
                .Font.Bold = True
                .Font.Color = RGB(255,255,255)
            End With
        Next i

        ' n% = .Cells(1, 1).End(xlToRight).Column
        ' For i = 1 To Columns.Count
        '     With .Cells(1, i%)
        '         w! = .EntireColumn.ColumnWidth
        '         .EntireColumn.ColumnWidth = w! + 4
        '         .HorizontalAlignment = xlCenter
        '         .Interior.Color = intColor&
        '         .Font.Bold = True
        '     End With ''// End For Loop WITH {.Cells(1,i)}

        ' Next i%

    End With ''// End oWorkSheet WITH

    With XlApp.ActiveWindow

        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True

    End With
        msg$ = vbNullString
procDone:
            Set oWorkSheet = Nothing
            Set oWorkBook = Nothing
            Set XlApp = Nothing
            
            Exit Sub
errHandler:
            msg$ = _
            Err.Number & ": " & Err.Description
            Resume procDone


End Sub


