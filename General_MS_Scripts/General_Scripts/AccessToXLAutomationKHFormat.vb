
''// Initial Test within a form - or button click
''// Try implement timer subsequently - is it a pre-requisite/requirement to keep the form open though? Confirm
''// See documentation for TransferSpreadSheet command details and arguments ::
''// https://docs.microsoft.com/en-us/office/vba/api/Access.DoCmd.TransferSpreadsheet

Private Sub SendToExcel()
    strPath = "\\srv-scd-fs-01.KHE.local\SCDData\BU3\Proj-Open\Hexion\68511-000\147-Instrument\1478-Specs-Reqs\Execute Phase\Piping Interface Test\"
    strExportDate = DateValue(Now)
    ' If Instr(strExportDate,"/") > 0 Then
    '     strExportDate = RemoveSlashes(CStr(strExportDate))
    ' End If
    strFileName = strPath & "68511-003-40-22Z-001_r0_PipingChecklist" & ".xlsx"
    DoCmd.TransferSpreadsheet acExport, 10 , "_Daniel_Piping_Interface", strFileName, True, "B13:T144" ''// Range is optional - just define range/columns in Access rather 
    ' Call FormatSheet(strFileName)
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


