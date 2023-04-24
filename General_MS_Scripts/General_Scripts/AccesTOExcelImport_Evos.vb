

Function SelectDataBase() As String
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: SelectDataBase ()
''/// Description   ::: Prompts user to select and ElecDes database
''/// Variables     ::: N/A
''/// Date          ::: 23/09/2022  11:36:38
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\ExcelTOElecDesAccessImporter.vb
''/// Author        ::: 55a618
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim fd As Office.FileDialog
    Dim strFile As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If Not fd Is Nothing Then
       With fd
            .Filters.Clear
            .Filters.Add "Access Files", "*.mdb;*.accdb", 1  ''// Will it always be an Access file??
            .Title = "Choose the Source EVOS Project Database :: "
            .AllowMultiSelect = False
            .InitialFileName = "K:\BU3\Proj-Open\Evos\69286-000\147-Instrument\1475-Model\7000 Database"  ''// TODO:: move this initial file name to project root directory/server location where database will actually be running
            If .Show = True Then
                If .SelectedItems(1) <> "" Then
                    strPrompt = "You have selected the following database: " & .SelectedItems(1) & vbCrLf & "Would you like to continue?"  ''// Optional
                    resp = MsgBox(strPrompt, vbSystemModal, "Warning")                                                                  ''// Optional
                    ''// Optional :: User to double check database that has been chosen within message box
                    SelectDataBase = .SelectedItems(1)
                End If ''// Ensure the selection string is not empty
            End If
        End With
    End If   ''// End fd != Null check
End Function ''// End SelectDataBase

Sub BuildAndExeSQLString()
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name :::  Sub BuildAndExeSQLString ()
''/// Description   :::
''/// Variables     ::: N/A
''/// Date          ::: 23/09/2022  11:42:25
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\ExcelTOElecDesAccessImporter.vb
''/// Author        ::: f93a1d
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim i As Integer, j As Integer              ''// For loop iterators
    Dim RowCount As Long, ColCount As Long      ''// Excel Row/Column containers
    Dim Sh As Worksheet                         ''// Excel Worksheet to be imported
    Dim oConnObj As Object                      ''// Database Connection Object
    Dim strValues As String, strInsert As String, strSheet As String

    
    ' If Not Sh Is Nothing Then
        strDev = "SELECT * FROM DDB_TestQuery" ''// SQL statement scaffolding - this query has been pre-configured for the loop sheet export
        ' RowCount = Sh.UsedRange.Rows.Count
        ' ColCount = Sh.UsedRange.Columns.Count
        Set oConnObj = CreateObject("ADODB.Connection") ''// Initialize ADO connection object
        
        If Not oConnObj Is Nothing Then
            strSourceDB = SelectDataBase()              ''//User selected database (ElecDes Project)
            If strSourceDB <> "" Then
                strConnTarget = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strSourceDB ''// Will it always be an Access file??
                oConnObj.Open strConnTarget
                Set rstObj = CreateObject("ADODB.Recordset")
                If Not rstObj Is Nothing Then
                    rstObj.Open strDev, oConnObj ''// run the damn query man
                    Call SplitData(rstObj)
                End If ''// end check for record set object existence
                oConnObj.Close     ''// Close database
                ' rstObj.Close
                Set rstObj = Nothing
                Set oConnObj = Nothing
            End If ''// End strSourceDb != Null check
        End If ''// End DB Connection Object != Null Check
    ' End If ''// End Sh != Null check
End Sub ''// End BuildAndExeSQLString

Sub SplitData(ByVal rstObj As Object)
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: Sub SplitData()
''/// Description   ::: Splits data from SQL query into different excel sheets based on loop typical type =>> for AutoCad automation
''/// Variables     ::: rstObj (ADODB.Recordset)
''/// Date          ::: 05/10/2022  22:30:40
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\AccesTOExcelImport_Evos.vb
''/// Author        ::: d84dcc
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If Not rstObj Is Nothing Then
        i = 1
        Dim Count_PT_2, Count_PT_1, Count_TT_1, Count_TT_2, Count_DCS_Valve, Count_SIS_Valve, Count_VSD, Count_DCS_LS, Count_SIS_LS, Count_DCS_HV, Count_SIS_NS, Count_DCSBeds As Integer
         Count_PT_2 = 2
         Count_PT_1 = 2
         Count_TT_1 = 2
         Count_TT_2 = 2
         Count_DCS_Valve = 2
         Count_SIS_Valve = 2
         Count_VSD = 2
         Count_DCS_LS = 2
         Count_SIS_LS = 2
         Count_DCS_HV = 2
         Count_SIS_NS = 2
         Count_DCSBeds = 2
        Do While Not rstObj.EOF
            If CStr(rstObj("typical")) <> "" Then
                Select Case CStr(rstObj("typical"))
                Case "01_SIS_NS"
                Set Sh = ThisWorkbook.Sheets("01_SIS_NS") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_SIS_NS = PrintData(rstObj, Sh, Count_SIS_NS)
                        i = i + 1
                    End If

                Case "05_DCS_Bed_cons"
                   Set Sh = ThisWorkbook.Sheets("05_DCS_Bed_cons") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_DCSBeds = PrintData(rstObj, Sh, Count_DCSBeds)
                        i = i + 1
                    End If

                Case "09_VSD_Pomp"
                    Set Sh = ThisWorkbook.Sheets("09_VSD_Pomp") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_VSD = PrintData(rstObj, Sh, Count_VSD)
                        i = i + 1
                    End If

                Case "11_DCS_LS"
                    Set Sh = ThisWorkbook.Sheets("11_DCS_LS") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_DCS_LS = PrintData(rstObj, Sh, Count_DCS_LS)
                        i = i + 1
                    End If

                Case "12_SIS_LS"
                    Set Sh = ThisWorkbook.Sheets("12_SIS_LS") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_SIS_LS = PrintData(rstObj, Sh, Count_SIS_LS)
                        i = i + 1
                    End If

                Case "13_ENRAF_LT"

                Case "14_RADAR_LT"

                Case "15_DCS_SEAL"

                Case "16_SIS_SEAL"

                Case "17_DCS_LS_PUT"

                Case "19_PT_1_2"
                    Set Sh = ThisWorkbook.Sheets("19_PT_1_2") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_PT_1 = PrintData(rstObj, Sh, Count_PT_1)
                        i = i + 1
                    End If

                Case "20_PT_2_2"
                    Set Sh = ThisWorkbook.Sheets("20_PT_2_2") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_PT_2 = PrintData(rstObj, Sh, Count_PT_2)
                        i = i + 1
                    End If

                Case "22_TT_1_2"
                    Set Sh = ThisWorkbook.Sheets("22_TT_1_2") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_TT_1 = PrintData(rstObj, Sh, Count_TT_1)
                        i = i + 1
                    End If

                Case "23_TT_2_2"
                     Set Sh = ThisWorkbook.Sheets("23_TT_2_2") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_TT_2 = PrintData(rstObj, Sh, Count_TT_2)
                        i = i + 1
                    End If

                Case "24_TT_OS"

                Case "25_SIG_DRAIN"

                Case "26_CLIXON"

                Case "27_DCS_HV"
                    Set Sh = ThisWorkbook.Sheets("27_DCS_HV") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                     ''//technically only need to format the sheet once, currently will format it every iteration of i, optimize (TODO)
                        Call FormatSheet(rstObj, Sh)
                        Count_DCS_HV = PrintData(rstObj, Sh, Count_DCS_HV)
                        i = i + 1
                    End If

                Case "28_DCS_VALVE"
                    Set Sh = ThisWorkbook.Sheets("28_DCS_VALVE") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_DCS_Valve = PrintData(rstObj, Sh, Count_DCS_Valve)
                        i = i + 1
                    End If

                Case "29_SIS_VALVE"
                 Set Sh = ThisWorkbook.Sheets("29_SIS_VALVE") ''//Sheet == Typical
                    ''//Consider printing first column headers
                    If Not Sh Is Nothing Then
                        Call FormatSheet(rstObj, Sh)
                        Count_SIS_Valve = PrintData(rstObj, Sh, Count_SIS_Valve)
                        i = i + 1
                    End If

            End Select

            End If ''// end check for no loop typical stated (really an overkill)
            rstObj.MoveNext
            
        Loop ''//End while loop
    End If ''// End rstObj != Null check
End Sub ''// End SplitData

Function PrintData(ByVal rstObj As Object, ByVal Shet As Object, ByVal Counter As Integer) As Integer
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: PrintData()
''/// Description   ::: Prints data into specified sheet, returns incremented counter
''/// Variables     ::: ByVal rstObj ({3:DataType})
''/// Date          ::: 06/10/2022  09:33:57
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\AccesTOExcelImport_Evos.vb
''/// Author        ::: 7a12d2
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If Not rstObj Is Nothing Then
        Dim j As Integer
        For j = 0 To rstObj.Fields.Count - 1
            Shet.Cells(Counter, j + 1) = rstObj.Fields.Item(j)
        Next j
      PrintData = Counter + 1
    End If ''// End ByVal rstObj!= Null check
End Function ''// End PrintData


Sub FormatSheet(ByVal rstObj As Object, ByVal Shet As Object)
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: Sub FormatSheet()
''/// Description   ::: Formats sheet with column headers from query (this will overwrite the sheet everytime this is run - possibly put in a confirmation prompt)
''/// Variables     ::: rstObj (Object), Shet (Object)
''/// Date          ::: 06/10/2022  09:14:52
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\AccesTOExcelImport_Evos.vb
''/// Author        ::: 9c8d38
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If Not rstObj Is Nothing Then
     If Not Shet Is Nothing Then
        Dim iCol As Integer, fldCount As Integer
        fldCount = rstObj.Fields.Count
        Shet.Cells.EntireColumn.AutoFit
        Shet.Cells.Font.Name = "ISOCTEUR"
        Shet.Cells.Font.Size = 9
        Shet.Cells.Font.Color = RGB(0,0,0)
        Shet.Range("A1:AB1").AutoFilter
            For iCol = 1 To fldCount ''// print column headers from recordset headers
                Shet.Cells(1, iCol).Value = rstObj.Fields(iCol - 1).Name ''// print column header
                ''//now format column header
                With Shet.Cells(1, iCol) ''//Iterating through columns on top row
                    ' .Interior.Color = RGB(255,0,0)
                    .Interior.Color = RGB(171, 37, 10)
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Size = 10 ''// why does this crash the entire import?????????????????????
                End With
            Next
        End If ''// end sheet != Null check
    End If ''// End rstObj and Shet != Null check
End Sub ''// End FormatSheet

Function ToXSO(strInput As String) As String
    If strInput <> "" Then
        If InStr(strInput, "XSV") > 0 Then
            ToXSO = "'" & Replace(strInput, "XSV", "XSO") & "'" ' ''// should probably use RegExp here (TODO)
        End If ''// end check for solenoid valve
    End If ''// end check for empty string
End Function

Function ToXSC(strInput As String)
    If strInput <> "" Then
        If InStr(strInput, "XSV") > 0 Then
            ToXSC = "'" & Replace(strInput, "XSV", "XSC") & "'" ''// should probably use RegExp here (TODO)
        End If ''// end check for solenoid valve
    End If ''// end check for empty string
End Function


' Private Sub FormatSheet(strFileName)
' ''// Formats Sheet at target location (strFileName = strPath + Date + .xlsx)
'     intColor& = RGB(100, 200, 200)
'     Set XlApp = CreateObject("Excel.Application")
'     With XlApp
'         .Visible = True
'         Set oWorkBook = .WorkBooks.Open(strFileName)
'     End With ''// End WITH XLApp

'     Set oWorkSheet = oWorkBook.Worksheets(1)
'     With oWorkSheet
'         .DisplayPageBreaks = True
'         With .Cells
'             .Select
'             .EntireColumn.AutoFit ''// Auto fit columns
'         End With ''// End oWorkSheet.Cells WITH

'         ''//Need to fix column header formatting
'         For i = 1 to 14
'             With .Cells(1,i) ''//Iterating through columns on top row
'                 .Interior.Color = RGB(255,0,0)
'                 .Font.Bold = True
'                 .Font.Color = RGB(255,255,255)
'             End With
'         Next i

'     End With ''// End oWorkSheet WITH

'     With XlApp.ActiveWindow

'         .SplitColumn = 0
'         .SplitRow = 1
'         .FreezePanes = True

'     End With
'         msg$ = vbNullString
' procDone:
'             Set oWorkSheet = Nothing
'             Set oWorkBook = Nothing
'             Set XlApp = Nothing
            
'             Exit Sub
' errHandler:
'             msg$ = _
'             Err.Number & ": " & Err.Description
'             Resume procDone


' End Sub
      
' Loop typicals
' typical naam
' 01_SIS_NS
' 02_PANEL_NS
' 03_BMI_HBM
' 04_BMI_QBM
' 05_DCS_Bed_cons
' 07_BMI_FSL
' 08_BMI_HRN
' 09_VSD_Pomp
' 11_DCS_LS
' 12_SIS_LS
' 13_ENRAF_LT
' 14_RADAR_LT
' 15_DCS_SEAL
' 16_SIS_SEAL
' 17_DCS_LS_PUT
' 19_PT_1_2
' 20_PT_2_2
' 22_TT_1_2
' 23_TT_2_2
' 24_TT_OS
' 25_SIG_DRAIN
' 26_CLIXON
' 27_DCS_HV
' 28_DCS_VALVE
' 29_SIS_VALVE


                    ' set Templ = rstObj("typical")
                    ' recArray = rstObj.GetRows ''// 0 based array, 1st dimension = fields, 2nd dimension records
                    ' recCount = UBound(recArray,2)+1 ''// transpose and begin index from 1(+1)

                    ' For i = 1 to recCount

                    ' Next i
                    ' Set Sh = ThisWorkbook.Sheets(strSheet) ''//Sheet == Typical
                    ' fldCount = rstObj.Fields.Count
                    ' stop
                    ' For iCol = 1 To fldCount
                    '     Sh.Cells(1, iCol).Value = rstObj.Fields(iCol - 1).Name
                    ' Next
                    ' Sh.Cells(2, 1).CopyFromRecordset rst


' Limit_Switch_Open: IIf(InStr([Tag_number New],"XSV")>0,Replace([Tag_number New],"XSV","XSO"),'0')







