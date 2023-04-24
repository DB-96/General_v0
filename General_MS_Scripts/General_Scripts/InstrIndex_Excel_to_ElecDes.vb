''// 1 ->> Create Sub Instrument in SubTable (with relevant SQL Query) :: Function must return IDX and Table Name to be written to Tag_Tags
''// 2 ->> Create Equipment in SubTable (Equipment_Equipment){with relevant SQL query} :: Function must return IDX and Table Name (Equipment_Equipment) to be written to Tag_Tags
''// 3 ->> Create Line Number in SubTable (Line_Lines) {with relevant SQL Query} :: Function must return IDX
''// 3 ->> Create Main Instrument in Tag_Tags {Associating Instrument->R1_IDX, Equipment->A20_IDX, Line ->}
''//
''//


Function SelectMyDataBase() As String
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
            .Filters.Add "Access Files", "*.mdb", 1  ''// Will it always be an Access file??
            .Title = "Choose the Target ElecDes Project Database :: "
            .AllowMultiSelect = False
            .InitialFileName = "C:\edstest"  ''// TODO:: move this initial file name to project root directory/server location where database will actually be running
            If .Show = True Then
                If .SelectedItems(1) <> "" Then
                    strPrompt = "You have selected the following database: " & .SelectedItems(1) & vbCrLf & "Would you like to continue?"   ''// Optional
                    resp = MsgBox(strPrompt, vbSystemModal, "Warning")                                                                  ''// Optional
                    ''// Optional :: User to double check database that has been chosen within message box
                    SelectMyDataBase = .SelectedItems(1)
                End If ''// Ensure the selection string is not empty
            End If
        End With
    End If   ''// End fd != Null check
End Function ''// End SelectDataBase

Sub CreateSubInstrument(oConnObj As Object, strInstrType As String, strInputs As String)
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: CreateSubInstrument()
''/// Description   ::: Create
''/// Variables     ::: oConnObj ({3:DataType}), strInstrType - new entries written to different tables based on instrument type
''/// Date          ::: 28/09/2022  11:28:57
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\InstrIndex_Excel_to_ElecDes.vb
''/// Author        ::: 743804
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If Not oConnObj Is Nothing Then
      If strInstrType <> "" Then
        Dim strInsert As String, strValues As String, strSQLFinal As String
        strInputs = Left(strInputs, Len(strInputs) - 1) ''// remove last comma
        strInsert = "INSERT INTO " & strInstrType & "(LastModified, PID_No, I_Service, I_LoopNumber, I_Plant, I_Location)"
        strValues = "VALUES( 'Admin_Import'," & strInputs & ")"
        strSQLFinal = strInsert & strValues
        oConnObj.Execute (strSQLFinal)
      End If
    End If ''// End oConnObj!= Null check
End Sub ''// End CreateSubInstrument


Sub CreateSubEquipment(oConnObj As Object, strInputs As String)
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: Sub CreateEquipment()
''/// Description   ::: Creates blank entry for equipment in Equipment_Equipment table, to be linked in Tag_Tags folder. This function solely writes to this table
''/// Variables     ::: oConnObj (Object), strInputs (String)
''/// Date          ::: 28/09/2022  12:38:16
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\InstrIndex_Excel_to_ElecDes.vb
''/// Author        ::: c76432
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If Not oConnObj Is Nothing Then
     If strInputs <> "" Then
      Dim strInsert As String, strValues As String, strSQLFinal As String
      strInsert = "INSERT INTO Equipment_Equipment(LastModified)"
      strValues = "VALUES('Admin_Import')"
      strSQLFinal = strInsert & strValues
      oConnObj.Execute (strSQLFinal)
     End If ''// End check for empty input string
    End If ''// End oConnObj and strInputs != Null check
End Sub ''// End CreateEquipment

Sub CreateMainEquipment(oConnObj As Object, strInput As String, strAppend As String)
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: Sub CreateMainEquipment()
''/// Description   ::: Creates Equipment in Tag_Tags and links to relational field R1
''/// Variables     ::: oConnObj (Object), strInput (String), strAppend (String) link from sub table
''/// Date          ::: 28/09/2022  13:17:13
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\InstrIndex_Excel_to_ElecDes.vb
''/// Author        ::: b13573
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If Not oConnObj Is Nothing Then
     If strInput <> "" Then ''//Tag number, descriptions etc - these are inputs from the import sheet
        If strAppend <> "" Then
            Dim strInsert As String, strValues As String, strSQLFinal As String
            strTag_Tags_Cols = "Tagname, Component_Type, R1_IDX, R1_TABLE" ''// Column headers from ElecDes Tag_Tags table TOD0::Input equipment description
            strInsert = "INSERT INTO Tag_Tags(" & strTag_Tags_Cols & ") " ''// SQL INSERT statement scaffolding
            strValues = "SELECT '" & strInput & "', 'Equipment_Equipment', " & strAppend & " FROM Equipment_Equipment;" ''// Remove final comma (,) - write final SQL query to Tag_Tags
            strSQLFinal = strInsert & strValues
            oConnObj.Execute (strSQLFinal)
        End If ''// End check for empty append string for association
     End If ''// end check for empty input string
    End If ''// End oConnObj and strInput != Null check
End Sub ''// End CreateMainEquipment

Sub BuildAndExeSQLString()
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name :::  Sub BuildAndExeSQLString ()
''/// Description   ::: Legacy, superceded by Sub Create Main Instrument
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
    Dim strValues As String, strInsert As String
    Dim strInstrRelate As String, strSubtable As String
    Dim strAppendEquip As String, strEquipTag As String

    Set Sh = ThisWorkbook.Sheets("ImportSheet") ''//Import sheet must be named 'ImportSheet'
    If Not Sh Is Nothing Then
        strTag_Tags_Cols = "Tagname, Description, Component_Type, R1_IDX, R1_TABLE" ''// Column headers from ElecDes Tag_Tags table
        strDev = "INSERT INTO Tag_Tags(" & strTag_Tags_Cols & ") " ''// SQL INSERT statement scaffolding
        RowCount = Sh.UsedRange.Rows.Count
        ColCount = Sh.UsedRange.Columns.Count
        Set oConnObj = CreateObject("ADODB.Connection") ''// Initialize ADO connection object
        If Not oConnObj Is Nothing Then
            strSourceDB = SelectMyDataBase()              ''//User selected database (ElecDes Project)
            If strSourceDB <> "" Then
                strConnTarget = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strSourceDB ''// Will it always be an Access file??
                oConnObj.Open strConnTarget                   ''// open database
                For i = 2 To RowCount                         ''// iterate among rows
                    strCols = ""                              ''// reset VALUES element to re-write new fields
                    strSubtable = ""
                    strRelate = ""
                    strAppendInstr = ""
                    strAppendEquip = ""
                    strEquipTag = ""
                    For j = 1 To ColCount                     ''// iterate among columns
                        If InStr(Sh.Cells(1, j).Value, "Ratings") > 0 Then ''// Check for relational Column Headers
                        ''// cue emotional trauma
                            strInstrRelate = Sh.Cells(i, 4).Value  ''// Column D hardcoded to instrument type (careful of pivot point) - basically dont change this in import sheet
                            strAppendInstr = "MAX(IDX) , '" & strInstrRelate & "'"
                            If i > 1 Then ''// do not actually write a column header to the table
                                If j > 4 Then ''// after instrument type - any relational table data must appear after column 4 basically
                                    strSubtable = strSubtable & " '" & Sh.Cells(i, j).Value & "',"
                                   ''// strCols = strCols & "(SELECT IDX FROM" & Sh.Cells(i, j).Value & "),"  ''//retrieve last added ID and add to table entry
                                End If
                            End If ''// end column header exclusion
                        ElseIf InStr(Sh.Cells(1, j).Value, "Equipment Object") > 0 Then ''// Check for Associated Equipment column header
                            ''//Check Equip existence in Tag_Tags fist
                            intEquip = j
                            If i > 1 Then
                             If Sh.Cells(i, intEquip).Value <> "" Then
                                Call CreateSubEquipment(oConnObj, "seed")
                                strAppendEquip = "MAX(IDX) , 'Equipment_Equipment'"
                                strEquipTag = Sh.Cells(i, intEquip).Value
                             End If ''// End check for associated equipment field being populated in main sheet
                            End If ''// end column header exclusion
                            
                        Else       ''// NORMAL Case - Tag_Tags table only
                            strCols = strCols & " '" & Sh.Cells(i, j).Value & "',"
                        End If     ''//End relational column headers checks
                    Next j
                    Stop
                    Call CreateSubInstrument(oConnObj, strInstrRelate, strSubtable) ''// sub instrument needs to be created ANYWAY
                    If strAppendEquip <> "" Then ''// there is actually an associated equipment
                        Call CreateMainEquipment(oConnObj, strEquipTag, strAppendEquip) ''//create object in main Tag_Tags Table Including SELECT of Max ID from Equipment_Equipment
                        strTag_Tags_Cols = "Tagname, Description, Component_Type, R1_IDX, R1_TABLE, A20_Table, A20_IDX " ''// Column headers from ElecDes Tag_Tags table :: here we overwrite strDev with additional equipment relation columns
                        strDev = "INSERT INTO Tag_Tags(" & strTag_Tags_Cols & ") " ''// SQL INSERT statement scaffolding
                        strCols = strCols & strAppendInstr
                        strCols = "SELECT " & strCols & "'Tag_Tags' FROM " & strInstrRelate & ","
                        strCols = strCols & "SELECT MAX(IDX),  FROM Tag_Tags ;"
                     Else ''// there is NOT an associated equipment
                      strCols = strCols & strAppendInstr ''//append strCols with ID from relational table and table name
                      strCols = "SELECT " & Left(strCols, Len(strCols) - 1) & " FROM " & strInstrRelate & ";" ''// Remove final comma (,) - write final SQL query to Tag_Tags
                    End If
                    
                    
                    
                    strValues = ""
                    ''// SQL string to be executed should look something along the lines of::
                    ''// INSERT INTO Tag_Tags( column_1 , column_2, ... , column n) SELECT 'string 1', 'string2',...., MAX(IDX), .... , 'string_n') FROM <relational_table>
                    ''// Need to use SELECT FROM to pull out relational table ID

                    Set oTemp = oConnObj.Execute(strDev & vbCrLf & strCols)
                Next i
                oConnObj.Close     ''// Close database
            End If ''// End strSourceDb != Null check
        End If ''// End DB Connection Object != Null Check
    End If ''// End Sh != Null check
End Sub ''// End BuildAndExeSQLString

''//--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x
Function CheckSQLDuplicates(oConnObj As Object, Table_Name As String, Tagname As String) As Boolean
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: Function CheckSQLDuplicates()
''/// Description   ::: POC existence function, to be done before creating instruments/pipes/equipment etc
''/// Variables     ::: oConnObj (Object), Table_Name (String), Tagname (String)
''/// Date          ::: 03/01/2023  10:04:16
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\InstrIndex_Excel_to_ElecDes.vb
''/// Author        ::: d428ca
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If Not oConnObj Is Nothing Then 
      If Table_Name <> "" Then 
         If Tagname <> "" Then 
            '' // FunctionBody 
             strSQL = "SELECT 1 FROM Tag_Tags WHERE Tagname = " & Tagname & " LIMIT 1;"
             intResult = oConnObj.Execute(strSQL)
             Select Case intResult
                Case 1
                    CheckSQLDuplicates = True
                Case Else
                    CheckSQLDuplicates = False
             End Select
         End If ''// End oConnObj and Table_Name and Tagname != Null check
      End If ''// End oConnObj and Table_Name and Tagname != Null check
    End If ''// End oConnObj and Table_Name and Tagname != Null check
End Function ''// End CheckSQLDuplicates 
''//--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x


















