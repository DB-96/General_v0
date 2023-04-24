
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
            .Filters.Add "Access Files", "*.mdb", 1  ''// Will it always be an Access file??
            .Title = "Choose the Target ElecDes Project Database :: "
            .AllowMultiSelect = False
            .InitialFileName = "C:\edstest"  ''// TODO:: move this initial file name to project root directory/server location where database will actually be running
            If .Show = True Then
                If .SelectedItems(1) <> "" Then 
                    strPrompt = "You have selected the following database: " & SelectDataBase & vbCrLf & "Would you like to continue?"  ''// Optional
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
    Dim strValues As String, strInsert As String

    Set Sh = ThisWorkbook.Sheets("ImportSheet") ''//Import sheet must be named 'ImportSheet'
    If Not Sh Is Nothing Then
        strTag_Tags_Cols = "Tagname, Description, Component_Type, R1_IDX, R1_TABLE" ''// Column headers from ElecDes Tag_Tags table
        strDev = "INSERT INTO Tag_Tags(" & strTag_Tags_Cols & ") " ''// SQL INSERT statement scaffolding
        RowCount = Sh.UsedRange.Rows.Count
        ColCount = Sh.UsedRange.Columns.Count
        Set oConnObj = CreateObject("ADODB.Connection") ''// Initialize ADO connection object 
        If Not oConnObj Is Nothing Then
            strSourceDB = SelectDataBase()              ''//User selected database (ElecDes Project)
            If strSourceDb <> "" Then 
                strConnTarget = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strSourceDB ''// Will it always be an Access file??
                oConnObj.Open strConnTarget                   ''// open database
                For i = 2 To RowCount                         ''// iterate among rows
                    strCols = ""                              ''// reset VALUES element to re-write new fields
                    strRelate = ""
                    For j = 1 To ColCount                     ''// iterate among columns
                        If InStr(Sh.Cells(1, j).Value, "Relational") > 0 Then ''// Check for relational Column Headers
                        ''// cue emotional trauma
                            If i > 1 Then ''// do not actually write a column header to the table
                                strInsert = "INSERT INTO " & Sh.Cells(i, j).Value & "(LastModified)"
                                strValues = "VALUES ('Admin_Import')"
                                strRelate = Sh.Cells(i, j).Value            ''// Relational table name
                                strSQLFinal = strInsert & strValues
                                Set rs = oConnObj.Execute(strSQLFinal)
                                strCols = strCols & " MAX(IDX) , '" & strRelate & "',"  
                                   ''// strCols = strCols & "(SELECT IDX FROM" & Sh.Cells(i, j).Value & "),"  ''//retrieve last added ID and add to table entry
                            End If ''// end column header exclusion
                        Else       ''// NORMAL Case - Tag_Tags table only
                            strCols = strCols & " '" & Sh.Cells(i, j).Value & "',"
                        End If     ''//End relational column headers checks
                    Next j
                    strCols = "SELECT " & Left(strCols, Len(strCols) - 1) & " FROM " & strRelate & " ;" ''// Remove final comma (,)
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





