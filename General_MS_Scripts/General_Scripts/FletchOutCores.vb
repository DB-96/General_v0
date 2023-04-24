''//--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x
Function DoSomeMaths(ByVal strInput As String) As  Integer 
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: DoSomeMaths()
''/// Description   ::: Takes cable size [YxYxY] and determines how many cores are required (as an integer)
''/// Variables     ::: strInput (String)
''/// Date          ::: 08/02/2023  15:41:26
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\FletchOutCores.vb
''/// Author        ::: a58ba2
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If strInput <> "" Then 
        If Instr(strInput, "x") > 1 Then
            Dim a,b as Integer
            Dim i as integer
            i = 1
            Dim Match, MatchV1 as Object
            Set regEx = CreateObject("vbscript.regexp")
            regEx.pattern = "^[0-9]{1,}x[0-9]{1,}"
            regEx.Global = True
            regEx.IgnoreCase = True
            Set Matches = regEx.Execute(strInput)
            If Not Matches Is Nothing Then
                If Matches.Count > 0 Then
                    ''//at this point you have 4x2 
                    For Each Match in Matches
                        Set regExAgain = CreateObject("vbscript.regexp")
                        regExAgain.pattern ="[0-9]{1,}"
                        regExAgain.Global = True
                        regExAgain.IgnoreCase = True
                        strTempString = Match.Value
                        Set MatchesV1 = regExAgain.Execute(strTempString)
                        If Not MatchesV1 Is Nothing Then
                            If MatchesV1.Count > 1 Then
                                For Each MatchV1 in MatchesV1 ''// y u make me do dis
                                    If i  = 1 Then
                                        a = CInt(MatchV1.Value)
                                    Else
                                        b = CInt(MatchV1.Value) 
                                    End If ''// silly silly counter check
                                    i = i+1
                                Next ''// sub match loop
                            End If ''// Hoping for two numbers to multiply - must be >1
                     End If ''// End sub match (number with in multiply)
                     DoSomeMaths = Round(a*b) ''//MATHS
                    Next ''// initial match loop
                End If  ''//see literally next line (double wrap)
            End IF  ''// ensure collection object is not empty
        End if ''// checking for {x} dimensioning - should exclude other garbage
    End If ''// End strInput!= Null check
End Function ''// End DoSomeMaths 
''//--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x

''//--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x
Sub ReadWriteData()
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: Function ReadWriteData()
''/// Description   ::: Reads data from cable list and prints a core list based on dimensions
''/// Variables     ::: MainShet (Object), CoreShet (Object)
''/// Date          ::: 10/02/2023  11:12:03
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\FletchOutCores.vb
''/// Author        ::: 90bfd2
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim MainShet as Worksheet ''//Main Cable List
    Dim CoreShet as Worksheet ''// Target Core List

    Set MainShet = ThisWorkbook.Sheets("Cables") ''//Source
    Set CoreShet = ThisWorkbook.Sheets("Cores")  ''//Target
    If Not MainShet Is Nothing  Then 
      Dim RowCount, ColCount as Long
      Dim i, j as Integer
      Dim iCoreCount As Integer
      Dim SheetCounter As Integer
      Dim strSize, strTag As String
      RowCount = MainShet.UsedRange.Rows.Count
      ColCount = MainShet.UsedRange.Columns.Count
      SheetCounter = 2 ''// CoreShet should start at row 2, will increment up
      stop
      For i = 2 to RowCount
      strSize = "" ''//reset/clear size string
      strTag = "" ''//reset/clear tag string
        For j = 1 to ColCount
            If Instr(MainShet.Cells(1,j).Value, "Size") > 0 Then
                strSize = MainShet.Cells(i,j).Value
                iCoreCount = DoSomeMaths(strSize) ''// determine core count from size attribute
            End If ''// Finding Size Column
            If Instr(MainShet.Cells(1,j).Value, "Cable_no") > 0 Then
                strTag = MainShet.Cells(i,j).Value
            End If ''// Finding Tag Column
        Next j''// iterating over columns
          If iCoreCount > 0 Then
                If strTag <> "" Then
                    SheetCounter = PrintCoreData(CoreShet,SheetCounter,iCoreCount,strTag)

                End If  ''//both tag and size must be determined
            End If
      Next i''// Iterating over all used rows (except titles)
    '   ReadWriteData = FunctionOutput
    End If ''// End MainShet and CoreShet != Null check
End Sub''// End ReadWriteData 
''//--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x



Function PrintCoreData(ByVal Shet As Object, ByVal Counter As Integer, ByVal CoreCount As Integer, ByVal strTag As String) As Integer
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: PrintCoreData()
''/// Description   ::: Prints data into specified sheet, returns incremented counter
''/// Variables     ::: ByVal rstObj ({3:DataType})
''/// Date          ::: 06/10/2022  09:33:57
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\AccesTOExcelImport_Evos.vb
''/// Author        ::: 7a12d2
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    If Not Shet Is Nothing Then
        Dim j As Integer
        For j = 1 To CoreCount ''//rows
            Shet.Cells(Counter, 1) = strTag
            Shet.Cells(Counter, 2) = CoreCount
            If j Mod 2 = 0 Then ''// Even case
                Shet.Cells(Counter, 3) = "0" + CStr(Ceil((j)/2)) + "WT"
             Else
                Shet.Cells(Counter, 3) = "0" + CStr(Ceil((j)/2)) + "ZT"
            End If ''// end odd/even check
            
            Counter = Counter + 1
        Next j
      PrintCoreData = Counter
    End If ''// End ByVal rstObj!= Null check
End Function ''// End PrintCoreData

Function Ceil(p_Number)
    Ceil = 0 - INT( 0 - p_Number)
End Function

' Sub FormatSheet(ByVal rstObj As Object, ByVal Shet As Object)
' ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' ''///
' ''/// Function Name ::: Sub FormatSheet()
' ''/// Description   ::: Formats sheet with column headers from query (this will overwrite the sheet everytime this is run - possibly put in a confirmation prompt)
' ''/// Variables     ::: rstObj (Object), Shet (Object)
' ''/// Date          ::: 06/10/2022  09:14:52
' ''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\AccesTOExcelImport_Evos.vb
' ''/// Author        ::: 9c8d38
' ''///
' ''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' ''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'     If Not rstObj Is Nothing Then
'      If Not Shet Is Nothing Then
'         Dim iCol As Integer, fldCount As Integer
'         fldCount = rstObj.Fields.Count
'         Shet.Cells.EntireColumn.AutoFit
'         Shet.Cells.Font.Name = "ISOCTEUR"
'         Shet.Cells.Font.Size = 9
'         Shet.Cells.Font.Color = RGB(0,0,0)
'         Shet.Range("A1:AB1").AutoFilter
'             For iCol = 1 To fldCount ''// print column headers from recordset headers
'                 Shet.Cells(1, iCol).Value = rstObj.Fields(iCol - 1).Name ''// print column header
'                 ''//now format column header
'                 With Shet.Cells(1, iCol) ''//Iterating through columns on top row
'                     ' .Interior.Color = RGB(255,0,0)
'                     .Interior.Color = RGB(171, 37, 10)
'                     .Font.Bold = True
'                     .Font.Color = RGB(255, 255, 255)
'                     .Font.Size = 10 ''// why does this crash the entire import?????????????????????
'                 End With
'             Next
'         End If ''// end sheet != Null check
'     End If ''// End rstObj and Shet != Null check
' End Sub ''// End FormatSheet
