Sub ProjectLookup()
    Dim ClientName As String
    Dim ProjectNumber As String
    Dim ProjectPhase As String
    Dim ProjectResult As String
    
    'Display user form and retrieve input values
    Set frmProjectLookup = UserForm1
    frmProjectLookup.Show
    With frmProjectLookup
        ' ClientName = .txtClientName.Value
        ProjectNumber = .txtProjectNumber.Value
        ProjectPhase = .cboProjectPhase.Value
    End With
    
    'Lookup project result based on project phase
    Select Case ProjectPhase
        Case "FEL1"
            ProjectResult = "1011"
        Case "FEL2"
            ProjectResult = "1012"
        Case "FEL3"
            ProjectResult = "1013"
        Case "FEL4"
            ProjectResult = "1014"
        Case Else
            ' MsgBox "Invalid project phase selected."
            Exit Sub
    End Select
    
    'Display results to user
    ' MsgBox "Client Name: " & ClientName & vbNewLine & "Project Number: " & ProjectNumber & vbNewLine & "Project Phase: " & ProjectPhase & vbNewLine & "Project Result: " & ProjectResult
End Sub


' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' USER FORM 2
' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub UserForm_Initialize()

'     Dim wbk     As Workbook
'     Set wbk = Workbooks("Estimating Tool.xlsm")
'     ListBox1.RowSource = "OverviewSheet!A12:D90"
'     ListBox1.ColumnCount = 4
'    'ListBox1.Grid
'     ListBox1.ColumnWidths = "28;155;30;25"
'     ListBox1.ColumnHeads = True
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OverviewSheet")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ''////////////////////////////////////////////////////////////////
    ''// ListView >>> ListBox
    ''//
    ''////////////////////////////////////////////////////////////
    ListView1.ColumnHeaders.Clear
    'ListView1.View = lvwReport
    'ListView1.View = lvwReport
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Add , , "WBS", 27, 0
    ListView1.ColumnHeaders.Add , , "Component Description", 165, 0
    ListView1.ColumnHeaders.Add , , "Code", 30, 2
    ListView1.ColumnHeaders.Add , , "Count", 35, 2
   ' ListView1.RowSource = "OverviewSheet!A13:D90"
   'ListView1.RowSource =
     Dim i As Long
    For i = 13 To lastRow
        Dim item As ListItem
        Set item = ListView1.ListItems.Add(, , ws.Cells(i, 1).Value)
        item.SubItems(1) = ws.Cells(i, 2).Value
        item.SubItems(2) = ws.Cells(i, 3).Value
        item.SubItems(3) = ws.Cells(i, 4).Value
    Next i
   

End Sub

Private Sub ListView1_DblClick()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OverviewSheet")
    
    Dim i As Integer
    Dim ItemSel As ListItem
    If Not ListView1.selectedItem Is Nothing Then
        Set ItemSel = ListView1.selectedItem
        Call EditData
        'Stop
    End If ''// ensure something is actually selected
    


End Sub

Sub EditData()
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name :::  Sub EditData()
''/// Description   :::
''/// Variables     ::: N/A
''/// Date          ::: 17/04/2023  23:43:10
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\Templates\EstimatingTool.vb
''/// Author        ::: 77ac61
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'    Stop
    If Not Me.ListView1.selectedItem Is Nothing Then
      Dim row, col, i As Integer
      Set ItemSel = ListView1.selectedItem
      i = 3
      On Error Resume Next
      
      With Frame1
        
        .Visible = True
        .Top = ItemSel.Top + ListView1.Top
        .Left = ListView1.ColumnHeaders(i + 1).Left + ListView1.Left
        .Width = ListView1.ColumnHeaders(i + 1).Width
        .Height = ItemSel.Height
        
        .ZOrder msoBringToFront
    End With
    
      With EditTextBox1
        .ZOrder msoBringToFront
        .Visible = True
        .Text = ItemSel.SubItems(i) ''// initial text to be placed in dummy text box
        .SetFocus
        .SelStart = 0
        .Left = 0
        .Top = ListView1.selectedItem.Top
        .Width = ListView1.ColumnHeaders(i + 1).Width
        ' .Top = ItemSel.Top + ListView1.Top
        ' .Left = ListView1.ColumnHeaders(i + 1).Left + ListView1.Left
        ' .Width = ListView1.ColumnHeaders(i + 1).Width
        ' .Height = ItemSel.Height
        .Height = ListView1.selectedItem.Height
        .SelLength = Len(.Text)
        .ZOrder msoBringToFront
        
      End With
    End If ''// End Me.UserForm2.ListBox1.SelectedItem != Null check
End Sub ''// End EditData

Private Sub ListView1_Click()
    Me.Frame1.Visible = False
    Me.  Me.EditTextBox1.Value 
    Me.EditTextBox1.Value = ""
   ' Me.infoCb.Value = ""
End Sub
' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' USER FORM 1
' ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Option Explicit
Dim ProjectNumber As String


Private Sub CommandButton1_Click()
    ''// TODO - Ensure that a project phase is actually selected and that the dictionary is populated before you hop into the other screen
    ''// print a message box warning in the case that this is not generated
    cboProjectPhase.Value = Null
    Me.DeliverableListBox.ListIndex = -1
    Me.DeliverableListBox.Clear
    ''// DOES THE DICTIONARY GET CLEARED THOUGH?
End Sub

Private Sub CommandButton2_Click()
    Dim ProjectPhase As String
    ProjectPhase = UserForm1.cboProjectPhase.Value
    If ProjectPhase <> "" Then
             UserForm2.Show
        Else
            MsgBox "Please select an appropriate project phase and associated deliverables before progressing!"
    End If ''// Ensure Project Phase has been selected before showing next screen
   
End Sub

Private Sub UserForm_Initialize()
    'Populate project phase options in combo box
    ''// Scaffolding 
    ''// Setting up combo box options and default fields

    cboProjectPhase.AddItem "FEL1"
    cboProjectPhase.AddItem "FEL2"
    cboProjectPhase.AddItem "FEL3"
    cboProjectPhase.AddItem "FEL4"
    framePhase.Visible = False

    cboClientName.AddItem "Evos"
    cboClientName.AddItem "Hexion"
    cboClientName.AddItem "Koole Terminals"
    cboClientName.AddItem "Shin-Etsu PVC"
    cboClientName.AddItem "Shell Nederland"
    cboClientName.AddItem "Neste Chemicals"
    cboClientName.AddItem "Vopak"
    cboClientName.AddItem "Westlake Epoxy"
    
    cboDiscipline.AddItem "120 - Process Engineering"
    cboDiscipline.AddItem "130 HSSE Engineering"
    cboDiscipline.AddItem "141 - CSA"
    cboDiscipline.AddItem "145 - Mechanical"
    cboDiscipline.AddItem "146 - Piping"
    cboDiscipline.AddItem "147 - I&PC"
    cboDiscipline.AddItem "148 - Electrical"

    cboDiscipline.ListIndex = 5

    cboProjectType.AddItem "Brownfield"
    cboProjectType.AddItem "Greenfield"

    

End Sub
''//--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x
Function PopulateDictionary(strPhase As String) As Object
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: PopulateDictionary()
''/// Description   ::: Returns a dictionary of deliverables (by WBS code) and document numbers based on the project phase selected, be aware that in this particular case
''///                   the DELIVERABLES worksheet is used as a reference to 'fetch' the deliverables
''/// Variables     ::: strPhase ({String})
''/// Date          ::: 12/04/2023  10:35:46
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\Templates\EstimatingTool.vb
''/// Author        ::: 8cee72
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// What about adding deliverables that are not part of the QPM??
    If strPhase <> "" Then
   ' Set deliverables = CreateObject("Scripting.Dictionary")
    Dim deliverables As New Scripting.Dictionary
    Dim DelivWS As Worksheet
    Set DelivWS = ThisWorkbook.Worksheets("DELIVERABLES") ''// deliverables worksheet {QPM}
    Dim RowCount, ColCount As Long
    Dim i, j As Integer
    RowCount = DelivWS.UsedRange.Rows.Count
    ColCount = DelivWS.UsedRange.Columns.Count

    For i = 2 To RowCount
      
      Dim WBSCode As String
      Dim WBSCodeDesc As String
      Dim DWGNo As String
        For j = 1 To ColCount
            If InStr(DelivWS.Cells(1, j).Value, strPhase) > 0 Then
                If InStr(DelivWS.Cells(i, j).Value, "x") > 0 Or InStr(DelivWS.Cells(i, j).Value, "o") > 0 Then
                    WBSCode = DelivWS.Range("D" & i).Value
                    WBSCodeDesc = DelivWS.Range("E" & i).Value
                    DWGNo = DelivWS.Range("G" & i).Value
                    If Not deliverables.Exists(strPhase) Then
                        deliverables.Add strPhase, CreateObject("Scripting.Dictionary")
                    End If ''// end primer populate for dictionary
                    If Not deliverables(strPhase).Exists(WBSCode) Then
                        deliverables(strPhase).Add WBSCode, Array(WBSCodeDesc, DWGNo)
                    End If
                End If ''// end check if column is marked for that phase
            End If ''// Finding Size Column
        Next j ''// iterating over columns
      Next i ''// Iterating along used rows (except titles)
       
      Set PopulateDictionary = deliverables
    End If ''// End strPhase!= Null check
End Function ''// End PopulateDictionary
''//--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x--x

Private Function UpdateProjNums(strInput, strProjNum)
    If strProjNum <> "" Then
        Dim regEx, Matches As Object
        If strInput <> "" Then
            Set regEx = CreateObject("vbscript.regexp")
            regEx.Pattern = "^[a-z]{5}"
            regEx.Global = True
            regEx.IgnoreCase = True
            Set Matches = regEx.Execute(strInput)
            If Not Matches Is Nothing Then
            If Matches.Count > 0 Then
                UpdateProjNums = regEx.Replace(strInput, strProjNum)
            End If
            End If
        End If  ''// End check for dwg number population
     Else ''// Project Number is NOT populated
        UpdateProjNums = strInput
    End If ''//{Project Number Is populated}

End Function
Private Sub cboProjectPhase_Change()
    'Show checkboxes based on selected project phase
    Dim selectedPhase As String
    selectedPhase = cboProjectPhase.Value
    Dim numCheckboxes As Integer
    Dim summaryText As String
    Dim dummy As String
    Dim i As Integer
    i = 1
    Dim deliverables As New Scripting.Dictionary  ''//empty dictionary

    Select Case selectedPhase
        Case "FEL1"
            numCheckboxes = 1
            summaryText = "Feasibility - This stage involves a more detailed evaluation of the project to determine its feasibility. A feasibility study is typically conducted to identify and evaluate alternative solutions and to identify the preferred option for the project."
            framePhase.Caption = "FEL1 :: Feasibility"
            framePhase.Visible = True
            framePhase.ZOrder msoSendToBack
            Set deliverables = PopulateDictionary(selectedPhase)

        Case "FEL2"
            numCheckboxes = 2
            summaryText = "Conceptual Design - In this stage, the preferred solution identified in the feasibility study is further developed into a conceptual design. The goal is to develop a high-level understanding of the project scope, cost, and schedule." & vbNewLine & "In dit stadium wordt de in de haalbaarheidsstudie vastgestelde voorkeursoplossing verder ontwikkeld tot een conceptueel ontwerp. Het doel is om op hoog niveau inzicht te krijgen in de omvang, de kosten en het tijdschema van het project."
            framePhase.Caption = "FEL2 :: Conceptual Design"
            framePhase.Visible = True
            framePhase.ZOrder msoSendToBack
            Set deliverables = PopulateDictionary(selectedPhase)

        Case "FEL3"
            numCheckboxes = 3
            summaryText = "Front End Engineering Design (FEED) - This stage involves the development of a detailed design for the project. The FEED typically includes the development of engineering specifications, project plans, and a detailed cost estimate."
            framePhase.Caption = "FEL3 :: Front End Engineering Design (FEED)"
            framePhase.Visible = True
            framePhase.ZOrder msoSendToBack
            Set deliverables = PopulateDictionary(selectedPhase)

        Case "FEL4"
            numCheckboxes = 0
            summaryText = "Execution/Implementation - This stage involves the execution of the project based on the detailed design developed in FEL 3. The goal is to ensure that the project is completed on time, within budget, and to the required quality standards."
            framePhase.Caption = "FEL4 :: Execution/Implementation"
            framePhase.Visible = True
            framePhase.ZOrder msoSendToBack
            Set deliverables = PopulateDictionary(selectedPhase)

        Case Else
           ' numCheckboxes = 0
            'summaryText = "Null"
            'framePhase.Caption = ""
            'framePhase.Visible = False
            'framePhase.ZOrder msoSendToBack
    End Select

    ' Show the appropriate number of checkboxes
    'framePhase.ZOrder = 0
    'Stop
    ProjectNumber = UserForm1.txtProjectNumber.Value
    If Not deliverables Is Nothing Then
     If selectedPhase <> "" Then
        Me.Controls.Add "Forms.Label.1", "LabelExplanation", True
        Me.Controls("LabelExplanation").Left = 22
        Me.Controls("LabelExplanation").Top = 145
        Me.Controls("LabelExplanation").Caption = "Please Select/UnSelect Applicable Phase Deliverables for your project;"
        Me.Controls("LabelExplanation").Width = 250
        Me.Controls("LabelExplanation").Visible = True
        Dim wbsKey As Variant
        With Me.DeliverableListBox
            .ColumnCount = 3
            .ColumnWidths = "55;225;65"
            .ColumnHeads = True
            .AddItem
            .List(0, 0) = "WBS"
            .List(0, 1) = "Deliverable Description"
            .List(0, 2) = "Drawing Number"
        For Each wbsKey In deliverables(selectedPhase).Keys
            If wbsKey <> "" Then
                    .AddItem
                    .List(i, 0) = CStr(wbsKey)
                    .List(i, 1) = deliverables(selectedPhase)(wbsKey)(0)
                    .List(i, 2) = UpdateProjNums(deliverables(selectedPhase)(wbsKey)(1), ProjectNumber)
            i = i + 1
         End If ''// end check for empty WBS codes
        Next wbsKey
       End With
     End If ''// Phase is Not Nothing Check
    End If ''// Dictionary != null check
    
    ' Show the project phase summary
    Me.phaseSummaryTextBox.Enabled = True
    Me.phaseSummaryTextBox.Visible = True
    Me.phaseSummaryTextBox.Value = summaryText

    ' Wrap the text in the summary textbox
    Dim txtWidth As Double
    Dim txtHeight As Double
    txtWidth = phaseSummaryTextBox.Width
    txtHeight = phaseSummaryTextBox.Height
    phaseSummaryTextBox.AutoSize = False

    phaseSummaryTextBox.Height = 27
    phaseSummaryTextBox.Width = 435
    
    deliverables.RemoveAll
    
End Sub

Private Sub btnSubmit_Click()
    'Close user form after submission
    Me.Hide
End Sub

' //TODO - IMPLEMENT RESET BUTTON TO REDO ESTIMATE ON THE SAME SHEET WITHOUT HAVING TO RE-RUN THE PROGRAM
' // The button should clear the phase selections and checkboxes, but should not remove the Project general information (Project no., Client Name, Project Description)
