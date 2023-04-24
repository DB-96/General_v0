Sub Main()
    Dim searchString As String
    Dim folderPath As String
    Dim foundFiles As String
    
    ' Set the search string and folder path
    searchString = InputBox("What cable are you looking for?","Enter Cable Tag")
    folderPath = SelectFolder()
    
    ' Call the FindStringInDWGFiles function to search for the string in .dwg files
    foundFiles = FindStringInDWGFiles(searchString, folderPath)
    
    ' Display the list of found filenames
    MsgBox "The string was found in the following files: " & vbCrLf & foundFiles
End Sub

Function SelectFolder() As String
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''///
''/// Function Name ::: SelectFolder ()
''/// Description   ::: Prompts user to select and Folder where .dwg files are stored
''/// Variables     ::: N/A
''/// Date          ::: 23/09/2022  11:36:38
''/// Location      ::: C:\Users\debarros\OneDrive - KH Engineering\Documenten\Scripts\ExcelTOElecDesAccessImporter.vb
''/// Author        ::: 55a618
''///
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim fd As Office.FileDialog
    Dim strFile As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        fd.InitialFileName = "K:\BU3\Proj-Open\Hexion\69214-000\147-Instrument\1472-Diagrams\DETAIL ENGG\7803-Wiring\Info_Copy"
   
    If fd.Show = -1 Then
        xFdItem = fd.SelectedItems(1) & Application.PathSeparator
        Filename = xFdItem
        SelectFolder = Filename
    ' If Not fd Is Nothing Then
    '    With fd
    '         .Filters.Clear
    '         .Title = "Choose the Source :: "
    '         .AllowMultiSelect = False
    '         .InitialFileName = "K:\BU3\Proj-Open\Hexion\69214-000\147-Instrument\1472-Diagrams\DETAIL ENGG\7803-Wiring\Info_Copy"  ''// TODO:: move this initial file name to project root directory/server location where database will actually be running
    '         If .Show = True Then
    '             If .SelectedItems(1) & Application.PathSeparator <> "" Then
    '                 strPrompt = "You have selected the following folder: " & .SelectedItems(1)& Application.PathSeparator & vbCrLf & "Would you like to continue?"  ''// Optional
    '                 resp = MsgBox(strPrompt, vbSystemModal, "Warning")                                                                  ''// Optional
    '                 ''// Optional :: User to double check database that has been chosen within message box
    '                 SelectFolder = .SelectedItems(1)  & Application.PathSeparator
    '             End If ''// Ensure the selection string is not empty
    '         End If
    '     End With
    End If   ''// End fd != Null check
End Function ''// End SelectFolder


Function FindStringInDWGFiles(searchString As String, folderPath As String) As String
    Dim fso As Object
    Dim app As AcadApplication
    Dim folder As Object
    Dim file As Object
    Dim filePath As String
    Dim foundInFiles As String
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create an AutoCAD.Application object
     'Check if AutoCAD application is open. If is not opened create a new instance and make it visible.

     On Error Resume Next
      Set app = GetObject(, "AutoCAD.Application")
     On Error GoTo 0
      If app Is Nothing Then
        Set app = New AcadApplication
        app.Visible = True
      End If


    
    ' Get the folder object
    Set folder = fso.GetFolder(folderPath)
    
    ' Loop through all files in the folder
    For Each file In folder.Files
        ' Check if the file is a .dwg file
        If Right(file.Name, 4) = ".dwg" Then
            ' Get the full path of the file
            filePath = folderPath & file.Name
            
            ' Open the file in AutoCAD
            app.Documents.Open filePath
            
            ' Search for the input string in the file
           
            For Each Entity In app.ActiveDocument.ModelSpace
                If Entity.ObjectName = "AcDbText" Then
                    If Entity.TextString = searchString Then
                        foundInFiles = foundInFiles & file.Name & vbCrLf
                        
                    End If
                End If
            Next
            ' If app.ActiveDocument.ModelSpace.Find(searchString) Then
            '     ' Add the filename to the list of found files
            '     foundInFiles = foundInFiles & file.Name & vbCrLf
            ' End If
            
            ' Close the file in AutoCAD
            app.Documents.Close
        End If
    Next
    
    ' Return the list of found filenames
    FindStringInDWGFiles = foundInFiles
End Function


