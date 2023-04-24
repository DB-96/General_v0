Sub Test()
    Dim I As Long
    Dim xRg As Range
    Dim xStr As String
    Dim xFd As FileDialog
    Dim xFdItem As Variant
    Dim xFileName As String
    Dim xFileNum As Long
    Dim RegExp As Object
    Set xFd = Application.FileDialog(msoFileDialogFolderPicker)
    If xFd.Show = -1 Then
        xFdItem = xFd.SelectedItems(1) & Application.PathSeparator
        xFileName = Dir(xFdItem & "*.pdf", vbDirectory)
        Set xRg = Range("A1")
        Range("A:B").ClearContents
        Range("A1:B1").Font.Bold = True
        xRg = "File Name"
        xRg.Offset(0, 1) = "Pages"
        I = 2
        xStr = ""
        Do While xFileName <> ""
            Cells(I, 1) = xFileName
            Set RegExp = CreateObject("VBscript.RegExp")
            RegExp.Global = True
            RegExp.Pattern = "/Type\s*/Page[^s]"
            xFileNum = FreeFile
            Open (xFdItem & xFileName) For Binary As #xFileNum
                xStr = Space(LOF(xFileNum))
                Get #xFileNum, , xStr
            Close #xFileNum
            Cells(I, 2) = RegExp.Execute(xStr).Count
            I = I + 1
            xFileName = Dir
        Loop
        Columns("A:B").AutoFit
    End If
End Sub