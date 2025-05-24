Sub JTCRenameFiles()
Dim JTCDirectory As String
Dim JTCFile As String
Dim JTCRow As Long
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
If .Show = -1 Then
    JTCDirectory = .SelectedItems(1)
    JTCFile = Dir(JTCDirectory & Application.PathSeparator & "*")
    Do Until JTCFile = ""
        JTCRow = 0
        On Error Resume Next
        JTCRow = Application.Match(JTCFile, Range("A:A"), 0)
        If JTCRow > 0 Then
            Name JTCDirectory & Application.PathSeparator & JTCFile As _
            JTCDirectory & Application.PathSeparator & Cells(JTCRow, "B").Value
        End If
        JTCFile = Dir
    Loop
End If
End With
End Sub
