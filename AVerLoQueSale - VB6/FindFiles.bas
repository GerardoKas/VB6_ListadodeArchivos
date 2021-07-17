Attribute VB_Name = "mFindFiles"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3

Public fs As FileSystemObject

Sub FindPreparando()
With Form1.lv1
    .AllowColumnReorder = True
    '.Checkboxes = True
    .FullRowSelect = True
    '.HoverSelection = True
    .MultiSelect = True
    .ColumnHeaders.Add , , "name"
    .ColumnHeaders.Add , , "path"
    .ColumnHeaders.Add , , "ext"
    .ColumnHeaders.Add , , "fecha"
    .ColumnHeaders.Add , , "Tamaño"
End With
End Sub
Function itsOk(name, path)
itsOk = True

End Function


Sub FindFiles(path As String)
    Dim fld As Scripting.Folder
    Dim fl As Scripting.File
    Dim flds As Scripting.Folders
    Dim fls As Scripting.Files
    Dim li As ListItem
    Dim lisu As ListSubItem
    Static oldDirKey As String
    If fs.FolderExists(path) = False Then MsgBox "No existe :" & path: Exit Sub
    
    Set fld = fs.GetFolder(path)
    Set flds = fld.SubFolders
    Set fls = fld.Files
    For Each fl In fls
        If itsOk(fl.name, fl.ParentFolder) Then
            Set li = Form1.lv1.ListItems.Add(, , fl.name)
            base = fs.GetParentFolderName(fl.path)
            li.SubItems(1) = base
            ext = fs.GetExtensionName(fl.name)
            li.SubItems(2) = ext
            fecha = FileDateTime(fl.path)
            li.SubItems(3) = fecha
            li.SubItems(4) = fl.Size
            'li.Checked = True
        End If
    Next
    
    For Each fld In flds
        DoEvents
        If doStop = True Then Exit Sub
        FindFiles (fld.path)
        DoEvents
    Next
End Sub
