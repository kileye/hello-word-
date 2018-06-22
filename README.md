# hello-word-
my first rep




Sub Test()

Dim iPath As String, i As Long

Dim t

t = Timer

With Application.FileDialog(msoFileDialogFolderPicker)

.Title = "请选择要查找的文件夹"
 .InitialFileName = ThisWorkbook.Path
'If .Show Then

iPath = ThisWorkbook.Path ' ‘.SelectedItems(1)

'End If

End With

If iPath = "False" Or Len(iPath) = 0 Then Exit Sub

i = 4

Call GetFolderFile(iPath, i)

'MsgBox Timer - t

MsgBox "文件名链接获取完毕。", vbOKOnly, "提示"

End Sub

Private Sub GetFolderFile(ByVal nPath As String, ByRef iCount As Long)

Dim iFileSys

'Dim iFile As Files, gFile As File

'Dim iFolder As Folder, sFolder As Folders, nFolder As Folder

Set iFileSys = CreateObject("Scripting.FileSystemObject")

Set iFolder = iFileSys.GetFolder(nPath)

Set sFolder = iFolder.SubFolders

Set iFile = iFolder.Files

With ActiveSheet
.Cells(3, 1) = "序号"
.Cells(3, 2) = "文件名"
.Cells(3, 3) = "链接"
For Each gFile In iFile
    .Cells(iCount, 1) = iCount - 3
    .Cells(iCount, 2) = gFile.Name
  .Hyperlinks.Add anchor:=.Cells(iCount, 3), Address:=gFile.Name, TextToDisplay:=gFile.Name

iCount = iCount + 1

Next

End With

'递归遍历所有子文件夹

'‘For Each nFolder In sFolder

'Call GetFolderFile(nFolder.Path, iCount)

'Next

End Sub
