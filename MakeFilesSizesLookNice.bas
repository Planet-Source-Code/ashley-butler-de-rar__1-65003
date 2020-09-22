Attribute VB_Name = "MakeFilesSizesLookNice"
Public Function FilesSize(ByVal TheFilesSize As String) As String
                

                
If TheFilesSize >= 1024# And TheFilesSize < 1048576# Then
        FilesSize = Round(TheFilesSize / 1024#, 2) & " KB"
ElseIf TheFilesSize >= 1048576# And TheFilesSize < 1073741824# Then
        FilesSize = Round(TheFilesSize / 1048576#, 2) & " MB"
ElseIf TheFilesSize >= 10737341824# Then
        FilesSize = Round(TheFilesSize / 10737341824#, 2) & " GB"
Else: FilesSize = TheFilesSize & " Bytes"
End If


End Function

