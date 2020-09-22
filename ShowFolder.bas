Attribute VB_Name = "ShowFolder"
'----------------------------------------------------------------
'
'                Browse for folders in VB5
'
'              written by D. Rijmenants 2004
'
'----------------------------------------------------------------
'
' This module enables you to get use the browse dialog to select
' a folder in vb5. Only one functionis required.  As you call
' the function, the browse dialog pops up. Easy to apply !
'
' return = BrowseFolder(Title, MyForm)
'
' Where:
'
' Title  (string) is the title you want to display on the dialog
' MyForm (form) is the form on wich you call the dialog
' return (string) the path of the selected folder after pressing OK
'
' Note: if cancel is selected, the dialog will return
'       an empty string !
'
'
' That's all folks...
'
' Comments or suggestions are most welcome at
' mail: dr.defcom@telenet.be
'
'----------------------------------------------------------------
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Public Function BrowseFolder(ByVal aTitle As String, ByVal aForm As Form) As String
Dim bInfo As BROWSEINFO
Dim rtn&, pidl&, Path$, pos%
Dim BrowsePath As String
bInfo.hOwner = aForm.hWnd
bInfo.lpszTitle = aTitle
'the type of folder(s) to return
'bInfo.ulFlags = &H1
bInfo.ulFlags = &H40 '&HFF & &H40
'show the dialog box
pidl& = SHBrowseForFolder(bInfo)
'set the maximum characters
Path = Space(512)
'get the selected path
t = SHGetPathFromIDList(ByVal pidl&, ByVal Path)
pos% = InStr(Path$, Chr$(0)) 'extracts the path from the string
'set the extracted path to SpecIn
BrowseFolder = Left(Path$, pos - 1)
'clean up the path string
If Right$(BrowseFolder, 1) = "\" Then
    BrowseFolder = BrowseFolder
    Else
    BrowseFolder = BrowseFolder + "\"
End If
If Right(BrowseFolder, 2) = "\\" Then BrowseFolder = Left(BrowseFolder, Len(BrowseFolder) - 1)
If BrowseFolder = "\" Then BrowseFolder = ""
End Function



