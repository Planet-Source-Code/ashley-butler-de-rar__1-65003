Attribute VB_Name = "ShowSaveAs"
      Option Explicit

       Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

       Private Type OPENFILENAME
         lStructSize As Long
         hwndOwner As Long
         hInstance As Long
         lpstrFilter As String
         lpstrCustomFilter As String
         nMaxCustFilter As Long
         nFilterIndex As Long
         lpstrFile As String
         nMaxFile As Long
         lpstrFileTitle As String
         nMaxFileTitle As Long
         lpstrInitialDir As String
         lpstrTitle As String
         flags As Long
         nFileOffset As Integer
         nFileExtension As Integer
         lpstrDefExt As String
         lCustData As Long
         lpfnHook As Long
         lpTemplateName As String
       End Type

Public Function ShowOpenDialog(ByVal TitleOfDialog As String, ByVal FilterText As String, ByVal FilterFileType As String, Form As Form, ByVal DefaultLocation As String) As String
         Dim OpenFile As OPENFILENAME
         Dim lReturn As Long
         Dim sFilter As String
         OpenFile.lStructSize = Len(OpenFile)
         OpenFile.hwndOwner = Form.hWnd
         OpenFile.hInstance = App.hInstance
         sFilter = FilterText & Chr(0) & FilterFileType & Chr(0)
         OpenFile.lpstrFilter = sFilter
         OpenFile.nFilterIndex = 1
         OpenFile.lpstrFile = String(257, 0)
         OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
         OpenFile.lpstrFileTitle = OpenFile.lpstrFile
         OpenFile.nMaxFileTitle = OpenFile.nMaxFile
         OpenFile.lpstrInitialDir = OpenFile.lpstrInitialDir 'DefaultLocation
         OpenFile.lpstrTitle = TitleOfDialog
         OpenFile.flags = 0
         lReturn = GetOpenFileName(OpenFile)
         If lReturn = 0 Then
            ShowOpenDialog = ""
         Else
            ShowOpenDialog = Trim(OpenFile.lpstrFile)
         End If
       End Function



