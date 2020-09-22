VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSource 
   Caption         =   "De-Rar"
   ClientHeight    =   7290
   ClientLeft      =   690
   ClientTop       =   4155
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7290
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView 
      Height          =   3495
      Left            =   0
      TabIndex        =   12
      Top             =   3000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6165
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList"
      SmallIcons      =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton btnReload 
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Reload Archive data"
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton btnShowComment 
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Show Archive Comments"
      Top             =   6600
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   5520
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   7260
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtOutputPath 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   5175
   End
   Begin VB.CommandButton btnPasswordHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtOpenRar 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1320
      Width           =   5175
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imglstToolbar"
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglstToolbar 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   5400
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   5520
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   5640
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   5760
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   4
      Left            =   5880
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Extract to:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblpassword 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Archive:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image pictoolbar 
      Height          =   615
      Index           =   2
      Left            =   1800
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image pictoolbar 
      Height          =   615
      Index           =   1
      Left            =   1200
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image pictoolbar 
      Height          =   615
      Index           =   0
      Left            =   600
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer


Private Sub btnProcess_Click()
If CorruptFile = True Then
    MsgBox "Corrupt File", vbCritical + vbOKOnly, "Corrupt File"
    Exit Sub
End If

'/////////check input path

If txtOpenRar.Text <> "" Then
    If Mid(Me.txtOpenRar.Text, Len(Me.txtOpenRar.Text) - 3, 4) = ".rar" Then 'if the ext on fullpathtorar is .rar
        If PasswordStatus = True Then 'there is a password associated with file
            If txtPassword.Text = "" Then ' checks to see if user has entered a password
                MsgBox "Please supply a password first!", vbExclamation, "Password Protected"
                Exit Sub
            Else 'there is text in password field
               'do nothing
            End If 'end text in password field check
        End If 'end checking for password
    Else ' file ext is not rar
        MsgBox "You have bypassed the 'select only Rar Archive' function, you cheeky monkey!" & vbCrLf & vbCrLf & "But I am smart and knew you would try!", vbExclamation + vbOKOnly, "Not a Rar Archive"
        Exit Sub
    End If ' end detection of correct filetype
Else
    MsgBox "No archive selected!", vbExclamation, "File Error" 'there can't have been a file selected or a correct file selected
    Exit Sub
End If


'///////////check output path
If Me.txtOutputPath.Text = "" Then MsgBox "Select a place to extract to", vbExclamation + vbOKOnly, "Error": Exit Sub

'show progress form
Me.Hide
frmProgress.Show
        


End Sub

Private Sub cmdChangeDir_Click()
Me.txtOutputPath.Text = BrowseFolder("Select Folder for Output", Me)
End Sub

Private Sub btnReload_Click()
Call ListFiles_Click
End Sub

Private Sub btnShowComment_Click()
MsgBox ArchiveComment, vbOKOnly, "Archive Comment"
End Sub

'Private Sub cmdListFiles_Click()

Private Sub Form_Load()
Dim TheCommand As String
Set Colx = Me.ListView.ColumnHeaders.Add(, , "File Name")
Set Colx = Me.ListView.ColumnHeaders.Add(, , "Packed Size")
Set Colx = Me.ListView.ColumnHeaders.Add(, , "Unpacked Size")
Set Colx = Me.ListView.ColumnHeaders.Add(, , "File Date")
Set Colx = Me.ListView.ColumnHeaders.Add(, , "File CRC")
Set Colx = Me.ListView.ColumnHeaders.Add(, , "Path")


    'set pictures for the rar comment button
    Me.btnShowComment.Picture = LoadResPicture(109, vbResIcon)
    'Me.btnShowComment.DownPicture = LoadResPicture(109, vbResIcon)
    Me.btnShowComment.DisabledPicture = LoadResPicture(110, vbResIcon)
    
    'set the reload icons
    Me.btnReload.Picture = LoadResPicture(111, vbResIcon)
    Me.btnReload.DisabledPicture = LoadResPicture(112, vbResIcon)
    
    
    

    Me.txtPassword.Visible = False 'disable password field unless user wants a password to enter
    Me.lblpassword.Visible = False
    Me.btnPasswordHelp.Visible = False
    Me.btnShowComment.Enabled = False
    Me.btnReload.Enabled = False
    'txtPassword.BackColor = RGB(230, 230, 230) ' sets password field to grey

'load images from the resource file into the picure boxes, and then into the imagelist
    For i = 0 To 2
        Me.pictoolbar(i).Picture = LoadResPicture(106 + i, vbResIcon)
        Me.imglstToolbar.ListImages.Add , , Me.pictoolbar(i).Picture
    Next
    
'assign the image list to the toolbar
    Me.Toolbar.ImageList = Me.imglstToolbar
'populate the toolbar with buttons
    With Me.Toolbar.Buttons
        .Add 1, , "1) Select the Rar archive", , 1
        .Item(1).ToolTipText = "Select the Rar archive that you wish to extract the contents from"
        .Add 2, , "2) Select Destination", , 2
        .Item(2).ToolTipText = "Select the destination where you would like the contents of the Rar archive to be extracted to"
        .Add 3, , "3) Extract Archive", , 3
        .Item(3).ToolTipText = "Extract the contents of the Rar archive"
    End With

'sort the treeview out. same as above code really
For i = 0 To 3
    Me.Image1(i).Picture = LoadResPicture(101 + i, vbResIcon)
    Me.ImageList.ListImages.Add , , Me.Image1(i).Picture
Next

'the followin code allows a RAR archive to be opened without Rar support directly
If Command$ <> "" Then

    'check to see if its a rar archive
    'trim off the " at he begining and end of the command
    TheCommand = Right(Command$, Len(Command$) - 1)
    TheCommand = Left(TheCommand, Len(TheCommand) - 1)

    If Right(TheCommand, 4) = ".rar" Then
    'if it is then put the name of the file in the text box
        Me.txtOpenRar.Text = TheCommand
        'reset the form back to normal
        Call ResetForm
        'obtain number of files in the archive
        Call ListFiles_Click
    Else ' the file is not a rar archive
        MsgBox "Not a Rar archive you scoundrel!", vbOKOnly + vbExclamation, "Rar archives only"
    End If
End If


End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call txtOpenRar_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Form_Resize()
If Me.Width < 5800 Then
    Me.Width = 5800
End If
Me.ListView.Width = Me.Width - 110
Me.txtOpenRar.Width = Me.Width - 1200
Me.txtOutputPath.Width = Me.Width - 1200
Me.lblCredits.Left = Me.Width - 1150
End Sub

'damn memory leak wont survive this!!
Private Sub Form_Terminate()
'Unload FilesInArchive 'logical order
Unload Me 'always me last
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload FilesInArchive
Unload Me
End Sub


Private Sub lblCredits_Click()
MsgBox "Thanks to:" _
& vbCrLf _
& vbCrLf & "- WinRar for making the Unrar.dll file free to use" _
& vbCrLf & "- Leigh Bowers for the basis of the Unrar.dll code" _
& vbCrLf & "- Pedro Lamas for the code on how to list files using the unrar.dll" _
& vbCrLf & "- Rossini Enrico for his breakdown of the Unrar.dll code and file flags" _
& vbCrLf & "- Microsoft for the Open Dialog API code" _
& vbCrLf & "- D. Rijmenants for the Open Directory API" _
& vbCrLf & "- Mike Bouffler for his Icon Suite/Edit program, which all the icons are made from" _
& vbCrLf & "- FrozenPea of VbCity.com for showing me how to read Rar comments" _
& vbCrLf & "- Me, Ashley Butler for implementing and editing all the code and creating the GUI" _
, vbInformation + vbOKOnly, "Program Credits"

'http://edais.mvps.org/Tutorials/Graphics/GFXch5.html for bitshifting



End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button
    Case "1) Select the Rar archive"
        Call btnOpenRar_Click
    Case "2) Select Destination"
        Call cmdChangeDir_Click
    Case "3) Extract Archive"
        Call btnProcess_Click
End Select

End Sub

Private Sub txtOpenRar_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtOpenRar_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'allows a rar archive to be dropped on the open rar text box


'check to see if its a rar archive
If Right(Data.Files(1), 4) = ".rar" Then
    'if it is then put the name of the file in the text box
    Me.txtOpenRar.Text = Data.Files(1)
    'reset the form back to normal
    Call ResetForm
    'obtain number of files in the archive
    Call ListFiles_Click
Else ' the file is not a rar archive
    MsgBox "Not a Rar archive you scoundrel!", vbOKOnly + vbExclamation, "Rar archives only"
End If
End Sub

Private Sub txtOutputPath_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub btnPasswordHelp_Click()
MsgBox "If the file has a password and none is given, you will not be able to proceed" _
& vbCrLf & vbCrLf & "If an incorrect password is given, nothing will be extracted", vbInformation, "Password Help"
End Sub

Sub ResetForm()

'
'
'
'
'clear treeview
Set Lisx = Nothing
Me.ListView.ListItems.Clear

Me.btnShowComment.Enabled = False
TotalArchiveSize = 0 'clears archive size, allows proper reporting of file lenght
NumberOfFilesInArchive = 0 'clears the counter of files in archive
PasswordStatus = False ' resets password status

Me.txtPassword.Visible = False 'disables password field
Me.lblpassword.Visible = False
Me.btnPasswordHelp.Visible = False

'Me.txtPassword.BackColor = RGB(230, 230, 230) 'vbgrey doesn't exist so....
Me.txtPassword.Text = "" 'clear text in password field
Me.btnReload.Enabled = False
End Sub
Private Sub btnOpenRar_Click()
Call ResetForm


On Error GoTo ErrorHandler
'MsgBox "err handle is off"
txtOpenRar.Text = ShowOpenDialog("Select Your Rar archive", "Rar Archives (*.Rar)", "*.Rar", Me, App.Path)
'go to old error handler otherwise an error is created further on
If txtOpenRar.Text = "" Then Err.Number = 32755: GoTo ErrorHandler
Call ListFiles_Click


Exit Sub

ErrorHandler: 'error handler
Debug.Print Err.Number
If Err.Number = 32755 Then 'the cancel error code
    txtOpenRar.Text = ""
'    lblFilesInArchive.Caption = "?? files and ?? folders in Archive"
    '
    '
    '
    '
    'clear treeview
    'FilesInArchive.List1.Clear
ElseIf Err.Number = 53 Then
    MsgBox "Unrar.dll file was not found. Make sure it is in the same folder as this program is run from or the System32 directory", vbCritical + vbOKOnly, "File Not Found"
Else
    MsgBox "Some unknown error occured", vbCritical, "Unknown error"
End If

End Sub

Private Sub ListFiles_Click()

'doevents mean disable all buttons
'Me.btnOpenRar.Enabled = False
'Me.btnPasswordHelp.Enabled = False
'FilesInArchive.List1.Enabled = False
'Me.lblCredits.Enabled = False
'Me.btnProcess.Enabled = False
'Me.cmdChangeDir.Enabled = False

Set Lisx = Nothing
Me.ListView.ListItems.Clear
'frmsource.listview.listitems.Clear
'set the root to be the archive that was selected
'Set Lisx = frmSource.listview.Nodes.Add(, , "Main Archive", Me.txtOpenRar.Text, 3, 3)

Call RARExtract("ObtainList", Me.txtOpenRar, , Me.txtPassword.Text)

're-enable all buttons
'Me.btnOpenRar.Enabled = True
'Me.btnPasswordHelp.Enabled = True
'FilesInArchive.List1.Enabled = True
'Me.lblCredits.Enabled = True
'Me.btnProcess.Enabled = True
'Me.cmdChangeDir.Enabled = True
Debug.Print ArchiveComment
If ArchiveComment <> "" Then
    Me.btnShowComment.Enabled = True
    MsgBox ArchiveComment, vbOKOnly, "Archive Comment"
Else
    Me.btnShowComment.Enabled = False
End If
'lblFilesInArchive = NumberOfFilesInArchive & " file(s) and " & iFolderCount & " folder(s) in Archive"   'displays the number of files in a nice label

End Sub


