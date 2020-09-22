VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   Caption         =   "De-Raring"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5640
      Top             =   120
   End
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "STOP"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblFileNumber 
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
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblProcessingFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Processing File:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblProgress 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Progress:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSTOP_Click()
Debug.Print "clicked"
StopProcessing = True 'exit de-rar
'prevent button from being pressed again
Me.Enabled = False
End Sub

Private Sub Form_Load()
Me.Enabled = True 'just incase it isn't
Me.Left = frmSource.Left
Me.Top = frmSource.Top
'Timer1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmSource.Show
End Sub

Private Sub Timer1_Timer()
'stop timer from fireing again
Me.Timer1.Enabled = False
    Call RARExtract("Extract", frmSource.txtOpenRar.Text, frmSource.txtOutputPath.Text, frmSource.txtPassword.Text) 'runs extraction function

        If ExtractedObjects = 0 Then 'nothing was extracted
            MsgBox "Nothing was Extracted", vbExclamation + vbOKOnly, "Error"
        ElseIf StopProcessing = True Then
            MsgBox "Archive Processing was stopped early by the user", vbInformation + vbOKOnly, "Partially extracted"
        Else
            MsgBox "Extraction complete", vbInformation + vbOKOnly, "Success"
        End If
        Unload frmProgress


End Sub
