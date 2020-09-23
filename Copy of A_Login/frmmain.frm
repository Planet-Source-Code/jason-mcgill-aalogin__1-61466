VERSION 5.00
Object = "{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form Main 
   BackColor       =   &H80000009&
   Caption         =   "Test Page"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4500
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin osenxpsuite2005.OsenXPButton cmdLogout 
      Default         =   -1  'True
      Height          =   1935
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3413
      Caption         =   "&Logout"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frmmain.frx":0442
      PICN            =   "frmmain.frx":045E
      UMCOL           =   -1  'True
      PICPOS          =   2
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Test Page.  Replace with your program."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   9015
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogout_Click()
'Close form
Dim strValue As String
strValue = MsgBox("Are you sure you want to logout?", vbQuestion + vbYesNo, "Logout?")
If strValue = vbYes Then Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Open login form
frmlogin.Show
End Sub

