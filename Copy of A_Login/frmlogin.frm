VERSION 5.00
Object = "{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frmlogin 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin osenxpsuite2005.OsenXPButton cmdChange 
      Height          =   975
      Left            =   2640
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1720
      Caption         =   "&Change"
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
      MICON           =   "frmlogin.frx":0442
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPListBox lstScheme 
      Height          =   930
      Left            =   1320
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1640
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSelected    =   16576
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      AllowEdit       =   0   'False
      WordWrap        =   0   'False
      ItemHeight      =   20
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      XPAlphaBlend    =   0   'False
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      HeaderCaption   =   "OsenXPListBox1"
   End
   Begin osenxpsuite2005.OsenXPButton cmdScheme 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&Scheme"
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
      MICON           =   "frmlogin.frx":045E
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPButton cmdok 
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&OK"
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
      MICON           =   "frmlogin.frx":047A
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPButton cmdCancel 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
      MICON           =   "frmlogin.frx":0496
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPButton cmdRegister 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Register"
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
      MICON           =   "frmlogin.frx":04B2
      UMCOL           =   -1  'True
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPTextBox txtpass 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      PasswordChar    =   "*"
   End
   Begin osenxpsuite2005.OsenXPTextBox txtname 
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "If you do not have a username and password type them in the spaces below and click register."
      Height          =   735
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmlogin.frx":04CE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Plese enter your username and password in the space provided below to login."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If you have any questions about how to use this program or ideas on how to
'better this program please contact me (Jason) at jaymcgill@gmail.com.
Option Explicit
Dim cn As New ADODB.Connection, strCNString As String
Dim rs As New ADODB.Recordset
Dim Txt As String

Private Sub cmdChange_Click()
'Change the Scheme of the form
    If lstScheme = "Default" Then txtname.ColorScheme = xpDefault
    If lstScheme = "Default" Then txtpass.ColorScheme = xpDefault
    If lstScheme = "Default" Then cmdRegister.ColorScheme = xpDefault
    If lstScheme = "Default" Then cmdCancel.ColorScheme = xpDefault
    If lstScheme = "Default" Then cmdok.ColorScheme = xpDefault
    If lstScheme = "Default" Then cmdScheme.ColorScheme = xpDefault
    If lstScheme = "XP Blue" Then txtpass.ColorScheme = xpBlue
    If lstScheme = "XP Blue" Then txtname.ColorScheme = xpBlue
    If lstScheme = "XP Blue" Then cmdRegister.ColorScheme = xpBlue
    If lstScheme = "XP Blue" Then cmdCancel.ColorScheme = xpBlue
    If lstScheme = "XP Blue" Then cmdok.ColorScheme = xpBlue
    If lstScheme = "XP Blue" Then cmdScheme.ColorScheme = xpBlue
    If lstScheme = "XP Olive" Then txtname.ColorScheme = xpOliveGreen
    If lstScheme = "XP Olive" Then txtpass.ColorScheme = xpOliveGreen
    If lstScheme = "XP Olive" Then cmdRegister.ColorScheme = xpOliveGreen
    If lstScheme = "XP Olive" Then cmdCancel.ColorScheme = xpOliveGreen
    If lstScheme = "XP Olive" Then cmdok.ColorScheme = xpOliveGreen
    If lstScheme = "XP Olive" Then cmdScheme.ColorScheme = xpOliveGreen
    If lstScheme = "XP Silver" Then txtname.ColorScheme = xpSilver
    If lstScheme = "XP Silver" Then txtpass.ColorScheme = xpSilver
    If lstScheme = "XP Silver" Then cmdCancel.ColorScheme = xpSilver
    If lstScheme = "XP Silver" Then cmdRegister.ColorScheme = xpSilver
    If lstScheme = "XP Silver" Then cmdok.ColorScheme = xpSilver
    If lstScheme = "XP Silver" Then cmdScheme.ColorScheme = xpSilver
    lstScheme.Visible = False
    cmdChange.Visible = False
'Change the form caption to show the Scheme selected
    If lstScheme = "Default" Then frmlogin.Caption = "Login - Default Color Scheme"
    If lstScheme = "XP Blue" Then frmlogin.Caption = "Login - XP Blue Color Scheme"
    If lstScheme = "XP Olive" Then frmlogin.Caption = "Login - XP Olive Color Scheme"
    If lstScheme = "XP Silver" Then frmlogin.Caption = "Login - XP Silver Color Scheme"
End Sub

Private Sub cmdOK_Click()

On Error GoTo ErrHandler
'Connect to database
strCNString = "Data Source=" & App.Path & "\dbpassword.mdb"
cn.Provider = "Microsoft Jet 4.0 OLE DB Provider"
cn.ConnectionString = strCNString
cn.Properties("Jet OLEDB:Database Password") = "jason"
cn.Open
'Open recordsource
With rs
   
         .Open "Select * from tblUsers where Username='" & txtname.Text & "' and Password='" & txtpass.Text & "'", cn, adOpenDynamic, adLockOptimistic
        'Check username and password
        If .EOF Then
            MsgBox "Access Denied...Please enter correct password!", vbOKOnly + vbCritical, "Security Login"
               txtname.Text = ""
               txtpass.Text = ""
               txtname.SetFocus
               cn.Close
        Else
           Txt = "" & " " & UCase$(txtname.Text) & ""
            MsgBox "Welcome!!!" & Txt, vbOKOnly + vbExclamation, "Security Login"
            cn.Close
            Unload Me
            Main.Show
            
        End If
    End With

     Exit Sub
     
ErrHandler:
MsgBox Err.Description, vbCritical, "Login"
cn.Close
End Sub

Private Sub cmdCancel_Click()
'Close form
Unload Me
End Sub

Private Sub cmdRegister_Click()
'Register a new username and password
On Error Resume Next
'Keep user from saving a blank username
If txtname.Text = "" Then GoTo message
'Connect to database
strCNString = "Data Source=" & App.Path & "\dbpassword.mdb"
cn.Provider = "Microsoft Jet 4.0 OLE DB Provider"
cn.ConnectionString = strCNString
cn.Properties("Jet OLEDB:Database Password") = "jason"
cn.Open
'Open recordsource and check for duplicate users
With rs
    .Open "Select * from tblUsers where Username='" & txtname.Text & "'", cn, adOpenDynamic, adLockOptimistic
    If Not .EOF Then
            MsgBox "Username in use.  Please choose again.", vbOKOnly + vbCritical, "Security Login"
               txtname.Text = ""
               txtpass.Text = ""
               txtname.SetFocus
               cn.Close
    'Ready recordsource for adding username and password
    Else
        .AddNew
        'Assign test boxes on form to their appropriate field in the recordsource
        rs(0) = txtname.Text
        rs(1) = txtpass.Text
        'Save record
        .Save
        'Close connections to database
        cn.Close
        .Close
        MsgBox "User Name and Password Created.", vbInformation, "Confirmation"
    End If
End With
Exit Sub
message:
    MsgBox "You must enter a User Name and Password.", vbCritical, "Error"
End Sub

Private Sub cmdScheme_Click()
'Show the list box with the Schemes available and also the Change button
lstScheme.Visible = True
cmdChange.Visible = True
End Sub

Private Sub Form_Load()
'Disable Register button until data is entered into both text boxes
cmdRegister.Enabled = False
'Add Schemes to the list box
lstScheme.AddItem ("Default")
lstScheme.AddItem ("XP Blue")
lstScheme.AddItem ("XP Olive")
lstScheme.AddItem ("XP Silver")
End Sub

Private Sub txtpass_Change()
'Enable Register button to allow user to save record
cmdRegister.Enabled = True
End Sub
