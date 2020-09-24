VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1605
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3375
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   948.287
   ScaleMode       =   0  'User
   ScaleWidth      =   3168.942
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000005&
      Height          =   350
      Left            =   2640
      Picture         =   "frmLogin.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Login"
      Top             =   1080
      Width           =   350
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A67D46&
      Height          =   345
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2445
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A67D46&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2445
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A67D46&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A67D46&
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1440
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdOK_Click()
On Error GoTo Err_Routine
Dim RS As New ADODB.Recordset


If Trim(txtPassword.Text) = "" Then
 MsgBox "Please Enter Password", vbOKOnly, App.Title
 Exit Sub
End If

Set OldDB = New Connection
OldDB.CursorLocation = adUseClient
OldDB.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\EM.mdb"

UserName = txtUserName.Text
RS.Open "select * from UserDetails where UserID = '" & txtUserName & "' and password = '" & txtPassword.Text & "'", OldDB

    'check for correct password
    If RS.EOF Then
      MsgBox "Invalid Password, try again!", , "Login"
      txtPassword.SetFocus
      SendKeys "{Home}+{End}"
      Exit Sub
    End If
 
    If RS!AcEnabled = True Then
      MsgBox "Acess Denied! Account Disabled", vbOKOnly, App.Title
      End
    End If
    
    If RS!LockedOut = True Then
      MsgBox "Acess Denied! Account Locked Out", vbOKOnly, App.Title
      End
    End If
            
    Load frmMain
    
    If RS!UserAdmin = False Then
         frmMain.Toolbar1.Buttons(6).Visible = False
     Else
         SuperUser = True
    End If
      
    If RS!Sales = False Then frmMain.Toolbar1.Buttons(1).Visible = False
        
    If RS!EmpDetails = False Then frmMain.Toolbar1.Buttons(5).Visible = False
        
    If RS!Reports = False Then frmMain.Toolbar1.Buttons(4).Visible = False
        
    If RS!Costing = False Then frmMain.Toolbar1.Buttons(2).Visible = False
            
    If RS!StockControl = False Then frmMain.Toolbar1.Buttons(3).Visible = False
        
    If RS!Customise = False Then frmMain.Toolbar1.Buttons(7).Visible = False
            
    LoginSucceeded = True
    Unload Me
    frmMain.Show
    
Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub Form_Load()
On Error Resume Next

End Sub







Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Len(txtPassword.Text) > 0 Then cmdOK_Click
End Sub
