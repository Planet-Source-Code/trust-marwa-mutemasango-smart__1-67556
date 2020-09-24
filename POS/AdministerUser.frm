VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAdministerUser 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Administration"
   ClientHeight    =   7005
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AdministerUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   5760
   Begin VB.Frame FrameUsers 
      BackColor       =   &H80000014&
      Caption         =   "Users"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.Frame Frame2 
         BackColor       =   &H80000014&
         Caption         =   "Status"
         Height          =   1215
         Left            =   3360
         TabIndex        =   22
         Top             =   3360
         Width           =   1935
         Begin VB.CheckBox chkLockedOut 
            BackColor       =   &H80000014&
            Caption         =   "Locked Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkAccEnabled 
            BackColor       =   &H80000014&
            Caption         =   "Account Enabled"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000014&
         Caption         =   "Access List (Tick to grant access)"
         Height          =   2895
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   3135
         Begin VB.CheckBox chkAccessModules7 
            BackColor       =   &H80000014&
            Caption         =   "Customise"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2520
            Width           =   2415
         End
         Begin VB.CheckBox chkAccessModules6 
            BackColor       =   &H80000014&
            Caption         =   "Stock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   2160
            Width           =   2415
         End
         Begin VB.CheckBox chkAccessModules1 
            BackColor       =   &H80000014&
            Caption         =   "User Administration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox chkAccessModules2 
            BackColor       =   &H80000014&
            Caption         =   "Charges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   2415
         End
         Begin VB.CheckBox chkAccessModules3 
            BackColor       =   &H80000014&
            Caption         =   "Bookings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CheckBox chkAccessModules4 
            BackColor       =   &H80000014&
            Caption         =   "Costing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   2415
         End
         Begin VB.CheckBox chkAccessModules5 
            BackColor       =   &H80000014&
            Caption         =   "Reports"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1800
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000014&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   5175
         Begin VB.CommandButton cmdPassword 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2040
            Picture         =   "AdministerUser.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   960
            Width           =   350
         End
         Begin VB.TextBox txtPwd 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtConPwd 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "New Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.TextBox txtEmpName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtEmpPosition 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox txtUserID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3840
         Picture         =   "AdministerUser.frx":0941
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4800
         Width           =   350
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3360
         Picture         =   "AdministerUser.frx":0AA7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4800
         Width           =   350
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame FrameLogFile 
      BackColor       =   &H80000014&
      Caption         =   "Log File"
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5535
      Begin MSComctlLib.ListView lvwlogfile 
         Height          =   5895
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   10398
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imglRunSearch"
         SmallIcons      =   "imglRunSearch"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time In"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time Out"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label lblLogFile 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Log File"
      ForeColor       =   &H00996533&
      Height          =   255
      Left            =   1200
      TabIndex        =   30
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblUsers 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      ForeColor       =   &H00996533&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmAdministerUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
On Error GoTo Err_Routine
Dim SQL As String
Dim NuMRS As New ADODB.Recordset

 SQL = "Select * from UserDetails where UserID = '" & txtUserID.Text & "'"
 NuMRS.Open SQL, OldDB
If NuMRS.EOF = True Then
 SQL = "Insert into UserDetails (UserID,FullName,JobTitle,Department,[password],[LockedOut],AcEnabled,UserAdmin,Sales,EmpDetails,Costing,Reports,StockControl,Customise) values ('" & txtUserID.Text & "', '" & txtEmpName.Text & "', '" & txtEmpPosition.Text & "' , '" & txtDepartment & "', ''," & ""
 
 If chkLockedOut.Value = vbChecked Then
  SQL = SQL & "True,"
 Else
  SQL = SQL & "False,"
 End If
 
 If chkAccEnabled.Value = vbChecked Then
  SQL = SQL & "True,"
 Else
  SQL = SQL & "False,"
 End If
 
 If chkAccessModules1.Value = vbChecked Then
  SQL = SQL & "True,"
 Else
  SQL = SQL & "False,"
 End If
 
 If chkAccessModules2.Value = vbChecked Then
  SQL = SQL & "True,"
 Else
  SQL = SQL & "False,"
 End If
 
 If chkAccessModules3.Value = vbChecked Then
  SQL = SQL & "True,"
 Else
  SQL = SQL & "False,"
 End If
 
 If chkAccessModules4.Value = vbChecked Then
  SQL = SQL & "True,"
 Else
  SQL = SQL & "False,"
 End If
 
 If chkAccessModules5.Value = vbChecked Then
  SQL = SQL & "True,"
 Else
  SQL = SQL & "False,"
 End If
 
 If chkAccessModules6.Value = vbChecked Then
  SQL = SQL & "True,"
 Else
  SQL = SQL & "False,"
 End If
 
 If chkAccessModules7.Value = vbChecked Then
  SQL = SQL & "True)"
 Else
  SQL = SQL & "False)"
 End If
 
 OldDB.Execute SQL
Else
 
 SQL = "Update UserDetails set "
 SQL = SQL & "FullName = '" & txtEmpName.Text & "',"
 SQL = SQL & "JobTitle = '" & txtEmpPosition.Text & "',"
 SQL = SQL & "Department = '" & txtDepartment.Text & "',"
  
 If chkLockedOut.Value = vbChecked Then
  SQL = SQL & "[LockedOut] = True,"
 Else
  SQL = SQL & "[LockedOut] = False,"
 End If
 
 If chkAccEnabled.Value = vbChecked Then
  SQL = SQL & "AcEnabled = True,"
 Else
  SQL = SQL & "AcEnabled = False,"
 End If
 
  If chkAccessModules1.Value = vbChecked Then
  SQL = SQL & "UserAdmin = True,"
 Else
  SQL = SQL & "UserAdmin = False,"
 End If
 
 If chkAccessModules2.Value = vbChecked Then
  SQL = SQL & "Sales = True,"
 Else
  SQL = SQL & "Sales = False,"
 End If
 
 If chkAccessModules3.Value = vbChecked Then
  SQL = SQL & "EmpDetails = True,"
 Else
  SQL = SQL & "EmpDetails = False,"
 End If
 
 If chkAccessModules4.Value = vbChecked Then
  SQL = SQL & "Costing = True,"
 Else
  SQL = SQL & "Costing = False,"
 End If
 
  If chkAccessModules5.Value = vbChecked Then
  SQL = SQL & "Reports = True,"
 Else
  SQL = SQL & "Reports = False,"
 End If
 
 If chkAccessModules6.Value = vbChecked Then
  SQL = SQL & "StockControl = True,"
 Else
  SQL = SQL & "StockControl = False,"
 End If
  
 If chkAccessModules7.Value = vbChecked Then
  SQL = SQL & "Customise = True"
 Else
  SQL = SQL & "Customise = False"
 End If
 
 SQL = SQL & " where UserID = '" & txtUserID.Text & "'"
 
 OldDB.Execute SQL
End If
 
 cmdAdd.Enabled = False
 Form_Load
 
Exit_Routine:
Exit Sub
Err_Routine:
MsgBox Err.Description, , App.Title
Resume Exit_Routine
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err_Routine
Dim SQL As String

  SQL = "Delete from UserDetails where UserID = '" & txtUserID.Text & "'"
  OldDB.Execute SQL
  MsgBox "User Deleted", vbOKOnly, App.Title
  cmdDelete.Enabled = False
Form_Load
Exit_Routine:
Exit Sub
Err_Routine:
MsgBox Err.Description, , App.Title
Resume Exit_Routine
End Sub


Private Sub cmdPassword_Click()
On Error GoTo Err_Routine

If txtPwd.Text <> txtConPwd.Text Then
 MsgBox "Passwords do not match", vbOKOnly, App.Title
 Exit Sub
End If

OldDB.Execute "Update UserDetails set [password] = '" & txtConPwd.Text & "' where UserID = '" & txtUserID.Text & "'"
MsgBox "Password has been succesfully changed", vbOKOnly, App.Title

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim lstmain1 As ListItem
Dim LogFileRS As New ADODB.Recordset

LogFileRS.Open "LogFile", OldDB
  
While Not LogFileRS.EOF
 Set lstmain1 = lvwlogfile.ListItems.Add(, , LogFileRS!UserName)
 lstmain1.SubItems(1) = LogFileRS!DateIn
 lstmain1.SubItems(2) = LogFileRS!TimeIn
 lstmain1.SubItems(3) = LogFileRS!TimeOut
 LogFileRS.MoveNext
Wend

cmdPassword.Enabled = False

LoadCbo txtUserID, "Select Distinct UserID from UserDetails"

End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUsers.ForeColor = &H996533
lblLogFile.ForeColor = &H996533
End Sub

Private Sub lblLogFile_Click()
FrameUsers.Visible = False
FrameLogFile.Visible = True
End Sub

Private Sub lblLogFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogFile.ForeColor = &H80FF&
End Sub

Private Sub lblUsers_Click()
FrameUsers.Visible = True
FrameLogFile.Visible = False
End Sub

Private Sub lblUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUsers.ForeColor = &H80FF&
End Sub

Private Sub txtUserID_Change()
 cmdAdd.Enabled = True
 cmdDelete.Enabled = True
 cmdPassword.Enabled = False
End Sub

Private Sub txtUserID_Click()
Dim RS As New ADODB.Recordset

RS.Open "Select * from UserDetails where UserID = '" & txtUserID.Text & "'", OldDB

If RS.EOF = True Then Exit Sub

txtUserID.Text = RS!UserID
If Not IsNull(RS!FullName) Then txtEmpName.Text = RS!FullName
If Not IsNull(RS!Department) Then txtDepartment.Text = RS!Department
If Not IsNull(RS!JobTitle) Then txtEmpPosition.Text = RS!JobTitle


If RS!AcEnabled = True Then
chkAccEnabled.Value = vbChecked
Else
chkAccEnabled.Value = vbUnchecked
End If

If RS!LockedOut = True Then
chkLockedOut.Value = vbChecked
Else
chkLockedOut.Value = vbUnchecked
End If

If RS!UserAdmin = True Then
chkAccessModules1.Value = vbChecked
Else
chkAccessModules1.Value = vbUnchecked
End If

If RS!Sales = True Then
chkAccessModules2.Value = vbChecked
Else
chkAccessModules2.Value = vbUnchecked
End If

If RS!EmpDetails = True Then
chkAccessModules3.Value = vbChecked
Else
chkAccessModules3.Value = vbUnchecked
End If

If RS!Costing = True Then
chkAccessModules4.Value = vbChecked
Else
chkAccessModules4.Value = vbUnchecked
End If

If RS!Reports = True Then
chkAccessModules5.Value = vbChecked
Else
chkAccessModules5.Value = vbUnchecked
End If

If RS!StockControl = True Then
chkAccessModules6.Value = vbChecked
Else
chkAccessModules6.Value = vbUnchecked
End If

If RS!Customise = True Then
chkAccessModules7.Value = vbChecked
Else
chkAccessModules7.Value = vbUnchecked
End If

cmdPassword.Enabled = True
cmdAdd.Enabled = True
End Sub
