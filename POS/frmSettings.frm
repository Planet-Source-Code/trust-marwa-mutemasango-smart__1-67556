VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5400
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5250
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame10 
      BackColor       =   &H80000014&
      Caption         =   "Company Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   5055
      Begin VB.TextBox txtCoName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtTel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CommandButton cmdSaveCoInfo 
         Height          =   350
         Left            =   1920
         Picture         =   "frmSettings.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2520
         Width           =   350
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000014&
         Caption         =   "Company Name"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000014&
         Caption         =   "Telephone"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000014&
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000014&
         Caption         =   "Fax"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000014&
         Caption         =   "Physical Address"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000014&
      Caption         =   "Overhead Class's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5055
      Begin VB.CommandButton cmdDelOverheadClass 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   4560
         Picture         =   "frmSettings.frx":0941
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   350
      End
      Begin VB.CommandButton cmdSaveOverheadClass 
         Height          =   350
         Left            =   4080
         Picture         =   "frmSettings.frx":0AA7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   350
      End
      Begin VB.ComboBox cboOverheadTypes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00996533&
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000014&
      Caption         =   "Category's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cboCategoryTypes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00996533&
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton cmdSaveCategory 
         Height          =   350
         Left            =   4080
         Picture         =   "frmSettings.frx":0B1E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   350
      End
      Begin VB.CommandButton cmdDelCategory 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   4560
         Picture         =   "frmSettings.frx":0B95
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   350
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdSaveCategory_Click()
On Error GoTo Err_Routine

If cboCategoryTypes.Text = "" Then Exit Sub

 OldDB.Execute "Insert Into Categorys (Category) Values ('" & cboCategoryTypes.Text & "')"
 MsgBox "Saved", vbOKOnly, App.Title
 LoadCbo cboCategoryTypes, "Select Distinct Category from Categorys"

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdDelCategory_Click()
On Error GoTo Err_Routine

 OldDB.Execute "Delete from Categorys where Category = '" & cboCategoryTypes.Text & "'"
 OldDB.Execute "Delete from PriceList where Category = '" & cboCategoryTypes.Text & "'"
 OldDB.Execute "Delete from ProductInfo where Category = '" & cboCategoryTypes.Text & "'"

 MsgBox "Deleted", vbOKOnly, App.Title
 LoadCbo cboCategoryTypes, "Select Distinct Category from Categorys"


Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Private Sub cmdSaveOverheadClass_Click()
On Error GoTo Err_Routine

 If cboOverheadTypes.Text = "" Then Exit Sub
 
 OldDB.Execute "Insert Into Overheads (Overhead) Values ('" & cboOverheadTypes.Text & "')"
 MsgBox "Saved", vbOKOnly, App.Title
 LoadCbo cboOverheadTypes, "Select Distinct Overhead from Overheads"


Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdDelOverheadClass_Click()
On Error GoTo Err_Routine

 OldDB.Execute "Delete from Overheads where Overhead = '" & cboOverheadTypes.Text & "'"
 MsgBox "Deleted", vbOKOnly, App.Title
 LoadCbo cboOverheadTypes, "Select Distinct Overhead from Overheads"


Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdSaveCoInfo_Click()
On Error GoTo Err_Routine

SaveSetting App.Title, "Settings", "txtCoName", txtCoName.Text
SaveSetting App.Title, "Settings", "txtTel", txtTel.Text
SaveSetting App.Title, "Settings", "txtEmail", txtEmail.Text
SaveSetting App.Title, "Settings", "txtFax", txtFax.Text
SaveSetting App.Title, "Settings", "txtAddress", txtAddress.Text

MsgBox "Saved", vbOKOnly, App.Title

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Sub LoadCoInfo()
 If GetSetting(App.Title, "Settings", "txtCoName", "") <> "" Then txtCoName.Text = GetSetting(App.Title, "Settings", "txtCoName", "")
 If GetSetting(App.Title, "Settings", "txtTel", "") <> "" Then txtTel.Text = GetSetting(App.Title, "Settings", "txtTel", "")
 If GetSetting(App.Title, "Settings", "txtEmail", "") <> "" Then txtEmail.Text = GetSetting(App.Title, "Settings", "txtEmail", "")
 If GetSetting(App.Title, "Settings", "txtFax", "") <> "" Then txtFax.Text = GetSetting(App.Title, "Settings", "txtFax", "")
 If GetSetting(App.Title, "Settings", "txtAddress", "") <> "" Then txtAddress.Text = GetSetting(App.Title, "Settings", "txtAddress", "")
End Sub

Private Sub Form_Load()
 On Error GoTo Err_Routine
 
 LoadCbo cboCategoryTypes, "Select Distinct Category from Categorys"
 LoadCbo cboOverheadTypes, "Select Distinct Overhead from Overheads"

 LoadCoInfo
 
Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub
