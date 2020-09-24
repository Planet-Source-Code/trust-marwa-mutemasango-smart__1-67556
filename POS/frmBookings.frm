VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmBookings 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bookings"
   ClientHeight    =   6720
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10890
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBookings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameBooking 
      BackColor       =   &H80000014&
      Caption         =   "Bookings"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.ComboBox cboLodge 
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
         Left            =   1440
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   645
         Left            =   1440
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox txtCusCon 
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
         Left            =   1440
         TabIndex        =   12
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CheckBox chkCheckedOut 
         BackColor       =   &H80000014&
         Caption         =   "Checked Out"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboCusName 
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
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtnotes 
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
         Height          =   735
         Left            =   1440
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3480
         Width           =   4575
      End
      Begin VB.TextBox txtemail 
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
         Left            =   1440
         TabIndex        =   8
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   4440
         Picture         =   "frmBookings.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   305
      End
      Begin VB.CommandButton cmdBooking 
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
         Left            =   4080
         Picture         =   "frmBookings.frx":0941
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "New"
         Top             =   240
         Width           =   350
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
         Left            =   4800
         Picture         =   "frmBookings.frx":0A31
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Delete"
         Top             =   240
         Width           =   350
      End
      Begin VB.TextBox txtPassport 
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
         Left            =   1440
         TabIndex        =   4
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtNationality 
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
         Left            =   1440
         TabIndex        =   3
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtCompany 
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
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin MSACAL.Calendar Arrival 
         Height          =   2775
         Left            =   6240
         TabIndex        =   1
         Top             =   480
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   4895
         _StockProps     =   1
         BackColor       =   -2147483628
         Year            =   2006
         Month           =   7
         Day             =   20
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10052915
         GridLinesColor  =   10052915
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10052915
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvwinfo 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Double Click To Export"
         Top             =   4680
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSACAL.Calendar Departure 
         Height          =   2775
         Left            =   6240
         TabIndex        =   28
         Top             =   3600
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   4895
         _StockProps     =   1
         BackColor       =   -2147483628
         Year            =   2006
         Month           =   7
         Day             =   20
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10052915
         GridLinesColor  =   10052915
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10052915
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Lodge/Table"
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
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         TabIndex        =   26
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         TabIndex        =   25
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number"
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
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Arrival"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Departure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   22
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblClientHistory 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Client History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67D46&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "e-mail"
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
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Passport #"
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
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
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
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
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
         Top             =   960
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   3
   End
End
Attribute VB_Name = "frmBookings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstmain1 As ListItem

Private Sub cboCusName_Click()
Dim RS As New ADODB.Recordset
Dim SQL As String

SQL = "Select * from bookings where CusName = '" & cboCusName.Text & "'"
RS.Open SQL, OldDB
lvwinfo.ListItems.Clear

While Not RS.EOF
 Set lstmain1 = lvwinfo.ListItems.Add(, , RS!ID)
 lstmain1.SubItems(1) = RS!Lodge
 lstmain1.SubItems(2) = RS!CusName
 If Not IsNull(RS!Arrival) Then lstmain1.SubItems(3) = RS!Arrival
 If Not IsNull(RS!Departure) Then lstmain1.SubItems(4) = RS!Departure
 RS.MoveNext
Wend

RS.Close

End Sub



Private Sub cmdBooking_Click()
 cboLodge.Text = ""
 txtAddress.Text = ""
 cboCusName.Text = ""
 txtCompany.Text = ""
 txtCusCon.Text = ""
 txtEmail.Text = ""
 txtNationality.Text = ""
 txtPassport.Text = ""
 chkCheckedOut.Value = vbUnchecked
 txtnotes.Text = ""
 FrameBooking.Caption = "New Booking"
 FormatDates
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err_Routine
Dim i As Integer
Dim SQL As String

 For i = 1 To lvwinfo.ListItems.count
   If lvwinfo.ListItems.Item(i).Checked = True Then
    OldDB.Execute "Delete from bookings where ID = " & lvwinfo.ListItems.Item(i)
   End If
 Next i
 
MsgBox "Record Deleted", vbOKOnly, App.Title

Exit_Routine:
Exit Sub
Err_Routine:
MsgBox Err.Description, , App.Title
Resume Exit_Routine
End Sub


Private Sub cmdSave_Click()
On Error GoTo Err_Routine

Dim SQL As String

If Trim(cboLodge.Text) = "" Then
 MsgBox "Select Lodge", vbOKOnly, App.Title
 Exit Sub
End If

If FrameBooking.Caption = "New Booking" Then
  SQL = "Insert into bookings (Arrival,Departure,email,Address,Lodge,CusName,CusCon,Notes,Passport,Nationality,Company,checkedout) values('" & Arrival.Value & "', '" & Departure.Value & "', '" & txtEmail.Text & "', '" & txtAddress.Text & "', '" & cboLodge.Text & "', '" & cboCusName.Text & "', '" & txtCusCon.Text & "','" & txtnotes.Text & "','" & txtPassport.Text & "','" & txtNationality.Text & "','" & txtCompany.Text & "',0)"
  OldDB.Execute SQL
 Else
  SQL = "Update bookings set Company = '" & txtCompany.Text & "', Address = '" & txtAddress.Text & "', CusName = '" & cboCusName.Text & "', CusCon = '" & txtCusCon.Text & "', Notes = '" & txtnotes.Text & "', email = '" & txtEmail.Text & "', Arrival = '" & Arrival.Value & "', Departure = '" & Departure.Value & "', Nationality = '" & txtNationality.Text & "', Passport = '" & txtPassport.Text & "',"
  
  If chkCheckedOut.Value = vbChecked Then
   SQL = SQL & " checkedout = 1"
  Else
   SQL = SQL & " checkedout = 0"
  End If
  
  SQL = SQL & " where ID = " & lvwinfo.SelectedItem.Text
  OldDB.Execute SQL

End If

MsgBox "Details Saved", vbOKOnly, App.Title

LoadCbo cboCusName, "Select Distinct CusName from bookings"
LoadCbo cboLodge, "Select Distinct Lodge from bookings"

Exit_Routine:
Exit Sub
Err_Routine:
MsgBox Err.Description, , App.Title
Resume Exit_Routine
End Sub

Private Sub Form_Load()
On Error GoTo Err_Routine

Dim clmx As ColumnHeader
Dim RS As New ADODB.Recordset

LoadCbo cboCusName, "Select Distinct CusName from bookings"
LoadCbo cboLodge, "Select Distinct Lodge from bookings"

Set clmx = lvwinfo.ColumnHeaders.Add(, , "#")
Set clmx = lvwinfo.ColumnHeaders.Add(, , "Lodge")
Set clmx = lvwinfo.ColumnHeaders.Add(, , "Client")
Set clmx = lvwinfo.ColumnHeaders.Add(, , "Date of arrival")
Set clmx = lvwinfo.ColumnHeaders.Add(, , "Date of departure")


lvwinfo.ColumnHeaders.Item(1).Width = 600
lvwinfo.ColumnHeaders.Item(2).Width = 2000
lvwinfo.ColumnHeaders.Item(3).Width = 2000
lvwinfo.ColumnHeaders.Item(4).Width = 2000
lvwinfo.ColumnHeaders.Item(5).Width = 2000

FormatDates

Exit_Routine:
Exit Sub
Err_Routine:
MsgBox Err.Description, , App.Title
Resume Exit_Routine
End Sub



Sub FormatDates()
Arrival.Value = Date
Departure.Value = Date
End Sub





Private Sub FrameBooking_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClientHistory.ForeColor = &HA67D46
End Sub

Private Sub lblClientHistory_Click()
frmClientHistory.Show
End Sub

Private Sub lblClientHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClientHistory.ForeColor = &H80FF&
End Sub

Private Sub lvwinfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 ' If the ListView is already sorted by the clicked column, _
 ' just reverse the order. Otherwise, sort the clicked column ascending.
   If lvwinfo.Sorted = True And ColumnHeader.SubItemIndex = lvwinfo.SortKey Then
      If lvwinfo.SortOrder = lvwAscending Then
        lvwinfo.SortOrder = lvwDescending
      Else
        lvwinfo.SortOrder = lvwAscending
      End If
   Else
      lvwinfo.Sorted = True
      lvwinfo.SortKey = ColumnHeader.SubItemIndex
      lvwinfo.SortOrder = lvwAscending
   End If
End Sub

Private Sub lvwinfo_DblClick()
On Error GoTo Err_Routine
Dim MyDataXport As New CDataXport

MyDataXport.lvwXport lvwinfo, dlgCommonDialog

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub lvwinfo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim RS As New ADODB.Recordset
Dim SQL As String

' Set the variable to the SelectedItem.
Set lvwinfo.SelectedItem = Item
    
If Item.Checked Then
 SQL = "Select * from bookings where ID = " & Item.Text
 RS.Open SQL, OldDB
 If Not RS.EOF Then
  cboLodge.Text = RS!Lodge
  txtAddress.Text = RS!Address
  txtEmail.Text = RS!email
  cboCusName.Text = RS!CusName
  txtCusCon.Text = RS!CusCon
  If Not IsNull(RS!Arrival) Then Arrival.Value = RS!Arrival
  If Not IsNull(RS!Arrival) Then Departure.Value = RS!Departure
  
  If RS!checkedout = 1 Then
   chkCheckedOut.Value = vbChecked
  Else
   chkCheckedOut.Value = vbUnchecked
  End If
  txtnotes.Text = RS!Notes
  FrameBooking.Caption = "View Booking"
 End If
 CheckLoadedItems lvwinfo, lvwinfo.SelectedItem.Index
Else
 UncheckLoadedItems lvwinfo, lvwinfo.SelectedItem.Index
End If

End Sub
