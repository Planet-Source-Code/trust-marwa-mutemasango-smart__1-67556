VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCharges 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Charges"
   ClientHeight    =   7725
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6945
   Icon            =   "frmCharges.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H80000014&
      Caption         =   "Payment Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   6735
      Begin VB.CommandButton cmdPayment 
         Height          =   350
         Left            =   1800
         Picture         =   "frmCharges.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Payment"
         Top             =   1320
         Width           =   350
      End
      Begin VB.TextBox txtVAT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000014&
         Caption         =   "VISA / Mastercard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optCheque 
         BackColor       =   &H80000014&
         Caption         =   "Cheque Payment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optCash 
         BackColor       =   &H80000014&
         Caption         =   "Cash Payment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox txtTotalCharge 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtPayment 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Text            =   "0"
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "VAT | Total Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Payed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Caption         =   "Charges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.OptionButton optSection3 
         BackColor       =   &H0091AA98&
         Caption         =   "Kitchen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   360
         Left            =   3360
         TabIndex        =   28
         Top             =   240
         Width           =   1550
      End
      Begin VB.OptionButton optSection1 
         BackColor       =   &H00008000&
         Caption         =   "Services"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   360
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Value           =   -1  'True
         Width           =   1550
      End
      Begin VB.OptionButton optSection2 
         BackColor       =   &H00800000&
         Caption         =   "Bar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   360
         Left            =   1800
         TabIndex        =   26
         Top             =   240
         Width           =   1550
      End
      Begin VB.OptionButton optSection4 
         BackColor       =   &H000040C0&
         Caption         =   "Restaurant"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   360
         Left            =   4920
         TabIndex        =   25
         Top             =   240
         Width           =   1550
      End
      Begin VB.CommandButton cmdSave 
         Height          =   350
         Left            =   1800
         Picture         =   "frmCharges.frx":0941
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Charge"
         Top             =   2160
         Width           =   350
      End
      Begin VB.ComboBox cboCategory 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cboUnits 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cboBookingCode 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1800
         Width           =   4695
      End
      Begin VB.ComboBox cboProdDesc 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Category"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Description"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000014&
      Caption         =   "Bill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   6735
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   600
         Picture         =   "frmCharges.frx":09B8
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Export"
         Top             =   240
         Width           =   350
      End
      Begin VB.CommandButton cmdDeleteTransaction 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   120
         Picture         =   "frmCharges.frx":0A2F
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Delete Transaction"
         Top             =   240
         Width           =   350
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         View            =   3
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
      Begin VB.Label lblChange 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00996533&
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   240
         Width           =   4815
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   3
      FontStrikeThru  =   -1  'True
      FontUnderLine   =   -1  'True
   End
End
Attribute VB_Name = "frmCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim lstmain1 As ListItem

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_OOM = 8
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_SHARE = 26
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_DLLNOTFOUND = 32

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    


Private Sub cboBookingCode_Click()
ShowBill
End Sub

Private Sub cboCategory_Click()
 LoadCbo cboProdDesc, "Select Distinct Menu From ProductInfo Where category = '" & cboCategory.List(cboCategory.ListIndex) & "'"
 cboProdDesc.Enabled = True
End Sub

Private Sub cboProdDesc_Click()
Dim MyCosting As New CCosting
txtPrice.Text = MyCosting.GetSellingPrice(cboProdDesc.Text)
cboUnits.ListIndex = 0
End Sub

Private Sub ShowBill()
On Error GoTo Err_Routine
Dim RS As New ADODB.Recordset
Dim TotalSale As Currency
Dim VarCharge As Currency

If Trim(cboBookingCode.Text) = "" Then
 MsgBox "Select Booking Code", vbOKOnly, App.Title
 Exit Sub
End If

RS.Open "select * from ChargesInfo where BookingCode = " & GetBookingCode(cboBookingCode.Text), OldDB
'open file that contains database path

TotalSale = 0
ListView1.ListItems.Clear

While Not RS.EOF
 Set lstmain1 = ListView1.ListItems.Add(, , RS!TransCode)
 If Not IsNull(RS!ProductDesc) Then lstmain1.SubItems(1) = RS!ProductDesc
 If Not IsNull(RS!Units) Then lstmain1.SubItems(2) = RS!Units
 If Not IsNull(RS!Charge) Then lstmain1.SubItems(3) = RS!Charge
 If Not IsNull(RS!STime) Then lstmain1.SubItems(4) = RS!STime
 If Not IsNull(RS!SDate) Then lstmain1.SubItems(5) = RS!SDate
 
 VarCharge = RS!Charge * RS!Units
 TotalSale = TotalSale + VarCharge
 RS.MoveNext
Wend

checkForPayment
txtTotalCharge.Text = TotalSale * (1 + (txtVAT.Text / 100))

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Sub checkForPayment()
Dim RS As New ADODB.Recordset

RS.Open "select Sum(BillTotal) as num from Payments where BookingCode = '" & cboBookingCode.Text & "'", OldDB

If RS.EOF Then
  cmdPayment.Enabled = True
Else
  cmdPayment.Enabled = False
  If Not IsNull(RS!num) Then txtPayment.Text = RS!num
End If

End Sub



Private Sub cmdDeleteTransaction_Click()
On Error Resume Next
Dim i As Integer
Dim SQL As String

 For i = 1 To ListView1.ListItems.count
   If ListView1.ListItems.Item(i).Checked = True Then
     SQL = "Delete from ChargesInfo where TransCode = " & ListView1.ListItems.Item(i)
     OldDB.Execute SQL
         
     SQL = "Delete from DailyIssues where DateOfIssue = '" & ListView1.ListItems.Item(i).SubItems(5) & "' and TimeOfIssue = '" & ListView1.ListItems.Item(i).SubItems(4) & "'"
     ListView1.ListItems.Remove i
   End If
 Next i
 MsgBox "Transaction Deleted", vbOKOnly, App.Title
End Sub



Private Sub cmdExport_Click()
On Error GoTo Err_Routine
Dim MyDataXport As New CDataXport

MyDataXport.lvwXport ListView1, dlgCommonDialog

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdPayment_Click()
On Error GoTo Err_Routine

Dim SQL As String, PaymentMethodStr As String


If Not IsNumeric(txtPayment.Text) Then
 MsgBox "Enter amount payed by customer", vbOKOnly, App.Title
 Exit Sub
End If

If optCash.Value = True Then
 PaymentMethodStr = "Cash"
ElseIf optCheque.Value = True Then
 PaymentMethodStr = "Cheque"
Else
PaymentMethodStr = "VISA/Mastercard"
End If


SQL = "Insert into Payments (BookingCode,SDate,BillTotal,PaymentMethod) Values('" & cboBookingCode.Text & "', '" & Format(Date, "dd/mm/yy") & "', " & txtTotalCharge.Text & ", '" & PaymentMethodStr & "')"
OldDB.Execute SQL

SaveSetting App.Title, "Settings", "txtVAT", txtVAT.Text

cmdPayment.Enabled = False

MsgBox "Payment received", vbOKOnly, App.Title

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub





Private Sub cmdSave_Click()
On Error GoTo Err_Routine
Dim SQL As String
Dim TimeOfSale As String

If cboBookingCode.Text = "" Then
  MsgBox "Please select Client", vbOKOnly, App.Title
  Exit Sub
End If

If cboProdDesc.Text = "" Then
  MsgBox "Please select service description", vbOKOnly, App.Title
  Exit Sub
End If

If cboCategory.Text = "" Then
  MsgBox "Please select service category", vbOKOnly, App.Title
  Exit Sub
End If

Screen.MousePointer = vbHourglass
    
TimeOfSale = Time
SQL = "Insert Into ChargesInfo (ProductDesc,Units,BookingCode,SDate,Charge,Stime) values ('" & cboProdDesc.Text & "', " & cboUnits.Text & ", '" & GetBookingCode(cboBookingCode.Text) & "', '" & Format(Date, "dd/mm/yy") & "', " & txtPrice.Text * CInt(cboUnits.Text) & ",'" & TimeOfSale & "')"
OldDB.Execute SQL
  

SQL = "Insert Into DailyIssues (ProductDesc,[Category],[Section],TimeOfIssue,DateOfIssue,UnitsIssued) values ('" & cboProdDesc.Text & "', '" & cboCategory.Text & "', '" & Section(Me) & "', '" & TimeOfSale & "','" & Format(Date, "dd/mm/yy") & "', " & cboUnits.Text & ")"
OldDB.Execute SQL

Screen.MousePointer = vbDefault
MsgBox "Transaction Saved", vbOKOnly, App.Title
cboUnits.ListIndex = 0
ShowBill

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub



Private Sub Form_Load()
On Error Resume Next

Dim i As Integer
Dim clmx As ColumnHeader

LoadCbo cboBookingCode, "select Distinct CusName from bookings where checkedout = 0"

cboProdDesc.Enabled = False
Set clmx = ListView1.ColumnHeaders.Add(, , "Transaction Code")
Set clmx = ListView1.ColumnHeaders.Add(, , "Service Description")
Set clmx = ListView1.ColumnHeaders.Add(, , "Units")
Set clmx = ListView1.ColumnHeaders.Add(, , "Price")
Set clmx = ListView1.ColumnHeaders.Add(, , "Time")
Set clmx = ListView1.ColumnHeaders.Add(, , "Date")

ListView1.ColumnHeaders.Item(1).Width = 1550
ListView1.ColumnHeaders.Item(2).Width = 2000
ListView1.ColumnHeaders.Item(3).Width = 550
ListView1.ColumnHeaders.Item(4).Width = 2000
ListView1.ColumnHeaders.Item(5).Width = 1000
ListView1.ColumnHeaders.Item(6).Width = 1000

LoadCbo cboCategory, "Select Distinct Category from ProductInfo"

If GetSetting(App.Title, "Settings", "txtVAT", "") <> "" Then txtVAT.Text = GetSetting(App.Title, "Settings", "txtVAT", "")

For i = 1 To 10
 cboUnits.AddItem i
Next i

cboUnits.ListIndex = 0

cmdPayment.Enabled = False
If SuperUser = False Then cmdDeleteTransaction.Visible = False

End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
' Set the variable to the SelectedItem.
Set ListView1.SelectedItem = Item
    
If Item.Checked Then
 CheckLoadedItems ListView1, ListView1.SelectedItem.Index
Else
 UncheckLoadedItems ListView1, ListView1.SelectedItem.Index
End If
End Sub


Private Sub txtPayment_Change()
If IsNumeric(txtPayment.Text) Then lblChange.Caption = "Change = " & CDbl(txtPayment.Text) - CDbl(txtTotalCharge.Text)

End Sub

Private Sub txtPayment_LostFocus()
If IsNumeric(txtPayment.Text) Then lblChange.Caption = "Change = " & CDbl(txtPayment.Text) - CDbl(txtTotalCharge.Text)
End Sub

Private Sub txtPayment_Validate(Cancel As Boolean)
 If Not IsNumeric(txtPayment.Text) Then
  MsgBox "A Numeric Value Is Expected", vbOKOnly, App.Title
  txtPayment.Text = "0"
 End If
End Sub


Private Sub txtTotalCharge_Validate(Cancel As Boolean)
 If Not IsNumeric(txtTotalCharge.Text) Then
  MsgBox "A Numeric Value Is Expected", vbOKOnly, App.Title
  txtTotalCharge.Text = "0"
 End If
End Sub

Private Sub txtVAT_LostFocus()
ShowBill
End Sub

Private Sub txtVAT_Validate(Cancel As Boolean)
 If Not IsNumeric(txtVAT.Text) Then
  MsgBox "A Numeric Value Is Expected", vbOKOnly, App.Title
  txtVAT.Text = "0"
 End If
End Sub
