VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReports 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reports"
   ClientHeight    =   3405
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6585
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkStock 
      BackColor       =   &H80000014&
      Caption         =   "Stock Purchases"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CheckBox chkSales 
      BackColor       =   &H80000014&
      Caption         =   "Sales Report"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000014&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdRun 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Picture         =   "frmReports.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Edit Product / Service List"
         Top             =   2640
         Width           =   360
      End
      Begin VB.CheckBox chkSalesSummary 
         BackColor       =   &H80000014&
         Caption         =   "Sales Sumary"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000014&
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6135
         Begin VB.CheckBox chkCUS 
            BackColor       =   &H80000014&
            Caption         =   "Cash Up Sheet"
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000014&
            Caption         =   "Report By Month"
            Height          =   255
            Left            =   4440
            TabIndex        =   14
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton optByDate 
            BackColor       =   &H80000014&
            Caption         =   "Report By Date"
            Height          =   255
            Left            =   2760
            TabIndex        =   13
            Top             =   1080
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.ComboBox cbomonthTo 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox cboyearTo 
            Height          =   315
            Left            =   4920
            TabIndex        =   9
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkSalesEmployees 
            BackColor       =   &H80000014&
            Caption         =   "Sales Summary By Lodge/Table"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2655
         End
         Begin VB.ComboBox cboSDate 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Month"
            Height          =   375
            Left            =   2640
            TabIndex        =   12
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            Height          =   255
            Left            =   4440
            TabIndex        =   11
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   3
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Date"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Dim outputstr As String
Dim Y As Integer, i As Integer, maxrows As Integer
Dim X As Integer
Dim StatementsRS As New ADODB.Recordset
Dim SalesRS As New ADODB.Recordset
Dim StoctRS As New ADODB.Recordset
Dim SalesReport1RS As New ADODB.Recordset
Dim SalesReport2RS As New ADODB.Recordset
Dim SalesReport3RS As New ADODB.Recordset
Dim SearchPatern As String
Dim TitleStr As String
Dim ReportDate As String





Private Sub cmdRun_Click()
On Error GoTo Err_Routine
Dim ToMcode As String
Dim ToYcode As String
Screen.MousePointer = vbHourglass

If optByDate.Value = False Then
  If Len(CStr(cbomonthTo.ListIndex + 1)) = 1 Then
   ToMcode = "0" & CStr(cbomonthTo.ListIndex + 1)
  Else
   ToMcode = CStr(cbomonthTo.ListIndex + 1)
  End If

  ToYcode = Right(cboyearTo.Text, 2)
  SearchPatern = "%/" & ToMcode & "/" & ToYcode & "%"
End If

 If optByDate.Value = True Then
   ReportDate = cboSDate.Text
 Else
  ReportDate = cbomonthTo.Text & " " & cboyearTo.Text
 End If
 
TitleStr = chkSales.Caption
If chkSales.Value = vbChecked Then RunSales
TitleStr = chkSalesSummary.Caption
If chkSalesSummary.Value = vbChecked Then RunSales2
TitleStr = chkSalesEmployees.Caption
If chkSalesEmployees.Value = vbChecked Then RunSales3
TitleStr = chkCUS.Caption
If chkCUS.Value = vbChecked Then RunCus
TitleStr = chkStock.Caption
If chkStock.Value = vbChecked Then RunStock

Screen.MousePointer = vbDefault

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Private Sub Form_Load()
On Error Resume Next

StatementsRS.Open "EmployeeInfo", OldDB
SalesRS.Open "Select Distinct SDate from SalesInfo", OldDB

While Not SalesRS.EOF
 cboSDate.AddItem SalesRS!SDate
 SalesRS.MoveNext
Wend

cboSDate.ListIndex = 0

For i = Year(Date) To 2004 Step -1
 cboyearTo.AddItem i
Next i

cbomonthTo.AddItem "January"
cbomonthTo.AddItem "February"
cbomonthTo.AddItem "March"
cbomonthTo.AddItem "April"
cbomonthTo.AddItem "May"
cbomonthTo.AddItem "June"
cbomonthTo.AddItem "July"
cbomonthTo.AddItem "August"
cbomonthTo.AddItem "September"
cbomonthTo.AddItem "October"
cbomonthTo.AddItem "November"
cbomonthTo.AddItem "December"

cboyearTo.Text = Year(Date)
cbomonthTo.ListIndex = Month(Date) - 1

End Sub


Sub RunSales()
 Dim sFile As String
 Dim stTemp As String
 
 If optByDate.Value = True Then
  SalesReport1RS.Open "Select * from SalesInfo where SDate='" & cboSDate.Text & "' Order By ProductDesc", OldDB
 Else
  SalesReport1RS.Open "Select * from SalesInfo where SDate Like '" & SearchPatern & "' Order By ProductDesc", OldDB
 End If
 
    stTemp = "CSV(MS-DOS)|*.csv"
 
         With dlgCommonDialog
            .DialogTitle = "Save Sales Report"
            .CancelError = False
            .DefaultExt = "*.csv"
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = stTemp
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
         End With
         
  Open sFile For Output As #2
   Print #2, ""
   Print #2, TitleStr & " As At " & ReportDate
   
  outputstr = "Product Description,Units,Waiter,Date Of Sale,Charge,VAT"
  Print #2, outputstr

 While Not SalesReport1RS.EOF
  outputstr = SalesReport1RS!ProductDesc & "," & SalesReport1RS!Units & "," & SalesReport1RS!EmployeeCode & "," & SalesReport1RS!SDate & "," & SalesReport1RS!Charge & "," & SalesReport1RS!VAT
  Print #2, outputstr
  SalesReport1RS.MoveNext
 Wend
 Close #2
SalesReport1RS.Close
End Sub

Sub RunSales2()
 Dim sFile As String
 Dim stTemp As String
 Dim SQL As String
 Dim DateStr As String
 
    stTemp = "CSV(MS-DOS)|*.csv"
 
         With dlgCommonDialog
            .DialogTitle = "Save Sales Summary Report"
            .CancelError = False
            .DefaultExt = "*.csv"
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = stTemp
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
         End With
         
  Open sFile For Output As #2
   Print #2, GetCoInfo
   Print #2, TitleStr & " As At " & ReportDate
   
  outputstr = "Product Description,Units,Charge,VAT"
  Print #2, outputstr
  
 StoctRS.Open "ProductInfo", OldDB
 
 While Not StoctRS.EOF
 If optByDate.Value = True Then
  SQL = "SELECT Sum(Units) AS unitsT, sum(Charge) AS chargeT, sum(VAT) as VATT From SalesInfo WHERE SDate ='" & cboSDate.Text & "' and ProductDesc = '" & StoctRS!Menu & "'"
  DateStr = cboSDate.Text
 Else
  SQL = "SELECT Sum(Units) AS unitsT, sum(Charge) AS chargeT, sum(VAT) as VATT From SalesInfo WHERE SDate Like '" & SearchPatern & "' and ProductDesc = '" & StoctRS!Menu & "'"
  DateStr = SearchPatern
 End If
 
  SalesReport2RS.Open SQL, OldDB
'  If SalesReport2RS!ChargeT > 0 Then outputstr = StoctRS!ProductDesc & "," & SalesReport2RS!UnitsT & "," & SalesReport2RS!ChargeT & "," & SalesReport2RS!VATT & "," & DateStr
  If SalesReport2RS!ChargeT > 0 Then
    outputstr = StoctRS!ProductDesc & "," & SalesReport2RS!UnitsT & "," & SalesReport2RS!ChargeT & "," & SalesReport2RS!VATT
    Print #2, outputstr
  End If
  
  StoctRS.MoveNext
  SalesReport2RS.Close
 Wend
 Close #2

End Sub

Sub RunSales3()
 Dim sFile As String
 Dim stTemp As String
 Dim SQL As String
 Dim DateStr As String
 
    stTemp = "CSV(MS-DOS)|*.csv"
 
         With dlgCommonDialog
            .DialogTitle = "Save Sales Summary By Table Report"
            .CancelError = False
            .DefaultExt = "*.csv"
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = stTemp
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
         End With
         
  Open sFile For Output As #2
   Print #2, GetCoInfo
   Print #2, TitleStr & " As At " & ReportDate
   
  outputstr = "Table,Total"
  Print #2, outputstr
  
  While Not StatementsRS.EOF
  If optByDate.Value = True Then
   SQL = "SELECT sum(Charge) AS chargeT From SalesInfo WHERE SDate ='" & cboSDate.Text & "' and TableName= '" & StatementsRS!EName & "'"
   DateStr = cboSDate.Text
  Else
   SQL = "SELECT sum(Charge) AS chargeT From SalesInfo WHERE SDate Like '" & SearchPatern & "' and TableName= '" & StatementsRS!EName & "'"
   DateStr = SearchPatern
  End If
  
  SalesReport3RS.Open SQL, OldDB
 
  outputstr = StatementsRS!EName & "," & SalesReport3RS!ChargeT
  Print #2, outputstr
  StatementsRS.MoveNext
  SalesReport3RS.Close
 Wend
 
 Close #2

End Sub

Sub RunCus()
 Dim sFile As String
 Dim stTemp As String

 
 If optByDate.Value = True Then
  SalesReport1RS.Open "Select * from CUS where SDate='" & cboSDate.Text & "'", OldDB
 Else
  SalesReport1RS.Open "Select * from CUS where SDate Like '" & SearchPatern & "'", OldDB
 End If
 
    stTemp = "CSV(MS-DOS)|*.csv"
 
         With dlgCommonDialog
            .DialogTitle = "Save Cash Up Sheet"
            .CancelError = False
            .DefaultExt = "*.csv"
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = stTemp
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
         End With
         
  Open sFile For Output As #2
   Print #2, GetCoInfo
   Print #2, TitleStr & " As At " & ReportDate
  outputstr = "Table,Customer,Waiter,Date Of Sale,Bill Total,VAT,Meal,Payment Method,Tip"
  Print #2, outputstr

 While Not SalesReport1RS.EOF
  outputstr = SalesReport1RS!TableName & "," & SalesReport1RS!Customer & "," & SalesReport1RS!Waiter & "," & SalesReport1RS!SDate & "," & SalesReport1RS!BillTotal & "," & SalesReport1RS!VAT & "," & SalesReport1RS!Meal & "," & SalesReport1RS!PaymentMethod & "," & SalesReport1RS!Tip
  Print #2, outputstr
  SalesReport1RS.MoveNext
 Wend
 Close #2
 SalesReport1RS.Close
End Sub


Sub RunStock()
 Dim sFile As String
  Dim stTemp As String
  Dim RS As New ADODB.Recordset
  
    stTemp = "CSV(MS-DOS)|*.csv"
 
         With dlgCommonDialog
            .DialogTitle = "Save Stock Report"
            .CancelError = False
            .DefaultExt = "*.csv"
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = stTemp
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
         End With
 
 Open sFile For Output As #2
 Print #2, GetCoInfo
 Print #2, TitleStr
 outputstr = "Product Description,Date,Quantity,Reason,Cost"
  Print #2, outputstr
 
 RS.Open "StockAdjustment", OldDB
 
 While Not RS.EOF
  outputstr = RS!ProductDesc & "," & RS!ADate & "," & RS!Qty & "," & RS!Reason & "," & RS!Cost
  Print #2, outputstr
  RS.MoveNext
 Wend
 Close #2
 
RS.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
StatementsRS.Close
SalesRS.Close
End Sub
