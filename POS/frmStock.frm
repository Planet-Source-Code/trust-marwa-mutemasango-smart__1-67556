VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStock 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Control"
   ClientHeight    =   6840
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10815
   Icon            =   "frmStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAddDel 
      BackColor       =   &H80000014&
      Caption         =   "+/-"
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
      Left            =   10210
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit Product / Service List"
      Top             =   120
      Width           =   360
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
      TabIndex        =   4
      Top             =   120
      Width           =   1550
   End
   Begin VB.ComboBox cboCat 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A67D46&
      Height          =   360
      ItemData        =   "frmStock.frx":08CA
      Left            =   6480
      List            =   "frmStock.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.OptionButton optSection1 
      BackColor       =   &H00293C6B&
      Caption         =   "Store"
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
      TabIndex        =   2
      Top             =   120
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
      TabIndex        =   1
      Top             =   120
      Width           =   1550
   End
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
      TabIndex        =   0
      Top             =   120
      Width           =   1550
   End
   Begin VB.Frame FrameStockSheets 
      BackColor       =   &H80000014&
      Caption         =   "Stock Sheets"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   10575
      Begin VB.CommandButton cmdNewStockSheet 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   9285
         MaskColor       =   &H80000014&
         Picture         =   "frmStock.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "New Stock Sheet"
         Top             =   360
         Width           =   350
      End
      Begin VB.CommandButton cmdExportStockSheet 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   10095
         MaskColor       =   &H80000014&
         Picture         =   "frmStock.frx":09BE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Export"
         Top             =   360
         Width           =   350
      End
      Begin VB.CommandButton cmdSaveStockSheet 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   9720
         MaskColor       =   &H80000014&
         Picture         =   "frmStock.frx":0A35
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save Stock Sheet"
         Top             =   360
         Width           =   305
      End
      Begin VB.ComboBox cboIssueDates 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67D46&
         Height          =   360
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   3735
      End
      Begin VB.CheckBox chkStatus 
         BackColor       =   &H80000014&
         Caption         =   "Stock Levels where updated"
         Enabled         =   0   'False
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
         Left            =   4560
         TabIndex        =   8
         Top             =   405
         Width           =   3495
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridStockSheet 
         Height          =   4455
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   4
         BackColorFixed  =   10911046
         ForeColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         GridColor       =   10911046
         GridColorFixed  =   10911046
         WordWrap        =   -1  'True
         FocusRect       =   2
         FillStyle       =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblHelp 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   10335
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000014&
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   435
         Width           =   615
      End
   End
   Begin VB.Frame FrameStockLevels 
      BackColor       =   &H80000014&
      Caption         =   "Stock Levels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   10575
      Begin VB.CommandButton cmdSaveStockLevels 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   9720
         Picture         =   "frmStock.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Save Stock Sheet"
         Top             =   360
         Width           =   305
      End
      Begin VB.CommandButton cmdExportStockLevels 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   10095
         Picture         =   "frmStock.frx":0B23
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Export"
         Top             =   360
         Width           =   350
      End
      Begin VB.ComboBox cboStockLevels 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67D46&
         Height          =   360
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmdNewStockLevels 
         BackColor       =   &H80000014&
         Height          =   350
         Left            =   9285
         Picture         =   "frmStock.frx":0B9A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "New Stock Sheet"
         Top             =   360
         Width           =   350
      End
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridStockLevels 
         Height          =   4455
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   4
         FixedCols       =   2
         BackColorFixed  =   10911046
         ForeColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         GridColor       =   10911046
         GridColorFixed  =   10911046
         WordWrap        =   -1  'True
         FocusRect       =   2
         FillStyle       =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   435
         Width           =   615
      End
      Begin VB.Label lblHelp2 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   4575
      End
   End
   Begin VB.Label lblStockLevels 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Levels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00996533&
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblStockSheets 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Sheets"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00996533&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cboCat_Click()
On Error GoTo Err_Routine

MSFlexGridStockLevels.Visible = False
ProgressBar3.Visible = True
Screen.MousePointer = vbHourglass
 SearchStockLevels
ProgressBar3.Visible = False
Screen.MousePointer = vbDefault
MSFlexGridStockLevels.Visible = True

If cboCat.Text = "All" Then
  cmdSaveStockSheet.Enabled = False
  cmdNewStockSheet.Enabled = False
  cmdNewStockLevels.Enabled = False
  cmdSaveStockLevels.Enabled = False
Else
  cmdSaveStockSheet.Enabled = True
  cmdNewStockSheet.Enabled = True
  cmdNewStockLevels.Enabled = True
  cmdSaveStockLevels.Enabled = True
End If

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Private Sub cboIssueDates_Click()
On Error GoTo Err_Routine

MSFlexGridStockSheet.Visible = False
ProgressBar2.Visible = True
Screen.MousePointer = vbHourglass
 SearchStockSheet
ProgressBar2.Visible = False
Screen.MousePointer = vbDefault
MSFlexGridStockSheet.Visible = True

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cboStockLevels_Click()
cmdSaveStockLevels.Enabled = False
SearchStockHistory
End Sub

Private Sub cmdAddDel_Click()
frmAddDelProducts.Show vbModal, Me
End Sub

Private Sub cmdExportStockLevels_Click()
On Error GoTo Err_Routine
Dim MyDataXport As New CDataXport

MyDataXport.FlexgridXport MSFlexGridStockLevels, frmMain.dlgCommonDialog

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdExportStockSheet_Click()
On Error GoTo Err_Routine
Dim MyDataXport As New CDataXport

MyDataXport.FlexgridXport MSFlexGridStockSheet, frmMain.dlgCommonDialog

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdNewStockLevels_Click()
cmdSaveStockLevels.Enabled = True
cboIssueDates_Click
End Sub

Private Sub cmdNewStockSheet_Click()
On Error GoTo Err_Routine

MSFlexGridStockSheet.Visible = False
ProgressBar2.Visible = True
Screen.MousePointer = vbHourglass
 
 SetcboTo cboIssueDates, Date
 If cboIssueDates.Text <> Date Then cboIssueDates.AddItem Date
 SetcboTo cboIssueDates, Date
 
 NewStockSheet
ProgressBar2.Visible = False
Screen.MousePointer = vbDefault
MSFlexGridStockSheet.Visible = True

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Sub SaveLevels()
Dim SQL As String
Dim i As Integer

With MSFlexGridStockLevels

For i = 1 To .Rows - 1
 SQL = "Update StockLevels Set Levels = " & .TextMatrix(i, 2) & " where "
 SQL = SQL & "ProductDesc = '" & .TextMatrix(i, 0)
 SQL = SQL & "' and [Section] = '" & Section(Me)
 SQL = SQL & "' and [Category] = '" & cboCat.Text & "'"
 OldDB.Execute SQL
 ProgressBar3.Value = ProgressBar3.Max * (i / .Rows)
Next i

End With
End Sub

Sub RemoveFromStore()
Dim SQL As String
Dim i As Integer

With MSFlexGridStockSheet
lblHelp.Visible = True
lblHelp.Caption = "Updating Stock Levels Please Wait..."
lblHelp.Refresh

If Section(Me) <> "Store" Then
  For i = 1 To .Rows - 1
   SQL = "Update StockLevels Set Levels = Levels - " & .TextMatrix(i, 1) & " where "
   SQL = SQL & "ProductDesc = '" & .TextMatrix(i, 0)
   SQL = SQL & "' and [Section] = 'Store'"
   SQL = SQL & " and [Category] = '" & cboCat.Text & "'"
   OldDB.Execute SQL
   ProgressBar2.Value = ProgressBar2.Max * (i / .Rows)
  Next i
End If

lblHelp.Visible = False

End With
End Sub





Private Sub cmdSaveStockLevels_Click()
On Error GoTo Err_Routine

ProgressBar3.Visible = True
Screen.MousePointer = vbHourglass
 SaveLevels
 SaveStockHistory
ProgressBar3.Visible = False
Screen.MousePointer = vbDefault

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdSaveStockSheet_Click()
On Error GoTo Err_Routine

MSFlexGridStockSheet.Visible = False
ProgressBar2.Visible = True
Screen.MousePointer = vbHourglass
 SaveStockSheet
 
 If chkStatus.Value = vbUnchecked Then
  If MsgBox("Do you wish to update stock levels?", vbYesNo, App.Title) = vbYes Then UpdateStockLevels
 End If
 
 RemoveFromStore
 
ProgressBar2.Visible = False
Screen.MousePointer = vbDefault
MSFlexGridStockSheet.Visible = True

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Private Sub Form_Load()
On Error GoTo Err_Routine

 LoadCbo cboCat, "Select Distinct Category from Categorys"
 cboCat.AddItem "All"
  
 cboCat.ListIndex = 0
 optSection1_Click
  
 MSFlexGridStockSheet.TextMatrix(0, 0) = "Product"
 MSFlexGridStockSheet.TextMatrix(0, 1) = "Stock IN"
 MSFlexGridStockSheet.TextMatrix(0, 2) = "Sales"
 MSFlexGridStockSheet.TextMatrix(0, 3) = "Breakages"
  
 MSFlexGridStockSheet.ColWidth(0) = 2000
 MSFlexGridStockSheet.ColWidth(1) = 2000
 MSFlexGridStockSheet.ColWidth(2) = 2000
 MSFlexGridStockSheet.ColWidth(3) = 2000
 
 MSFlexGridStockLevels.TextMatrix(0, 0) = "Product"
 MSFlexGridStockLevels.TextMatrix(0, 1) = "Expected Levels"
 MSFlexGridStockLevels.TextMatrix(0, 2) = "Actual Levels"
 MSFlexGridStockLevels.TextMatrix(0, 3) = "Notes"
 
 MSFlexGridStockLevels.ColWidth(0) = 2000
 MSFlexGridStockLevels.ColWidth(1) = 2000
 MSFlexGridStockLevels.ColWidth(2) = 2000
 MSFlexGridStockLevels.ColWidth(3) = 4000
 
Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStockSheets.ForeColor = &H996533
lblStockLevels.ForeColor = &H996533
End Sub



Private Sub lblStockLevels_Click()
FrameStockLevels.Visible = True
FrameStockSheets.Visible = False
End Sub

Private Sub lblStockLevels_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStockLevels.ForeColor = &H80FF&
End Sub

Private Sub lblStockSheets_Click()
FrameStockLevels.Visible = False
FrameStockSheets.Visible = True
End Sub

Private Sub lblStockSheets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStockSheets.ForeColor = &H80FF&
End Sub

Private Sub MSFlexGridStockLevels_KeyPress(KeyAscii As Integer)
    With MSFlexGridStockLevels
              If KeyAscii = vbKeyBack Then
                If .Text <> "" Then
                .Text = Left$(.Text, (Len(.Text) - 1))
                End If
              ElseIf KeyAscii = vbKeyReturn Then
                If .Row < .Rows - 1 Then
                 .Row = .Row + 1
                Else
                 .Row = 1
                End If
              Else
                .Text = .Text + Chr$(KeyAscii)
              End If
    End With
End Sub

Private Sub MSFlexGridStockLevels_LeaveCell()
With MSFlexGridStockLevels
   If .Col = 3 Then Exit Sub
   If .TextMatrix(.Row, .Col) = "" Then Exit Sub

   If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
      MsgBox "Numeric Value Expected", vbOKOnly, App.Title
      .TextMatrix(.Row, .Col) = "0"
   End If
End With
End Sub

Private Sub MSFlexGridStockSheet_KeyPress(KeyAscii As Integer)
    With MSFlexGridStockSheet
    If Section(Me) = "Store" And .Col = 2 Then Exit Sub
              If KeyAscii = vbKeyBack Then
                If .Text <> "" Then
                .Text = Left$(.Text, (Len(.Text) - 1))
                End If
              ElseIf KeyAscii = vbKeyReturn Then
                If .Row < .Rows - 1 Then
                 .Row = .Row + 1
                Else
                 .Row = 1
                End If
              Else
                .Text = .Text + Chr$(KeyAscii)
              End If
    End With
End Sub


Sub NewStockSheet()
Dim RS As New ADODB.Recordset
Dim i As Integer

RS.Open "Select Distinct Menu from ProductInfo where Category = '" & cboCat.Text & "'", OldDB

With MSFlexGridStockSheet
  .Rows = RS.RecordCount + 1

  i = 1
 
 While Not RS.EOF
   
   If Not IsNull(RS!Menu) Then
      .TextMatrix(i, 0) = RS!Menu
      .TextMatrix(i, 1) = "0"
      .TextMatrix(i, 2) = GetDailyIssues(RS!Menu, cboCat.Text, Section(Me))
      .TextMatrix(i, 3) = "0"
   End If
   
   RS.MoveNext
   ProgressBar2.Value = ProgressBar2.Max * (i / RS.RecordCount)
   i = i + 1
Wend

End With

End Sub

Sub SearchStockSheet()
Dim RS As New ADODB.Recordset
Dim i As Integer

If cboCat.Text = "All" Then
 RS.Open "Select * from StockSheet where DDate = '" & cboIssueDates.Text & "' and [Section] = '" & Section(Me) & "'", OldDB
Else
 RS.Open "Select * from StockSheet where Category = '" & cboCat.Text & "' and DDate = '" & cboIssueDates.Text & "' and [Section] = '" & Section(Me) & "'", OldDB
End If

 MSFlexGridStockSheet.Rows = RS.RecordCount + 1

If RS.EOF Then
 chkStatus.Value = vbUnchecked
 Exit Sub
ElseIf RS!Status = "Processed" Then
 chkStatus.Value = vbChecked
Else
 chkStatus.Value = vbUnchecked
End If

i = 1
While Not RS.EOF
If Not IsNull(RS!ProductDesc) Then
 MSFlexGridStockSheet.TextMatrix(i, 0) = RS!ProductDesc
 MSFlexGridStockSheet.TextMatrix(i, 1) = RS!StockIN
 MSFlexGridStockSheet.TextMatrix(i, 2) = RS!Issues
 MSFlexGridStockSheet.TextMatrix(i, 3) = RS!Breakages
End If
 RS.MoveNext
 ProgressBar2.Value = ProgressBar2.Max * (i / RS.RecordCount)
 i = i + 1
Wend

End Sub

Sub SaveStockSheet()
Dim SQL As String
Dim i As Integer

SQL = "Delete from StockSheet where Category = '" & cboCat.Text & "' and DDate = '" & cboIssueDates.Text & "' and [Section] = '" & Section(Me) & "'"
OldDB.Execute SQL

With MSFlexGridStockSheet

For i = 1 To MSFlexGridStockSheet.Rows - 1
 SQL = "Insert Into StockSheet ([Section],ProductDesc,StockIN,Category,Issues,Breakages,DDate,MM,YY) Values ('" & Section(Me) & "','" & .TextMatrix(i, 0)
 SQL = SQL & "'," & .TextMatrix(i, 1)
 SQL = SQL & ",'" & cboCat.Text
 SQL = SQL & "'," & .TextMatrix(i, 2)
 SQL = SQL & "," & .TextMatrix(i, 3)
 SQL = SQL & ",'" & cboIssueDates.Text & "'," & Month(cboIssueDates.Text) & "," & Year(cboIssueDates.Text)
 SQL = SQL & ")"
 OldDB.Execute SQL

 ProgressBar2.Value = ProgressBar2.Max * (i / MSFlexGridStockSheet.Rows)
Next i

End With


End Sub

Sub SaveStockHistory()
Dim SQL As String
Dim i As Integer

lblHelp2.Caption = "Saving Stock History"
SQL = "Delete from StockHistory where Category = '" & cboCat.Text & "' and DDate = '" & cboIssueDates.Text & "' and [Section] = '" & Section(Me) & "'"
OldDB.Execute SQL

With MSFlexGridStockLevels

For i = 1 To MSFlexGridStockLevels.Rows - 1
 SQL = "Insert Into StockHistory ([Section],Category,ProductDesc,ExpectedLevels,ActualLevels,StockNotes,DDate) Values ('" & Section(Me) & "','" & cboCat.Text
 SQL = SQL & "','" & .TextMatrix(i, 0)
 SQL = SQL & "'," & .TextMatrix(i, 1)
 SQL = SQL & "," & .TextMatrix(i, 2)
 SQL = SQL & ",'" & .TextMatrix(i, 3)
 SQL = SQL & "','" & Date
 SQL = SQL & "')"
 OldDB.Execute SQL

 ProgressBar3.Value = ProgressBar3.Max * (i / MSFlexGridStockLevels.Rows)
Next i

End With

lblHelp2.Caption = ""

End Sub
Sub SearchStockLevels()
Dim RS As New ADODB.Recordset
Dim i As Integer

If cboCat.Text <> "All" Then
 RS.Open "Select * from StockLevels where [Category] = '" & cboCat.Text & "' and [Section] ='" & Section(Me) & "'", OldDB
Else
 RS.Open "Select * from StockLevels where [Section] ='" & Section(Me) & "'", OldDB
End If

With MSFlexGridStockLevels

 .Rows = RS.RecordCount + 1
 i = 1
 
While Not RS.EOF
  If Not IsNull(RS!ProductDesc) Then
     .TextMatrix(i, 0) = RS!ProductDesc
     .TextMatrix(i, 1) = RS!Levels
     .TextMatrix(i, 2) = RS!Levels
     .TextMatrix(i, 3) = ""
  End If
  RS.MoveNext
  ProgressBar3.Value = ProgressBar3.Max * (i / RS.RecordCount)
  i = i + 1
Wend

End With
End Sub


Sub SearchStockHistory()
Dim RS As New ADODB.Recordset
Dim i As Integer

If cboCat.Text = "All" Then
 RS.Open "Select * from StockHistory where [Section] ='" & Section(Me) & "'", OldDB
Else
 RS.Open "Select * from StockHistory where [Category] = '" & cboCat.Text & "' and [Section] ='" & Section(Me) & "'", OldDB
End If

With MSFlexGridStockLevels

 .Rows = RS.RecordCount + 1

 i = 1
 
While Not RS.EOF
  If Not IsNull(RS!ProductDesc) Then
     .TextMatrix(i, 0) = RS!ProductDesc
     .TextMatrix(i, 1) = RS!ExpectedLevels
     .TextMatrix(i, 2) = RS!ActualLevels
     .TextMatrix(i, 3) = RS!StockNotes
  End If
  RS.MoveNext
  ProgressBar3.Value = ProgressBar3.Max * (i / RS.RecordCount)
  i = i + 1
Wend

End With
End Sub

Sub UpdateStockLevels()
Dim SQL As String
Dim i As Integer
Dim StockOut As Double
Dim StockIN As Double
Dim RS As New ADODB.Recordset

With MSFlexGridStockSheet

For i = 1 To .Rows - 1
    StockOut = CDbl(.TextMatrix(i, 2)) + CDbl(.TextMatrix(i, 3))
    StockIN = CDbl(.TextMatrix(i, 1)) - StockOut

    SQL = "Select ProductDesc from StockLevels where "
    SQL = SQL & "ProductDesc = '" & .TextMatrix(i, 0)
    SQL = SQL & "' and [Section] = '" & Section(Me)
    SQL = SQL & "' and [Category] = '" & cboCat.Text & "'"
    
    RS.Open SQL, OldDB
  
 If RS.EOF Then
    SQL = "Insert Into StockLevels (Levels,ProductDesc,[Section],[Category]) Values ("
    SQL = SQL & StockIN & ",'" & .TextMatrix(i, 0) & "','" & Section(Me) & "','" & cboCat.Text & "')"
 Else
    SQL = "Update StockLevels Set Levels = Levels + " & StockIN & " where "
    SQL = SQL & "ProductDesc = '" & .TextMatrix(i, 0)
    SQL = SQL & "' and [Section] = '" & Section(Me)
    SQL = SQL & "' and [Category] = '" & cboCat.Text & "'"
 End If
 
 OldDB.Execute SQL
 RS.Close
 ProgressBar2.Value = ProgressBar2.Max * (i / .Rows)
Next i

OldDB.Execute "Update StockSheet Set Status = 'Processed' where Category = '" & cboCat.Text & "' and DDate = '" & cboIssueDates.Text & "' and [Section] = '" & Section(Me) & "'"
chkStatus.Value = vbChecked
End With
End Sub

Private Sub MSFlexGridStockSheet_LeaveCell()
With MSFlexGridStockSheet
   If .TextMatrix(.Row, .Col) = "" Then Exit Sub

   If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
      MsgBox "Numeric Value Expected", vbOKOnly, App.Title
      .TextMatrix(.Row, .Col) = "0"
   End If
End With
End Sub

Private Sub optSection1_Click()
MSFlexGridStockSheet.BackColor = optSection1.BackColor
MSFlexGridStockSheet.ForeColor = optSection1.ForeColor

MSFlexGridStockLevels.BackColor = optSection1.BackColor
MSFlexGridStockLevels.ForeColor = optSection1.ForeColor

cboIssueDates_Click

LoadCbo cboStockLevels, "Select Distinct DDate from StockHistory where [Section] ='" & Section(Me) & "'"
LoadCbo cboIssueDates, "Select Distinct DDate from StockSheet where [Section] ='" & Section(Me) & "'"

End Sub

Private Sub optSection2_Click()
MSFlexGridStockSheet.BackColor = optSection2.BackColor
MSFlexGridStockSheet.ForeColor = optSection2.ForeColor

MSFlexGridStockLevels.BackColor = optSection2.BackColor
MSFlexGridStockLevels.ForeColor = optSection2.ForeColor
cboIssueDates_Click

LoadCbo cboStockLevels, "Select Distinct DDate from StockHistory where [Section] ='" & Section(Me) & "'"
LoadCbo cboIssueDates, "Select Distinct DDate from StockSheet where [Section] ='" & Section(Me) & "'"

End Sub

Private Sub optSection3_Click()
MSFlexGridStockSheet.BackColor = optSection3.BackColor
MSFlexGridStockSheet.ForeColor = optSection3.ForeColor

MSFlexGridStockLevels.BackColor = optSection3.BackColor
MSFlexGridStockLevels.ForeColor = optSection3.ForeColor
cboIssueDates_Click

LoadCbo cboStockLevels, "Select Distinct DDate from StockHistory where [Section] ='" & Section(Me) & "'"
LoadCbo cboIssueDates, "Select Distinct DDate from StockSheet where [Section] ='" & Section(Me) & "'"

End Sub

Private Sub optSection4_Click()
MSFlexGridStockSheet.BackColor = optSection4.BackColor
MSFlexGridStockSheet.ForeColor = optSection4.ForeColor

MSFlexGridStockLevels.BackColor = optSection4.BackColor
MSFlexGridStockLevels.ForeColor = optSection4.ForeColor
cboIssueDates_Click

LoadCbo cboStockLevels, "Select Distinct DDate from StockHistory where [Section] ='" & Section(Me) & "'"
LoadCbo cboIssueDates, "Select Distinct DDate from StockSheet where [Section] ='" & Section(Me) & "'"

End Sub
