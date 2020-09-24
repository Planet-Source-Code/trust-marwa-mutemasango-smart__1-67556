VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClientHistory 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client History"
   ClientHeight    =   3345
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7200
   Icon            =   "frmClientHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboCusName 
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
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   6975
   End
   Begin MSComctlLib.ListView lvwinfo 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click To Export"
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4895
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   3
      FontStrikeThru  =   -1  'True
      FontUnderLine   =   -1  'True
   End
End
Attribute VB_Name = "frmClientHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub cboCusName_Click()
 LoadReports lvwinfo, "Select CusName as 'Client', Company as 'Company', Lodge as 'Lodge/table', Address as 'Address',CusCon as 'Contact Number',email as 'E-mail',Notes as 'Notes',Arrival as 'Arrival',Departure as 'Departure',Passport as 'Passport',Nationality as 'Nationality' from bookings where CusName ='" & cboCusName.Text & "'"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Routine

 LoadCbo cboCusName, "Select Distinct CusName from bookings"
 LoadReports lvwinfo, "Select CusName as 'Client', Company as 'Company', Lodge as 'Lodge/table', Address as 'Address',CusCon as 'Contact Number',email as 'E-mail',Notes as 'Notes',Arrival as 'Arrival',Departure as 'Departure',Passport as 'Passport',Nationality as 'Nationality' from bookings"

Exit_Routine:
Exit Sub
Err_Routine:
MsgBox Err.Description, , App.Title
Resume Exit_Routine
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
' Set the variable to the SelectedItem.
Set lvwinfo.SelectedItem = Item
    
If Item.Checked Then
 CheckLoadedItems lvwinfo, lvwinfo.SelectedItem.Index
Else
 UncheckLoadedItems lvwinfo, lvwinfo.SelectedItem.Index
End If

End Sub

