VERSION 5.00
Begin VB.Form frmAddDelProducts 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Products / Services"
   ClientHeight    =   2010
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3945
   Icon            =   "frmAddDelProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSaveProductServices 
      BackColor       =   &H80000014&
      Height          =   350
      Left            =   120
      Picture         =   "frmAddDelProducts.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   350
   End
   Begin VB.CommandButton cmdDelProductServices 
      BackColor       =   &H80000014&
      Height          =   350
      Left            =   600
      Picture         =   "frmAddDelProducts.frx":0941
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   350
   End
   Begin VB.ComboBox cboProductService 
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
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
   End
   Begin VB.ComboBox cboCategoryProdServ 
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
      ItemData        =   "frmAddDelProducts.frx":0AA7
      Left            =   120
      List            =   "frmAddDelProducts.frx":0AA9
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddDelProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdDelProductServices_Click()
On Error GoTo Err_Routine

 OldDB.Execute "Delete from ProductInfo where Menu = '" & cboProductService.Text & "' and Category = '" & cboCategoryProdServ.Text & "'"
 OldDB.Execute "Delete from PriceList where ProductDesc = '" & cboProductService.Text & "' and Category = '" & cboCategoryProdServ.Text & "'"

 MsgBox "Deleted", vbOKOnly, App.Title
 cboCategoryProdServ_Click

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Private Sub cmdSaveProductServices_Click()
On Error GoTo Err_Routine
Dim RS As New ADODB.Recordset
Dim SQL As String

If cboCategoryProdServ.Text = "" Then
 MsgBox "Select Category", vbOKOnly, App.Title
 Exit Sub
End If

RS.Open "Select * from ProductInfo where Menu = '" & cboProductService.Text & "' and Category = '" & cboCategoryProdServ.Text & "'", OldDB

If RS.EOF Then SQL = "Insert Into ProductInfo (Menu,Category,FreePrice,MarkUp,NumberPple,DOC) Values ('" & cboProductService.Text & "','" & cboCategoryProdServ.Text & "',0,0,0,'" & Date & "')"

OldDB.Execute SQL

RS.Close

MsgBox "Saved", vbOKOnly, App.Title

cboCategoryProdServ_Click

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cboCategoryProdServ_Click()
LoadCbo cboProductService, "Select * from ProductInfo where Category= '" & cboCategoryProdServ.Text & "'"
End Sub


Private Sub Form_Load()
On Error GoTo Err_Routine

LoadCbo cboCategoryProdServ, "Select Distinct Category from Categorys"

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub
