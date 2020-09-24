VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmproducts 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costing"
   ClientHeight    =   6390
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmproducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCosts 
      BackColor       =   &H80000014&
      Caption         =   "Costs"
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   10575
      Begin VB.Frame Frame6 
         BackColor       =   &H80000014&
         Caption         =   "Overhead Apportionment"
         Height          =   3015
         Left            =   120
         TabIndex        =   58
         Top             =   2640
         Width           =   10335
         Begin VB.CommandButton cmdSaveApportionment 
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
            Left            =   9360
            Picture         =   "frmproducts.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   240
            Width           =   350
         End
         Begin VB.CommandButton cmdExportApportionment 
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
            Left            =   9840
            Picture         =   "frmproducts.frx":0941
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Export"
            Top             =   240
            Width           =   350
         End
         Begin VB.CommandButton cmdNewApportionment 
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
            Left            =   8880
            Picture         =   "frmproducts.frx":09B8
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "New Apportionment"
            Top             =   240
            Width           =   350
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2175
            Left            =   120
            TabIndex        =   62
            Top             =   720
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   10000
            Cols            =   5
            FixedCols       =   3
            BackColorFixed  =   10911046
            ForeColorFixed  =   -2147483628
            BackColorBkg    =   -2147483628
            GridColor       =   10911046
            GridColorFixed  =   10911046
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblTotalOverLabel 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Total Overheads"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblTotalOver 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00A67D46&
            Height          =   255
            Left            =   2520
            TabIndex        =   63
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000014&
         Caption         =   "Ingredients"
         Height          =   2175
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton cmdDeleteIngred 
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
            Left            =   2400
            Picture         =   "frmproducts.frx":0AA8
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   1680
            Width           =   350
         End
         Begin VB.TextBox txtIngredAmount 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   52
            Top             =   960
            Width           =   2895
         End
         Begin VB.ComboBox cboIngredUnits 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtPrice 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   50
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdSaveIngred 
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
            Left            =   1920
            Picture         =   "frmproducts.frx":0C0E
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1680
            Width           =   350
         End
         Begin VB.ComboBox txtIngredName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label24 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label25 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Units"
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
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label26 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Of Measurement"
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
            Left            =   120
            TabIndex        =   55
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label27 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Cost Price"
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
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000014&
         Caption         =   "Overheads"
         Height          =   2175
         Left            =   5400
         TabIndex        =   38
         Top             =   360
         Width           =   5055
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
            Left            =   1920
            Picture         =   "frmproducts.frx":0C85
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1680
            Width           =   350
         End
         Begin VB.ComboBox cboClass 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox txtAmount 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   41
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdDelOverheads 
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
            Left            =   2400
            Picture         =   "frmproducts.frx":0CFC
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1680
            Width           =   350
         End
         Begin VB.ComboBox txtName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Overhead Class"
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
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Overhead Cost"
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
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Name Of Overhead"
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
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Frame FrameMenus 
      BackColor       =   &H80000014&
      Caption         =   "Menus"
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   10575
      Begin VB.Frame Frame7 
         BackColor       =   &H80000014&
         Caption         =   "Select Ingredients Used"
         Height          =   2655
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   5175
         Begin VB.ComboBox cboIngredient 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdEnter 
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
            Left            =   2160
            Picture         =   "frmproducts.frx":0E62
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1560
            Width           =   350
         End
         Begin VB.TextBox txtAmountUsed 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   32
            Text            =   "0"
            Top             =   720
            Width           =   2895
         End
         Begin VB.ComboBox cboUnits 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000014&
            Caption         =   "Ingredient"
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
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Used"
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
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label18 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Units "
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
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000014&
         Caption         =   "Menu Information"
         Height          =   5295
         Left            =   5400
         TabIndex        =   21
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton cmdShow 
            BackColor       =   &H80000014&
            Caption         =   "Show Menu Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdRemove 
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
            Left            =   120
            Picture         =   "frmproducts.frx":0F66
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4800
            Width           =   350
         End
         Begin VB.CommandButton cmdExportIngridients 
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
            Left            =   600
            Picture         =   "frmproducts.frx":1069
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Export"
            Top             =   4800
            Width           =   350
         End
         Begin MSComctlLib.ProgressBar ProgressBar4 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   2040
            Visible         =   0   'False
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            Max             =   3
            Scrolling       =   1
         End
         Begin MSComctlLib.ListView lvwIncludedIngred 
            Height          =   2175
            Left            =   120
            TabIndex        =   26
            Top             =   2520
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   3836
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
               Text            =   "Ingredient"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Amount Used"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Units"
               Object.Width           =   2364
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Unit Cost"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblCostPerIndividual 
            BackStyle       =   0  'Transparent
            Caption         =   "Cost Per Individual"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
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
            TabIndex        =   29
            Top             =   1080
            Width           =   5055
         End
         Begin VB.Label lblApportionedOver 
            BackStyle       =   0  'Transparent
            Caption         =   "Aportioned Overhead"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
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
            TabIndex        =   28
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label lblTotalIngred 
            BackStyle       =   0  'Transparent
            Caption         =   "Cost of Production"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
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
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000014&
         Caption         =   "Menu"
         Height          =   2535
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txtNumberPple 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Text            =   "1"
            Top             =   1080
            Width           =   2895
         End
         Begin VB.ComboBox cboMenuCategory 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmproducts.frx":10E0
            Left            =   2160
            List            =   "frmproducts.frx":10E2
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   360
            Width           =   2895
         End
         Begin VB.ComboBox cboMenuName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton cmdDeletePrice 
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
            Left            =   2640
            Picture         =   "frmproducts.frx":10E4
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1560
            Width           =   350
         End
         Begin VB.CommandButton cmdSavePrice 
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
            Left            =   2160
            Picture         =   "frmproducts.frx":124A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1560
            Width           =   350
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Number Of People Served"
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
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label29 
            BackColor       =   &H80000014&
            Caption         =   "Category"
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
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label23 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin VB.Frame FrameLists 
      BackColor       =   &H80000014&
      Caption         =   "Lists"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10575
      Begin VB.OptionButton optSection2 
         BackColor       =   &H000040C0&
         Caption         =   "Ingrediants"
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
         Left            =   5040
         TabIndex        =   10
         Top             =   390
         Width           =   1550
      End
      Begin VB.OptionButton optSection1 
         BackColor       =   &H00008000&
         Caption         =   "Prices"
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
         Left            =   3130
         TabIndex        =   9
         Top             =   390
         Width           =   1550
      End
      Begin VB.OptionButton optSection3 
         BackColor       =   &H0091AA98&
         Caption         =   "Overheads"
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
         Left            =   6600
         TabIndex        =   8
         Top             =   390
         Value           =   -1  'True
         Width           =   1550
      End
      Begin VB.CommandButton cmdExport 
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
         Left            =   9960
         Picture         =   "frmproducts.frx":12C1
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Export"
         Top             =   360
         Width           =   350
      End
      Begin VB.CommandButton cmdUpdate 
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
         Left            =   9585
         Picture         =   "frmproducts.frx":1338
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Update Prices"
         Top             =   360
         Width           =   305
      End
      Begin VB.ComboBox cboPriceList 
         Enabled         =   0   'False
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
         ItemData        =   "frmproducts.frx":13AF
         Left            =   240
         List            =   "frmproducts.frx":13B1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   390
         Width           =   2895
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridPriceList 
         Height          =   4455
         Left            =   120
         TabIndex        =   7
         Top             =   1230
         Visible         =   0   'False
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   5
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   10215
      End
   End
   Begin VB.Label lblLists 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Lists"
      ForeColor       =   &H00996533&
      Height          =   255
      Left            =   2040
      TabIndex        =   67
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblMenus 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Menus"
      ForeColor       =   &H00996533&
      Height          =   255
      Left            =   1080
      TabIndex        =   66
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblCosts 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Costs"
      ForeColor       =   &H00996533&
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmproducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboMenuCategory_Click()
LoadCbo cboMenuName, "Select * from ProductInfo where Category= '" & cboMenuCategory.Text & "'"
End Sub

Private Sub cboMenuName_Click()
Dim RS As New ADODB.Recordset
Dim SQL As String

RS.Open "Select * from ProductInfo where Menu = '" & cboMenuName.Text & "'", OldDB

If Not RS.EOF Then
 SetcboTo cboMenuCategory, RS!Category
 If Not IsNull(RS!NumberPple) Then txtNumberPple.Text = RS!NumberPple
  
End If

RS.Close

LoadMenuIngredients
End Sub

Private Sub cboPriceList_Click()
If optSection1.Value = True Then ShowList
End Sub

Private Sub cmdDeleteIngred_Click()
On Error GoTo Err_Routine

 OldDB.Execute "Delete from Ingredients where Name = '" & txtIngredName.Text & "'"
 OldDB.Execute "Delete from Apportionment where Name = '" & txtIngredName.Text & "'"

 MsgBox "Deleted", vbOKOnly, App.Title
 
 LoadCbo txtIngredName, "Select Distinct Name from Ingredients order by Name ASC"
 LoadCbo cboIngredient, "Select Distinct Name from Ingredients order by Name ASC"

 txtPrice.Text = ""
 txtIngredAmount.Text = ""
 ConfigMSFlexGrid1

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdDeletePrice_Click()
On Error GoTo Err_Routine

 OldDB.Execute "Delete from ProductInfo where Menu = '" & cboMenuName.Text & "' and Category = '" & cboMenuCategory & "'"
 MsgBox "Deleted", vbOKOnly, App.Title
 cboMenuCategory_Click

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub



Private Sub cmdDelOverheads_Click()
On Error GoTo Err_Routine

 OldDB.Execute "Delete from OverheadInfo where Name = '" & txtName.Text & "'"
 MsgBox "Deleted", vbOKOnly, App.Title
 LoadCbo txtName, "Select Distinct Name from OverheadInfo order by Name ASC"

 txtAmount.Text = ""
 lblTotalOver_Click
 
Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Private Sub cmdEnter_Click()
On Error GoTo Err_Routine

Dim RS As New ADODB.Recordset
Dim SQL As String

If ValidateForm4 = False Then Exit Sub

RS.Open "Select * from MenuCost Where Menu = '" & cboMenuName.Text & "' and Ingredient = '" & cboIngredient.Text & "'", OldDB

If RS.EOF Then
 SQL = "Insert Into MenuCost (Menu,Ingredient,AmountUsed,Units) Values ('" & cboMenuName.Text & "','" & cboIngredient.Text & "'," & txtAmountUsed.Text & ",'" & cboUnits.Text & "')"
Else
 SQL = "Update MenuCost Set AmountUsed = " & txtAmountUsed.Text & ",[Units] = '" & cboUnits.Text & "' Where [Menu] = '" & cboMenuName.Text & "' and Ingredient = '" & cboIngredient.Text & "'"
End If
 OldDB.Execute SQL

MsgBox "Saved", vbOKOnly, App.Title
LoadMenuIngredients

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Function ValidateForm4() As Boolean

If txtAmountUsed.Text = "" Then
MsgBox "Enter The Amount Used", vbOKOnly, App.Title
ValidateForm4 = False
Exit Function
End If

If Not IsNumeric(txtAmountUsed.Text) Then
MsgBox "Amount Used Must Be A Numeric Value", vbOKOnly, App.Title
ValidateForm4 = False
Exit Function
End If

If cboIngredient.Text = "" Then
MsgBox "Select Ingredient", vbOKOnly, App.Title
ValidateForm4 = False
Exit Function
End If

If cboUnits.Text = "" Then
MsgBox "Select Units Of Measurement", vbOKOnly, App.Title
ValidateForm4 = False
Exit Function
End If


ValidateForm4 = True
End Function

Private Sub cmdExport_Click()
On Error GoTo Err_Routine

Dim MyDataXport As New CDataXport

MyDataXport.FlexgridXport MSFlexGridPriceList, CommonDialog1

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Private Sub cmdExportApportionment_Click()
On Error GoTo Err_Routine
Dim MyDataXport As New CDataXport

MyDataXport.FlexgridXport MSFlexGrid1, CommonDialog1

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdExportIngridients_Click()
On Error GoTo Err_Routine
Dim MyDataXport As New CDataXport

MyDataXport.lvwXport lvwIncludedIngred, CommonDialog1

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Private Sub cmdNewApportionment_Click()
On Error GoTo Err_Routine
Dim RS As New ADODB.Recordset
Dim count As Integer

RS.Open "Select Menu,Category from ProductInfo order by Category ASC", OldDB

 MSFlexGrid1.Rows = RS.RecordCount + 1
 MSFlexGrid1.ColWidth(0) = 2000
 MSFlexGrid1.ColWidth(1) = 2200
 MSFlexGrid1.ColWidth(2) = 1500
 
 MSFlexGrid1.TextMatrix(0, 0) = "Category"
 MSFlexGrid1.TextMatrix(0, 1) = "Product"
 MSFlexGrid1.TextMatrix(0, 2) = "Apportionment"
 MSFlexGrid1.TextMatrix(0, 3) = "Sales"
 MSFlexGrid1.TextMatrix(0, 4) = "Weight"

While Not RS.EOF
 MSFlexGrid1.TextMatrix(count + 1, 0) = RS!Category
 MSFlexGrid1.TextMatrix(count + 1, 1) = RS!Menu
 MSFlexGrid1.TextMatrix(count + 1, 2) = "0"
 MSFlexGrid1.TextMatrix(count + 1, 3) = "0"
 MSFlexGrid1.TextMatrix(count + 1, 4) = "0"
 count = count + 1
 RS.MoveNext
Wend

RS.Close

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Sub ShowList()
On Error GoTo Err_Routine

MSFlexGridPriceList.Visible = False
ProgressBar1.Visible = True
Screen.MousePointer = vbHourglass

If optSection2.Value = True Then
 cboPriceList.Enabled = False
 ShowIngredientsCost
 lblTitle.Caption = optSection2.Caption
ElseIf optSection3.Value = True Then
 cboPriceList.Enabled = False
 ShowOverheadsCost
 lblTitle.Caption = optSection3.Caption
ElseIf optSection1.Value = True Then
 cboPriceList.Enabled = True
 LoadPriceList
 lblTitle.Caption = optSection1.Caption
End If

ProgressBar1.Visible = False
Screen.MousePointer = vbDefault
MSFlexGridPriceList.Visible = True

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub





Private Sub cmdRemove_Click()
On Error GoTo Err_Routine

OldDB.Execute "Delete from MenuCost Where Menu = '" & cboMenuName.Text & "' and Ingredient = '" & lvwIncludedIngred.SelectedItem.Text & "'"
MsgBox "Deleted", vbOKOnly, App.Title
LoadMenuIngredients

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_Routine

Dim RS As New ADODB.Recordset

If ValidateForm2 = False Then Exit Sub

RS.Open "Select Name from OverheadInfo where Name = '" & txtName.Text & "'", OldDB

If RS.EOF Then
 OldDB.Execute "Insert Into OverheadInfo (Name,Amount,OverheadClass) Values ('" & txtName.Text & "'," & txtAmount.Text & ",'" & cboClass.Text & "')"
Else
 OldDB.Execute "Update OverheadInfo Set Amount = " & txtAmount.Text & ", OverheadClass = '" & cboClass.Text & "' where Name = '" & txtName.Text & "'"
End If

 MsgBox "Saved", vbOKOnly, App.Title
 txtAmount.Text = ""
 LoadCbo txtName, "Select Distinct Name from OverheadInfo order by Name ASC"

 lblTotalOver_Click

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub
Function ValidateForm2() As Boolean

If txtName.Text = "" Then
MsgBox "Enter The Overhead Name", vbOKOnly, App.Title
ValidateForm2 = False
Exit Function
End If

If txtAmount.Text = "" Then
MsgBox "Enter The Overhead Amount", vbOKOnly, App.Title
ValidateForm2 = False
Exit Function
End If

If Not IsNumeric(txtAmount.Text) Then
MsgBox "The Overhead Amount Must Be A Numeric Value", vbOKOnly, App.Title
ValidateForm2 = False
Exit Function
End If

If txtAmount.Text <= 0 Then
MsgBox "Overhead Amount Must Be Greater Than Zero (0)", vbOKOnly, App.Title
ValidateForm2 = False
Exit Function
End If

If cboClass.Text = "" Then
MsgBox "Select Overhead Class", vbOKOnly, App.Title
ValidateForm2 = False
Exit Function
End If

ValidateForm2 = True

End Function


Private Sub cmdSaveApportionment_Click()
Dim i As Integer

OldDB.Execute "Delete from Apportionment"

For i = 1 To MSFlexGrid1.Rows - 1
 OldDB.Execute "Insert Into Apportionment (Category,Menu,Apportionment,SalesPerMonth,Weight) Values('" & MSFlexGrid1.TextMatrix(i, 0) & "','" & MSFlexGrid1.TextMatrix(i, 1) & "'," & MSFlexGrid1.TextMatrix(i, 2) & "," & MSFlexGrid1.TextMatrix(i, 3) & "," & MSFlexGrid1.TextMatrix(i, 4) & ")"
Next i

End Sub






Private Sub cmdSaveIngred_Click()
On Error GoTo Err_Routine

Dim RS As New ADODB.Recordset

If ValidateForm = False Then Exit Sub

RS.Open "Select Name from Ingredients where Name = '" & txtIngredName.Text & "'", OldDB

If RS.EOF Then
 OldDB.Execute "Insert Into Ingredients (Name,CostPrice,Amount,UnitofMeasurement) Values ('" & txtIngredName.Text & "'," & txtPrice.Text & "," & txtIngredAmount.Text & ",'" & cboIngredUnits.Text & "')"
 OldDB.Execute "Insert Into Apportionment (Name,Consumption) Values('" & txtIngredName.Text & "',0)"
Else
 OldDB.Execute "Update Ingredients Set CostPrice = " & txtPrice.Text & ", Amount = " & txtIngredAmount.Text & ",UnitofMeasurement = '" & cboIngredUnits.Text & "' where Name = '" & txtIngredName.Text & "'"
End If
 
 OldDB.Execute "Insert Into StockAdjustment (ProductDesc,ADate,Qty,Reason,Cost) Values ('" & txtIngredName.Text & "','" & Date & "','" & txtIngredAmount.Text & "','Purchase', " & txtPrice.Text & ")"

 MsgBox "Saved", vbOKOnly, App.Title
 LoadCbo txtIngredName, "Select Distinct Name from Ingredients order by Name ASC"
 LoadCbo cboIngredient, "Select Distinct Name from Ingredients order by Name ASC"

 ConfigMSFlexGrid1

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Function ValidateForm() As Boolean
If txtIngredName.Text = "" Then
MsgBox "Select/Enter The Ingredient Name", vbOKOnly, App.Title
ValidateForm = False
Exit Function
End If

If txtPrice.Text = "" Then
MsgBox "Enter The Ingredient Price", vbOKOnly, App.Title
ValidateForm = False
Exit Function
End If

If txtPrice.Text < 0 Then
MsgBox "Price Should Be Greater Than or Equal to Zero (0)", vbOKOnly, App.Title
ValidateForm = False
Exit Function
End If

If Not IsNumeric(txtPrice.Text) Then
MsgBox "The Ingredient Price Must Be A Numeric Value", vbOKOnly, App.Title
ValidateForm = False
Exit Function
End If

If txtIngredAmount.Text = "" Then
MsgBox "Enter Amount Of Ingredient Purchased", vbOKOnly, App.Title
ValidateForm = False
Exit Function
End If

If Not IsNumeric(txtIngredAmount.Text) Then
MsgBox "The Amount Must Be A Numeric Value", vbOKOnly, App.Title
ValidateForm = False
Exit Function
End If

If txtIngredAmount.Text <= 0 Then
MsgBox "Amount Purchased Must Be Greater Than Zero (0)", vbOKOnly, App.Title
ValidateForm = False
Exit Function
End If

If cboIngredUnits.Text = "" Then
MsgBox "Select Ingredient Units", vbOKOnly, App.Title
ValidateForm = False
Exit Function
End If

ValidateForm = True

End Function




Private Sub cmdSavePrice_Click()
On Error GoTo Err_Routine

Dim RS As New ADODB.Recordset
Dim SQL As String

If ValidateForm3 = False Then Exit Sub

RS.Open "Select * from ProductInfo where Menu = '" & cboMenuName.Text & "' and Category = '" & cboMenuCategory & "'", OldDB

If RS.EOF Then
 SQL = "Insert Into ProductInfo (Menu,Category,NumberPple,DOC) Values ('" & cboMenuName.Text & "','" & cboMenuCategory.Text & "'," & txtNumberPple.Text & ",'" & Date & "')"
Else
 SQL = "Update ProductInfo Set Category ='" & cboMenuCategory.Text & "', "
 SQL = SQL & "NumberPple = " & txtNumberPple.Text
 SQL = SQL & " where Menu = '" & cboMenuName.Text & "'"
End If

OldDB.Execute SQL
RS.Close

MsgBox "Saved", vbOKOnly, App.Title

cboMenuCategory_Click

LoadCbo cboPriceList, "Select Distinct Category from Categorys"
cboPriceList.ListIndex = 0
 
Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub
Function ValidateForm3() As Boolean

If cboMenuCategory.Text = "" Then
MsgBox "Select Menu Category", vbOKOnly, App.Title
ValidateForm3 = False
Exit Function
End If

If cboMenuName.Text = "" Then
MsgBox "Enter/Select Menu Name", vbOKOnly, App.Title
ValidateForm3 = False
Exit Function
End If

If txtNumberPple.Text = "" Then
MsgBox "Enter The Number Of People Served", vbOKOnly, App.Title
ValidateForm3 = False
Exit Function
End If

If Not IsNumeric(txtNumberPple.Text) Then
MsgBox "The Number Of People Served Must Be A Numeric Value", vbOKOnly, App.Title
ValidateForm3 = False
Exit Function
End If

If txtNumberPple.Text <= 0 Then
MsgBox "The Number Of People Served Should Be Greater Than Zero (0)", vbOKOnly, App.Title
ValidateForm3 = False
Exit Function
End If



ValidateForm3 = True

End Function







Private Sub cmdShow_Click()
On Error GoTo Err_Routine
Dim MyCosting As New CCosting

ProgressBar4.Visible = True
Screen.MousePointer = vbHourglass

 lblTotalIngred.Caption = "Cost of Production = " & MyCosting.TotalIngredCost(cboMenuName.Text)
 ProgressBar4.Value = 1
 lblApportionedOver.Caption = "Aportioned Overhead =" & MyCosting.ApportionedOverhead(cboMenuName.Text)
 ProgressBar4.Value = 2
 lblCostPerIndividual.Caption = "Cost Per Individual =" & MyCosting.CostPerIndividual(cboMenuName.Text)
 ProgressBar4.Value = 3

Screen.MousePointer = vbDefault
ProgressBar4.Visible = False

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Err_Routine
Dim SQL As String
Dim i As Integer

ProgressBar1.Visible = True

With MSFlexGridPriceList

If lblTitle.Caption = "Ingredients" Then OldDB.Execute "Delete from PriceList where Category = '" & cboPriceList.Text & "'"

For i = 1 To .Rows - 1

If lblTitle.Caption = "Ingredients" Then
 SQL = "Update Ingredients Set CostPrice = " & .TextMatrix(i, 3) & ", Amount = " & .TextMatrix(i, 2) & " where Name = '" & .TextMatrix(i, 0) & "'"
ElseIf lblTitle.Caption = "Overheads" Then
 SQL = "Update OverheadInfo Set Amount = " & .TextMatrix(i, 2) & " where Name = '" & .TextMatrix(i, 0) & "' and OverheadClass = '" & .TextMatrix(i, 1) & "'"
Else
 OldDB.Execute "Insert Into PriceList (Category,ProductDesc,Price) Values ('" & cboPriceList.Text & "','" & .TextMatrix(i, 0) & "'," & .TextMatrix(i, 4) & ")"
End If

 OldDB.Execute SQL
 ProgressBar1.Value = ProgressBar1.Max * (i / MSFlexGridPriceList.Rows)

Next i

End With

ProgressBar1.Visible = False

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
 
 
 
End Sub

Private Sub Form_Load()
On Error GoTo Err_Routine
 
 LoadCbo cboMenuCategory, "Select Distinct Category from Categorys"
 LoadCbo cboPriceList, "Select Distinct Category from Categorys"
 cboPriceList.ListIndex = 0

 LoadUnitsOfMeasurement
 LoadCbo cboClass, "Select Distinct Overhead from Overheads"

 LoadCbo txtName, "Select Distinct Name from OverheadInfo order by Name ASC"

 LoadCbo txtIngredName, "Select Distinct Name from Ingredients order by Name ASC"
 LoadCbo cboIngredient, "Select Distinct Name from Ingredients order by Name ASC"

 lblTotalOver_Click
 
 ConfigMSFlexGrid1
 ShowList

Exit_Routine:
 Exit Sub
Err_Routine:
 MsgBox Err.Description, , App.Title
 Resume Exit_Routine
End Sub


Sub ConfigMSFlexGrid1()
Dim RS As New ADODB.Recordset
Dim count As Integer

RS.Open "Select Menu,Category,Weight,SalesPerMonth,Apportionment from Apportionment order by Category ASC", OldDB

 MSFlexGrid1.Rows = RS.RecordCount + 1
 MSFlexGrid1.ColWidth(0) = 2000
 MSFlexGrid1.ColWidth(1) = 2200
 MSFlexGrid1.ColWidth(2) = 1500
 
 MSFlexGrid1.TextMatrix(0, 0) = "Category"
 MSFlexGrid1.TextMatrix(0, 1) = "Product"
 MSFlexGrid1.TextMatrix(0, 2) = "Apportionment"
 MSFlexGrid1.TextMatrix(0, 3) = "Sales"
 MSFlexGrid1.TextMatrix(0, 4) = "Weight"

While Not RS.EOF
 MSFlexGrid1.TextMatrix(count + 1, 0) = RS!Category
 MSFlexGrid1.TextMatrix(count + 1, 1) = RS!Menu
 MSFlexGrid1.TextMatrix(count + 1, 2) = RS!Apportionment
 MSFlexGrid1.TextMatrix(count + 1, 3) = RS!SalesPerMonth
 MSFlexGrid1.TextMatrix(count + 1, 4) = RS!Weight
 count = count + 1
 RS.MoveNext
Wend

RS.Close

End Sub




Sub LoadUnitsOfMeasurement()
 cboIngredUnits.AddItem "Kilograms"
 cboIngredUnits.AddItem "Grams"
 cboIngredUnits.AddItem "Litres"
 cboIngredUnits.AddItem "Mililitres"
 cboIngredUnits.AddItem "Units"
 
 cboUnits.AddItem "Kilograms"
 cboUnits.AddItem "Grams"
 cboUnits.AddItem "Litres"
 cboUnits.AddItem "Mililitres"
 cboUnits.AddItem "Units"
End Sub

Sub LoadMenuIngredients()
Dim RS As New ADODB.Recordset
Dim lstmain1 As ListItem
Dim MyCosting As New CCosting
lvwIncludedIngred.ListItems.Clear

RS.Open "Select * from MenuCost Where Menu = '" & cboMenuName.Text & "'", OldDB

While Not RS.EOF
 Set lstmain1 = lvwIncludedIngred.ListItems.Add(, , RS!Ingredient)
 If Not IsNull(RS!AmountUsed) Then lstmain1.SubItems(1) = RS!AmountUsed
 If Not IsNull(RS!Units) Then lstmain1.SubItems(2) = RS!Units
 If Not IsNull(RS!Ingredient) Then lstmain1.SubItems(3) = MyCosting.PricePerSmallestUnit(RS!Ingredient)
 RS.MoveNext
Wend

End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLists.ForeColor = &H996533
lblCosts.ForeColor = &H996533
lblMenus.ForeColor = &H996533
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTotalOver.ForeColor = &HA67D46
End Sub

Private Sub lblCosts_Click()
FrameMenus.Visible = False
FrameLists.Visible = False
FrameCosts.Visible = True
End Sub

Private Sub lblCosts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCosts.ForeColor = &H80FF&
End Sub

Private Sub lblLists_Click()
FrameMenus.Visible = False
FrameLists.Visible = True
FrameCosts.Visible = False
End Sub

Private Sub lblLists_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLists.ForeColor = &H80FF&
End Sub

Private Sub lblMenus_Click()
FrameMenus.Visible = True
FrameLists.Visible = False
FrameCosts.Visible = False
End Sub

Private Sub lblMenus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMenus.ForeColor = &H80FF&
End Sub

Private Sub lblTotalOver_Click()
Dim MyCosting As New CCosting

lblTotalOver.Caption = MyCosting.TotalOverheads
End Sub

Private Sub lblTotalOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTotalOver.ForeColor = &H80FF&
End Sub

Private Sub lvwIncludedIngred_ItemCheck(ByVal Item As MSComctlLib.ListItem)
' Set the variable to the SelectedItem.
Set lvwIncludedIngred.SelectedItem = Item
    
If Item.Checked Then
 CheckLoadedItems lvwIncludedIngred, lvwIncludedIngred.SelectedItem.Index
Else
 UncheckLoadedItems lvwIncludedIngred, lvwIncludedIngred.SelectedItem.Index
End If
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
   With MSFlexGrid1
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


Private Sub MSFlexGrid1_LeaveCell()
With MSFlexGrid1
   If .TextMatrix(.Row, .Col) = "" Then Exit Sub

   If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
      MsgBox "Numeric Value Expected", vbOKOnly, App.Title
      .TextMatrix(.Row, .Col) = "0"
   Else
   If .TextMatrix(.Row, 3) > 0 Then .TextMatrix(.Row, 2) = Round((lblTotalOver.Caption * (.TextMatrix(.Row, 4) / 100)) / .TextMatrix(.Row, 3), 0)
   End If
End With
End Sub

Private Sub MSFlexGridPriceList_KeyPress(KeyAscii As Integer)
    With MSFlexGridPriceList
    If .Col <> 4 And optSection1.Value = True Then Exit Sub
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


Private Sub MSFlexGridPriceList_LeaveCell()
With MSFlexGridPriceList
   If .TextMatrix(.Row, .Col) = "" Then Exit Sub

   If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
      MsgBox "Numeric Value Expected", vbOKOnly, App.Title
      .TextMatrix(.Row, .Col) = "0"
   End If
End With
End Sub


Private Sub optSection1_Click()
ShowList
End Sub

Private Sub optSection2_Click()
ShowList
End Sub

Private Sub optSection3_Click()
ShowList
End Sub

Private Sub txtIngredName_Click()
Dim RS As New ADODB.Recordset

RS.Open "Select * from Ingredients where Name = '" & txtIngredName.Text & "'", OldDB

If RS.EOF Then Exit Sub
 txtPrice.Text = RS!CostPrice
 txtIngredAmount.Text = RS!Amount
 SetcboTo cboIngredUnits, RS!UnitofMeasurement

End Sub



Private Sub txtName_Click()
Dim RS As New ADODB.Recordset

RS.Open "Select * from OverheadInfo where Name = '" & txtName.Text & "'", OldDB
If RS.EOF Then Exit Sub

txtAmount.Text = RS!Amount
SetcboTo cboClass, RS!OverheadClass

End Sub

Sub LoadPriceList()
Dim RS As New ADODB.Recordset
Dim i As Integer
Dim MyCosting As New CCosting

RS.Open "Select Distinct Menu from ProductInfo where Category = '" & cboPriceList.Text & "'", OldDB
MSFlexGridPriceList.Rows = RS.RecordCount + 1
 
 MSFlexGridPriceList.Cols = 5
 MSFlexGridPriceList.TextMatrix(0, 0) = "Product"
 MSFlexGridPriceList.TextMatrix(0, 1) = "Cost of Production"
 MSFlexGridPriceList.TextMatrix(0, 2) = "Aportioned Overhead"
 MSFlexGridPriceList.TextMatrix(0, 3) = "Cost Per Individual"
 MSFlexGridPriceList.TextMatrix(0, 4) = "Final Selling Price"
 
 MSFlexGridPriceList.ColWidth(0) = 2000
 MSFlexGridPriceList.ColWidth(1) = 2000
 MSFlexGridPriceList.ColWidth(2) = 2000
 MSFlexGridPriceList.ColWidth(3) = 2000
 MSFlexGridPriceList.ColWidth(4) = 2000

i = 1
While Not RS.EOF
If Not IsNull(RS!Menu) Then
 MSFlexGridPriceList.TextMatrix(i, 0) = RS!Menu
 MSFlexGridPriceList.TextMatrix(i, 1) = MyCosting.TotalIngredCost(RS!Menu)
 MSFlexGridPriceList.TextMatrix(i, 2) = MyCosting.ApportionedOverhead(RS!Menu)
 MSFlexGridPriceList.TextMatrix(i, 3) = MyCosting.CostPerIndividual(RS!Menu)
 MSFlexGridPriceList.TextMatrix(i, 4) = MyCosting.GetSellingPrice(RS!Menu)
End If
 RS.MoveNext
 ProgressBar1.Value = ProgressBar1.Max * (i / RS.RecordCount)
 i = i + 1
 
Wend

End Sub


Sub ShowIngredientsCost()
Dim RS As New ADODB.Recordset
Dim i As Integer

RS.Open "Select Name,UnitofMeasurement,CostPrice,Amount from Ingredients Order By Name", OldDB
MSFlexGridPriceList.Rows = RS.RecordCount + 1

MSFlexGridPriceList.Cols = 4
MSFlexGridPriceList.FixedCols = 2

MSFlexGridPriceList.TextMatrix(0, 0) = "Name"
MSFlexGridPriceList.TextMatrix(0, 1) = "Unit Of Measurement"
MSFlexGridPriceList.TextMatrix(0, 2) = "Units"
MSFlexGridPriceList.TextMatrix(0, 3) = "Cost Price"

MSFlexGridPriceList.ColWidth(0) = 3000
MSFlexGridPriceList.ColWidth(1) = 2000
MSFlexGridPriceList.ColWidth(2) = 2000
MSFlexGridPriceList.ColWidth(3) = 2000

i = 1
 
While Not RS.EOF
If Not IsNull(RS!Name) Then
 MSFlexGridPriceList.TextMatrix(i, 0) = RS!Name
 MSFlexGridPriceList.TextMatrix(i, 1) = RS!UnitofMeasurement
 MSFlexGridPriceList.TextMatrix(i, 2) = RS!Amount
 MSFlexGridPriceList.TextMatrix(i, 3) = RS!CostPrice
End If
 RS.MoveNext
 ProgressBar1.Value = ProgressBar1.Max * (i / RS.RecordCount)
 i = i + 1
Wend

End Sub


Sub ShowOverheadsCost()
Dim RS As New ADODB.Recordset
Dim i As Integer

RS.Open "Select Name,OverheadClass,Amount from OverheadInfo Order By Name", OldDB
MSFlexGridPriceList.Rows = RS.RecordCount + 1

MSFlexGridPriceList.Cols = 3
MSFlexGridPriceList.FixedCols = 2

MSFlexGridPriceList.TextMatrix(0, 0) = "Name Of Overhead"
MSFlexGridPriceList.TextMatrix(0, 1) = "Overhead Class"
MSFlexGridPriceList.TextMatrix(0, 2) = "Overhead Cost"
  
MSFlexGridPriceList.ColWidth(0) = 3000
MSFlexGridPriceList.ColWidth(1) = 3000
MSFlexGridPriceList.ColWidth(2) = 3000

i = 1
 
While Not RS.EOF
If Not IsNull(RS!Name) Then
 MSFlexGridPriceList.TextMatrix(i, 0) = RS!Name
 MSFlexGridPriceList.TextMatrix(i, 1) = RS!OverheadClass
 MSFlexGridPriceList.TextMatrix(i, 2) = RS!Amount
End If
 RS.MoveNext
 ProgressBar1.Value = ProgressBar1.Max * (i / RS.RecordCount)
 i = i + 1
Wend

End Sub


