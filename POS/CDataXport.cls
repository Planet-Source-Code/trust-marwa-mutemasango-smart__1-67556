VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataXport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FlexgridXport(FlexGrid As MSFlexGrid, dlgCommonDialog As CommonDialog)
Dim sFile As String

sFile = xPortFilename(dlgCommonDialog)
If Len(sFile) > 0 Then FlexGridSave sFile, FlexGrid
End Sub

Public Sub lvwXport(lvwinfo As ListView, dlgCommonDialog As CommonDialog)
Dim sFile As String

sFile = xPortFilename(dlgCommonDialog)
If Len(sFile) > 0 Then lvwSave sFile, lvwinfo
End Sub

Function xPortFilename(dlgCommonDialog As CommonDialog) As String
    With dlgCommonDialog
         .DialogTitle = "Save Data"
         .CancelError = False
         .DefaultExt = "*.csv"
         .Filter = "CSV(MS-DOS)|*.csv"
         .ShowSave
         If Len(.Filename) = 0 Then
            Exit Function
         End If
         xPortFilename = .Filename
    End With
End Function

Sub lvwSave(fname As String, lvwinfo As ListView)
 Dim i As Long
 Dim j As Long
 Dim outputX As String
 
Open fname For Output As #1

For i = 1 To lvwinfo.ColumnHeaders.Count
 outputX = outputX & lvwinfo.ColumnHeaders(i).Text & ","
Next i
Print #1, outputX
         
For i = 1 To lvwinfo.ListItems.Count
 outputX = lvwinfo.ListItems.Item(i) & ","
 For j = 1 To lvwinfo.ColumnHeaders.Count - 1
   outputX = outputX & CStr(lvwinfo.ListItems.Item(i).SubItems(j)) & ","
 Next j
 Print #1, outputX
Next i

Close (1)

End Sub

Sub FlexGridSave(fname As String, FlexGrid As MSFlexGrid)
 Dim i As Long
 Dim j As Long
 Dim outputX As String
 

Open fname For Output As #1

For i = 0 To FlexGrid.Rows - 1
 For j = 0 To FlexGrid.Cols - 1
   outputX = outputX & FlexGrid.TextMatrix(i, j) & ","
 Next j
   Print #1, outputX
   outputX = ""
Next i
Close (1)

End Sub
