Attribute VB_Name = "variables"
Option Explicit
Public OldDB As New ADODB.Connection
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Const conSwNormal = 1
Public UserName As String
Public SuperUser As Boolean

Public Sub FormDrag(TheForm As Form)
 ReleaseCapture
 SendMessage TheForm.hwnd, &HA1, 2, 0&
End Sub

Sub Main()
SuperUser = False
frmLogin.Show
End Sub

Sub SetcboTo(cbo As ComboBox, ID As String)
Dim i As Integer

For i = 0 To cbo.ListCount - 1
 If cbo.List(i) = ID Then cbo.ListIndex = i
Next i

End Sub

Sub LoadCbo(cbo As ComboBox, SQL As String)
Dim RS As New ADODB.Recordset

RS.Open SQL, OldDB

cbo.Clear
While Not RS.EOF
  If Not IsNull(RS.Fields(0).Value) Then cbo.AddItem RS.Fields(0).Value
  RS.MoveNext
Wend

End Sub



Function GetStockLevel(MenuStr As String, CategoryStr As String, SectionStr As String) As Currency
Dim RS As New ADODB.Recordset

RS.Open "Select Levels From StockLevels Where ProductDesc = '" & MenuStr & "' and [Section] ='" & SectionStr & "' and [Category] = '" & CategoryStr & "'", OldDB

If Not RS.EOF Then GetStockLevel = RS!Levels

End Function

Function GetDailyIssues(MenuStr As String, CategoryStr As String, SectionStr As String) As Currency
Dim RS As New ADODB.Recordset

RS.Open "Select Sum (UnitsIssued) as num From DailyIssues Where ProductDesc = '" & MenuStr & "' and [Section] ='" & SectionStr & "' and [Category] = '" & CategoryStr & "' and DateOfIssue = '" & Format(Date, "dd/mm/yy") & "'", OldDB

If Not RS.EOF And Not IsNull(RS!num) Then GetDailyIssues = RS!num

End Function


Function GetCoInfo() As String
Dim Temp As String
 Temp = GetSetting(App.Title, "Settings", "txtCoName", "") & vbNewLine
 Temp = Temp & "Telephone: " & GetSetting(App.Title, "Settings", "txtTel", "") & vbNewLine
 Temp = Temp & "E-Mail: " & GetSetting(App.Title, "Settings", "txtEmail", "") & vbNewLine
 Temp = Temp & "Fax: " & GetSetting(App.Title, "Settings", "txtFax", "") & vbNewLine
 Temp = Temp & "Physical Address: " & GetSetting(App.Title, "Settings", "txtAddress", "") & vbNewLine
End Function

Function GetBookingCode(CusNameStr As String) As String
Dim RS As New ADODB.Recordset

 RS.Open "SELECT ID from bookings where checkedout = 0 and CusName = '" & CusNameStr & "'", OldDB
 
 If Not IsNull(RS!ID) And Not RS.EOF Then GetBookingCode = RS!ID
 
End Function

Sub bug(str As String)
 Open "C:\fname.txt" For Output As #1
 Print #1, str
 Close (1)
End Sub

Sub UncheckLoadedItems(lvwinfo As ListView, i As Integer)
Dim X As Integer

   lvwinfo.ListItems.Item(i).Bold = False
   lvwinfo.ListItems.Item(i).ForeColor = vbBlack
   lvwinfo.ListItems.Item(i).Checked = False
         
   For X = 1 To lvwinfo.ListItems.Item(i).ListSubItems.count
     lvwinfo.ListItems.Item(i).ListSubItems(X).Bold = False
     lvwinfo.ListItems.Item(i).ListSubItems(X).ForeColor = vbBlack
   Next X

End Sub

Sub CheckLoadedItems(lvwinfo As ListView, i As Integer)
Dim X As Integer

   lvwinfo.ListItems.Item(i).Bold = True
   lvwinfo.ListItems.Item(i).ForeColor = &HA67D46
   lvwinfo.ListItems.Item(i).Checked = True
         
   For X = 1 To lvwinfo.ListItems.Item(i).ListSubItems.count
     lvwinfo.ListItems.Item(i).ListSubItems(X).Bold = True
     lvwinfo.ListItems.Item(i).ListSubItems(X).ForeColor = &HA67D46
   Next X

End Sub

Sub UnCheckItems(lvwinfo As ListView)
Dim i As Integer
Dim X As Integer

   For i = 1 To lvwinfo.ListItems.count
        lvwinfo.ListItems.Item(i).Bold = False
        lvwinfo.ListItems.Item(i).ForeColor = vbBlack
        lvwinfo.ListItems.Item(i).Checked = False
        For X = 1 To lvwinfo.ListItems.Item(i).ListSubItems.count
         lvwinfo.ListItems.Item(i).ListSubItems(X).Bold = False
         lvwinfo.ListItems.Item(i).ListSubItems(X).ForeColor = vbBlack
        Next X
   Next i
End Sub

Sub CheckItems(lvwinfo As ListView)
Dim i As Integer
Dim X As Integer

   For i = 1 To lvwinfo.ListItems.count
        lvwinfo.ListItems.Item(i).Bold = True
        lvwinfo.ListItems.Item(i).ForeColor = &HA67D46
        lvwinfo.ListItems.Item(i).Checked = True
        
        For X = 1 To lvwinfo.ListItems.Item(i).ListSubItems.count
         lvwinfo.ListItems.Item(i).ListSubItems(X).Bold = True
         lvwinfo.ListItems.Item(i).ListSubItems(X).ForeColor = &HA67D46
        Next X
   Next i
End Sub

Function Section(frm As Form) As String

If frm.optSection1.Value = True Then
 Section = frm.optSection1.Caption
ElseIf frm.optSection2.Value = True Then
 Section = frm.optSection2.Caption
ElseIf frm.optSection3.Value = True Then
 Section = frm.optSection3.Caption
ElseIf frm.optSection4.Value = True Then
 Section = frm.optSection4.Caption
End If

End Function

Public Sub LoadReports(lvwinfo As ListView, SQL As String)
Dim RS As New ADODB.Recordset
Dim X As Integer
Dim lstmain1 As ListItem

RS.Open SQL, OldDB
lvwinfo.Visible = True
lvwinfo.ColumnHeaders.Clear

 For X = 0 To RS.Fields.count - 1
  lvwinfo.ColumnHeaders.Add , , Replace(RS.Fields(X).Name, "'", "")
 Next X
 
lvwinfo.ListItems.Clear
 While Not RS.EOF
   If Not IsNull(RS.Fields(0).Value) Then
    Set lstmain1 = lvwinfo.ListItems.Add(, , RS.Fields(0).Value)
   Else
    Set lstmain1 = lvwinfo.ListItems.Add(, , "")
   End If
   
    For X = 1 To RS.Fields.count - 1
    If Not IsNull(RS.Fields(X).Value) Then lstmain1.SubItems(X) = Replace(RS.Fields(X).Value, vbNewLine, " ")
   Next X

   RS.MoveNext
  Wend
  
End Sub
