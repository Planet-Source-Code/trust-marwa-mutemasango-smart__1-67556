VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim TotalApportionedOverhead As Currency
Dim CostOfIngredients As Currency

Function GetUnitofMeasurement(Ingredient As String)
Dim RS As New ADODB.Recordset

RS.Open "Select * from Ingredients where Name = '" & Ingredient & "'", OldDB
If Not IsNull(RS!UnitofMeasurement) And Not RS.EOF Then GetUnitofMeasurement = RS!UnitofMeasurement

End Function

Function GetAmtUsed(MenuStr As String, Ingredient As String) As Integer
Dim RS As New ADODB.Recordset

RS.Open "SELECT MenuCost.AmountUsed FROM Ingredients INNER JOIN MenuCost ON Ingredients.Name = MenuCost.Ingredient Where MenuCost.Menu = '" & MenuStr & "' and MenuCost.Ingredient = '" & Ingredient & "'", OldDB

If Not RS.EOF Then GetAmtUsed = RS!AmountUsed

End Function

Function ApportionedOverhead(MenuStr As String) As Currency
Dim RS As New ADODB.Recordset

RS.Open "Select Apportionment from Apportionment where Menu ='" & MenuStr & "'", OldDB

If Not RS.EOF And Not IsNull(RS!Apportionment) Then ApportionedOverhead = RS!Apportionment

End Function

Function TotalIngredCost(MenuStr As String) As Currency
Dim RS As New ADODB.Recordset
Dim c As Currency

RS.Open "SELECT Ingredients.*,MenuCost.Ingredient, MenuCost.AmountUsed,MenuCost.Units FROM Ingredients INNER JOIN MenuCost ON Ingredients.Name = MenuCost.Ingredient Where MenuCost.Menu = '" & MenuStr & "'", OldDB

While Not RS.EOF

If RS!UnitofMeasurement = "Kilograms" Then
c = (RS!CostPrice / 1000) / RS!Amount
ElseIf RS!UnitofMeasurement = "Litres" Then
c = (RS!CostPrice / 1000) / RS!Amount
Else
c = (RS!CostPrice / RS!Amount)
End If

If RS!Units = "Kilograms" Then
CostOfIngredients = CostOfIngredients + ((c * RS!AmountUsed) * 1000)
ElseIf RS!Units = "Litres" Then
CostOfIngredients = CostOfIngredients + ((c * RS!AmountUsed) * 1000)
Else
CostOfIngredients = CostOfIngredients + (c * RS!AmountUsed)
End If
 
RS.MoveNext
Wend

TotalIngredCost = CostOfIngredients

End Function


Function PricePerSmallestUnit(Ingredient As String) As Currency
Dim RS As New ADODB.Recordset
RS.Open "SELECT Ingredients.CostPrice, Ingredients.Amount, Ingredients.UnitofMeasurement FROM Ingredients INNER JOIN MenuCost ON Ingredients.Name = MenuCost.Ingredient Where Ingredients.Name = '" & Ingredient & "'", OldDB

If RS.EOF Then Exit Function

If RS!UnitofMeasurement = "Kilograms" Then
PricePerSmallestUnit = (RS!CostPrice / 1000) / RS!Amount
ElseIf RS!UnitofMeasurement = "Litres" Then
PricePerSmallestUnit = (RS!CostPrice / 1000) / RS!Amount
Else
PricePerSmallestUnit = (RS!CostPrice / RS!Amount)
End If

End Function

Function CostPerIndividual(MenuStr As String) As Variant
Dim RS As New ADODB.Recordset
RS.Open "Select * from ProductInfo where Menu = '" & MenuStr & "'", OldDB

If Not RS.EOF And Not IsNull(RS!NumberPple) Then CostPerIndividual = Round((TotalIngredCost(MenuStr) + ApportionedOverhead(MenuStr)) / RS!NumberPple, 0)

End Function


Function TotalOverheads() As Currency
Dim RS As New ADODB.Recordset

RS.Open "Select Sum(Amount) as TotalOver from OverheadInfo", OldDB
If Not IsNull(RS!TotalOver) Then TotalOverheads = RS!TotalOver

End Function

Function GetSellingPrice(MenuStr As String) As Currency
Dim RS As New ADODB.Recordset

RS.Open "Select Price From Pricelist Where ProductDesc = '" & MenuStr & "'", OldDB

If Not RS.EOF Then GetSellingPrice = RS!Price

End Function
