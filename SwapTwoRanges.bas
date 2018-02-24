Attribute VB_Name = "SwapTwoRanges"
Sub SwapTwoRanges()

Dim Rng1 As Range, rng2 As Range
Dim arr1 As Variant, arr2 As Variant
xTitleId = "Change range"
On Error Resume Next
Set Rng1 = Application.Selection
Set Rng1 = Application.InputBox("Range1:", xTitleId, Rng1.Address, Type:=8)
Rng1.Interior.ColorIndex = 3
 
Set rng2 = Application.InputBox("Range2:", xTitleId, Type:=8)
rng2.Interior.ColorIndex = 15

'Application.ScreenUpdating = False
arr1 = Rng1.Value
arr2 = rng2.Value
Rng1.Interior.ColorIndex = 0
rng2.Interior.ColorIndex = 0
Rng1.Value = arr2
rng2.Value = arr1

End Sub
