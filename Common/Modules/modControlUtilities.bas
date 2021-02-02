Attribute VB_Name = "modControlUtilities"
''' modControlUtilities
''' A helper module with a whole bunch of utilities related to controls.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Gets the number from a TextBox.
Public Function GetTextNum(txtBox As TextBox) As Double
    Dim dblValue As Double
    On Error Resume Next
    
    dblValue = ParseNumber(txtBox.Text)
    GetTextNum = dblValue
End Function
