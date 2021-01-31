Attribute VB_Name = "modMagnitudeParser"
''' modMagnitudeParser
''' Parses manitude strings like "10m" into numbers like 0.01 automagically.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Parse a string with magnitude into a number.
Public Function ParseNumber(strNumber As String) As Double
    Dim strValue As String
    Dim dblNumber As Double
    
    ' Calculate the number.
    On Error Resume Next
    dblNumber = Val(strNumber)
    dblNumber = dblNumber * MagnitudeValue(Right$(strNumber, 1))
    
    ParseNumber = dblNumber
End Function

' Converts a number into a string with magnitude.
Public Function NumberToMagString(dblNumber As Double) As String
    Dim strBuffer As String
    
    NumberToMagString = strBuffer
End Function

' Gets the size of a magnitude.
Private Function MagnitudeValue(ByVal strMagChar As String) As Double
    ' Check if the magnitude is valid.
    If IsNumeric(strMagChar) Then
        MagnitudeValue = 1
        Exit Function
    End If
    
    ' Try to determine the magnitude of the character.
    Select Case strMagChar
        Case "f": MagnitudeValue = 0.000000000000001
        Case "p": MagnitudeValue = 0.000000000001
        Case "n": MagnitudeValue = 0.000000001
        Case "u": MagnitudeValue = 0.000001
        Case "m": MagnitudeValue = 0.001
        Case "c": MagnitudeValue = 0.01
        Case "d": MagnitudeValue = 0.1
        Case "D": MagnitudeValue = 10#
        Case "h": MagnitudeValue = 100#
        Case "k": MagnitudeValue = 1000#
        Case "M": MagnitudeValue = 1000000#
        Case "G": MagnitudeValue = 1000000000#
        Case "T": MagnitudeValue = 1000000000000#
        Case Else: MagnitudeValue = 0
    End Select
End Function
