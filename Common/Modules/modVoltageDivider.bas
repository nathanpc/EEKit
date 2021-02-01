Attribute VB_Name = "modVoltageDivider"
''' modVoltageDivider
''' A simple module to help out with voltage divider calculations.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Calculates the input voltage of a voltage divider circuit.
Public Function VoltDividerVin(dblR1 As Double, dblVout As Double, dblR2 As Double, _
        dblVgnd As Double)
    VoltDividerVin = dblVout + dblVgnd + ((dblR1 * dblVout) / dblR2)
End Function

' Calculates the output voltage of a voltage divider circuit.
Public Function VoltDividerVout(dblVin As Double, dblR1 As Double, dblR2 As Double, _
        dblVgnd As Double)
    VoltDividerVout = ((dblVin - dblVgnd) * dblR2) / (dblR1 + dblR2)
End Function

' Calculates the negative rail of a voltage divider circuit.
Public Function VoltDividerVgnd(dblVin As Double, dblR1 As Double, dblVout As Double, _
        dblR2 As Double)
    VoltDividerVgnd = dblVin - ((dblVout * (dblR1 + dblR2)) / dblR2)
End Function

' Calculates the input resistor of a voltage divider circuit.
Public Function VoltDividerR1(dblVin As Double, dblVout As Double, dblR2 As Double, _
        dblVgnd As Double)
    VoltDividerR1 = (dblR2 * (dblVin - dblVout - dblVgnd)) / dblVout
End Function

' Calculates the grounded resistor of a voltage divider circuit.
Public Function VoltDividerR2(dblVin As Double, dblR1 As Double, dblVout As Double, _
        dblVgnd As Double)
    VoltDividerR2 = (dblR1 * dblVout) / (dblVin - dblVout - dblVgnd)
End Function
