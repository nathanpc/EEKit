Attribute VB_Name = "modOhmsLaw"
''' modOhmsLaw
''' A simple module to help out with Ohm's Law type of calculations.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Calculates the voltage.
Public Function OhmsLawVoltage(dblCurrent As Double, dblResistance As Double, _
        dblPower As Double) As Double
    If dblCurrent = 0 Then
        OhmsLawVoltage = Sqr(dblPower * dblResistance)
    ElseIf dblResistance = 0 Then
        OhmsLawVoltage = dblPower / dblCurrent
    ElseIf dblPower = 0 Then
        OhmsLawVoltage = dblCurrent * dblResistance
    End If
End Function

' Calculates the current.
Public Function OhmsLawCurrent(dblVoltage As Double, dblResistance As Double, _
        dblPower As Double) As Double
    If dblVoltage = 0 Then
        OhmsLawCurrent = Sqr(dblPower / dblResistance)
    ElseIf dblResistance = 0 Then
        OhmsLawCurrent = dblPower / dblVoltage
    ElseIf dblPower = 0 Then
        OhmsLawCurrent = dblVoltage / dblResistance
    End If
End Function

' Calculates the resistance.
Public Function OhmsLawResistance(dblVoltage As Double, dblCurrent As Double, _
        dblPower As Double) As Double
    If dblVoltage = 0 Then
        OhmsLawResistance = dblPower / (dblCurrent * dblCurrent)
    ElseIf dblCurrent = 0 Then
        OhmsLawResistance = (dblVoltage * dblVoltage) / dblPower
    ElseIf dblPower = 0 Then
        OhmsLawResistance = dblVoltage / dblCurrent
    End If
End Function

' Calculates the Power
Public Function OhmsLawPower(dblVoltage As Double, dblCurrent As Double, _
        dblResistance As Double) As Double
    If dblVoltage = 0 Then
        OhmsLawPower = dblResistance * (dblCurrent * dblCurrent)
    ElseIf dblCurrent = 0 Then
        OhmsLawPower = (dblVoltage * dblVoltage) / dblResistance
    ElseIf dblResistance = 0 Then
        OhmsLawPower = dblVoltage * dblCurrent
    End If
End Function

