VERSION 5.00
Begin VB.Form frmOhmsLaw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ohm's Law"
   ClientHeight    =   1605
   ClientLeft      =   7140
   ClientTop       =   5250
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4545
   Begin VB.CheckBox chkPower 
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chkResistance 
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkCurrent 
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkVoltage 
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picHelp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      Picture         =   "frmOhmsLaw.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   1270
      Width           =   255
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtPower 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtResistance 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtCurrent 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtVoltage 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Power (W): "
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Resistance (R): "
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Current (A): "
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Voltage (V): "
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmOhmsLaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmOhmsLaw
''' Ohm's Law applicationf for the EEKit suite.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Shows a nice help dialog.
Public Sub ShowHelp()
    MsgBox "Enter only two fields and either press Return or click on the " & _
        "Calculate button. " & vbCrLf & vbCrLf & "You always need to leave 2 " & _
        "fields blank so that the calculator knows what you want to " & _
        "calculate. You can also use the check boxes to auto-clear a field " & _
        "before calculating.", vbOKOnly + vbInformation, "Help"
End Sub

' Performs the calculations.
Public Sub Calculate()
    ' Automatically clear some fields.
    AutoClearFields
    
    ' Check empty field requirements.
    If EmptyFields <> 2 Then
        MsgBox "To perform a calculation we need to have exactly 2 empty fields.", _
            vbOKOnly + vbExclamation, "Invalid Input"
        Exit Sub
    End If
    
    ' Calculate voltage.
    If txtVoltage.Text = vbNullString Then
        txtVoltage.Text = NumberToMagString(OhmsLawVoltage(GetNum(txtCurrent), _
            GetNum(txtResistance), GetNum(txtPower)))
    End If
    
    ' Calculate current.
    If txtCurrent.Text = vbNullString Then
        txtCurrent.Text = NumberToMagString(OhmsLawCurrent(GetNum(txtVoltage), _
            GetNum(txtResistance), GetNum(txtPower)))
    End If
    
    ' Calculate resistance.
    If txtResistance.Text = vbNullString Then
        txtResistance.Text = NumberToMagString(OhmsLawResistance(GetNum(txtVoltage), _
            GetNum(txtCurrent), GetNum(txtPower)))
    End If
    
    ' Calculate power.
    If txtPower.Text = vbNullString Then
        txtPower.Text = NumberToMagString(OhmsLawPower(GetNum(txtVoltage), _
            GetNum(txtCurrent), 0))
    End If
End Sub

' Automatically clears checked fields.
Public Sub AutoClearFields()
    If chkVoltage.Value = vbChecked Then
        txtVoltage.Text = ""
    End If
    
    If chkCurrent.Value = vbChecked Then
        txtCurrent.Text = ""
    End If
    
    If chkResistance.Value = vbChecked Then
        txtResistance.Text = ""
    End If
    
    If chkPower.Value = vbChecked Then
        txtPower.Text = ""
    End If
End Sub

' Resets the form.
Public Sub ResetFields()
    txtVoltage.Text = ""
    txtCurrent.Text = ""
    txtResistance.Text = ""
    txtPower.Text = ""
End Sub

' Counts the number of empty fields.
Public Function EmptyFields() As Integer
    Dim intCount As Integer
    
    ' Count how many empty fields we have.
    If txtVoltage.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    If txtCurrent.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    If txtResistance.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    If txtPower.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    EmptyFields = intCount
End Function

' Gets the number from a TextBox.
Public Function GetNum(txtBox As TextBox) As Double
    Dim dblValue As Double
    On Error Resume Next
    
    dblValue = ParseNumber(txtBox.Text)
    GetNum = dblValue
End Function

' Calculate button clicked.
Private Sub cmdCalculate_Click()
    Calculate
End Sub

' Reset button clicked.
Private Sub cmdReset_Click()
    ResetFields
End Sub

' Form just loaded.
Private Sub Form_Load()
    ResetFields
End Sub

' Help button clicked.
Private Sub picHelp_Click()
    ShowHelp
End Sub

' Current text field key pressed.
Private Sub txtCurrent_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        Calculate
    End If
End Sub

' Power text field key pressed.
Private Sub txtPower_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        Calculate
    End If
End Sub

' Resistance text field key pressed.
Private Sub txtResistance_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        Calculate
    End If
End Sub

' Voltage text field key pressed.
Private Sub txtVoltage_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        Calculate
    End If
End Sub
