VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voltage Divider"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHelp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   20
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtVgnd 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Text            =   "0"
      Top             =   2520
      Width           =   735
   End
   Begin VB.OptionButton optVgnd 
      Caption         =   "Option1"
      Height          =   245
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   255
   End
   Begin VB.OptionButton optVin 
      Caption         =   "Option1"
      Height          =   245
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtVin 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox picSchematic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   960
      Picture         =   "frmMain.frx":0A02
      ScaleHeight     =   2580
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton optR2 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   245
         Left            =   1080
         TabIndex        =   8
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton optVout 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   245
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton optR1 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   245
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtR1 
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtR2 
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtVout 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "R1"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "R2"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Vout"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdReset 
      Cancel          =   -1  'True
      Caption         =   "Reset"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblCurrent 
      Caption         =   "Current:"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblPower 
      Caption         =   "Power:"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Vgnd"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Vin"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmVoltageDivider
''' Voltage divider program main form.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Shows a nice help dialog.
Public Sub ShowHelp()
    MsgBox "Enter values into all fields except for the one you want to calculate " & _
        "for and either press Return or click on the Calculate button. " & vbCrLf & _
        vbCrLf & "You always need to leave 1 field blank so that the calculator " & _
        "knows what you want to calculate. You can also use the option buttons to " & _
        "auto-clear a field before calculating.", vbOKOnly + vbInformation, "Help"
End Sub

' Calculates the voltage divider parameters.
Public Sub Calculate()
    Dim dblPower As Double
    Dim dblCurrent As Double
    
    ' Automatically clear some fields.
    AutoClearFields
    
    ' Check empty field requirements.
    If EmptyFields <> 1 Then
        MsgBox "To perform a calculation we need to have exactly 1 empty field.", _
            vbOKOnly + vbExclamation, "Invalid Input"
        Exit Sub
    End If
    
    ' Calculate input voltage.
    If txtVin.Text = vbNullString Then
        txtVin.Text = NumberToMagString(VoltDividerVin(GetTextNum(txtR1), _
            GetTextNum(txtVout), GetTextNum(txtR2), GetTextNum(txtVgnd)))
    End If
    
    ' Calculate R1 value.
    If txtR1.Text = vbNullString Then
        txtR1.Text = NumberToMagString(VoltDividerR1(GetTextNum(txtVin), _
            GetTextNum(txtVout), GetTextNum(txtR2), GetTextNum(txtVgnd)))
    End If
    
    ' Calculate output voltage.
    If txtVout.Text = vbNullString Then
        txtVout.Text = NumberToMagString(VoltDividerVout(GetTextNum(txtVin), _
            GetTextNum(txtR1), GetTextNum(txtR2), GetTextNum(txtVgnd)))
    End If
    
    ' Calculate R2 value.
    If txtR2.Text = vbNullString Then
        txtR2.Text = NumberToMagString(VoltDividerR2(GetTextNum(txtVin), _
            GetTextNum(txtR1), GetTextNum(txtVout), GetTextNum(txtVgnd)))
    End If
    
    ' Calculate ground voltage.
    If txtVgnd.Text = vbNullString Then
        txtVgnd.Text = NumberToMagString(VoltDividerVgnd(GetTextNum(txtVin), _
            GetTextNum(txtR1), GetTextNum(txtVout), GetTextNum(txtR2)))
    End If
    
    ' Calculate power.
    dblPower = OhmsLawPower(ParseNumber(txtVin.Text) - ParseNumber(txtVgnd.Text), _
        0, ParseNumber(txtR1.Text) + ParseNumber(txtR2.Text))
    lblPower.Caption = "Power: " & vbCrLf & NumberToMagString(dblPower)
    
    ' Calculate current.
    dblCurrent = OhmsLawCurrent(ParseNumber(txtVin.Text) - ParseNumber(txtVgnd.Text), _
        ParseNumber(txtR1.Text) + ParseNumber(txtR2.Text), 0)
    lblCurrent.Caption = "Current: " & vbCrLf & NumberToMagString(dblCurrent)
End Sub

' Automatically clears checked fields.
Public Sub AutoClearFields()
    If optVin.Value Then
        txtVin.Text = ""
    End If
    
    If optR1.Value Then
        txtR1.Text = ""
    End If
    
    If optVout.Value Then
        txtVout.Text = ""
    End If
    
    If optR2.Value Then
        txtR2.Text = ""
    End If
    
    If optVgnd.Value Then
        txtVgnd.Text = ""
    End If
End Sub

' Counts the number of empty fields.
Public Function EmptyFields() As Integer
    Dim intCount As Integer
    
    ' Count how many empty fields we have.
    If txtVin.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    If txtR1.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    If txtVout.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    If txtR2.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    If txtVgnd.Text = vbNullString Then
        intCount = intCount + 1
    End If
    
    EmptyFields = intCount
End Function

' Resets all the fields.
Public Sub ResetFields()
    ' Reset text fields.
    txtVin.Text = ""
    txtVgnd.Text = "0"
    txtR1.Text = ""
    txtVout.Text = ""
    txtR2.Text = ""
    
    ' Reset option buttons.
    ResetOptions optVout
    
    ' Reset labels.
    lblPower.Caption = ""
    lblCurrent.Caption = ""
End Sub

' Resets the option buttons.
Public Sub ResetOptions(optCaller As OptionButton)
    optVin.Value = (optCaller.Name = "optVin")
    optVgnd.Value = (optCaller.Name = "optVgnd")
    optR1.Value = (optCaller.Name = "optR1")
    optVout.Value = (optCaller.Name = "optVout")
    optR2.Value = (optCaller.Name = "optR2")
End Sub

' Calculate button clicked.
Private Sub cmdCalculate_Click()
    Calculate
End Sub

' Reset button clicked.
Private Sub cmdReset_Click()
    ResetFields
End Sub

' Form just loaded up.
Private Sub Form_Load()
    ResetFields
End Sub

' R1 option clicked.
Private Sub optR1_Click()
    ResetOptions optR1
End Sub

' R2 option clicked.
Private Sub optR2_Click()
    ResetOptions optR2
End Sub

' Vgnd option clicked.
Private Sub optVgnd_Click()
    ResetOptions optVgnd
End Sub

' Vin option clicked.
Private Sub optVin_Click()
    ResetOptions optVin
End Sub

' Vout option clicked.
Private Sub optVout_Click()
    ResetOptions optVout
End Sub

' Help button clicked.
Private Sub picHelp_Click()
    ShowHelp
End Sub
