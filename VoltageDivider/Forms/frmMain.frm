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
      Picture         =   "frmMain.frx":0000
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

' Resets all the fields.
Public Sub ResetFields()
    txtVin.Text = ""
    txtVgnd.Text = "0"
    txtR1.Text = ""
    txtVout.Text = ""
    txtR2.Text = ""
    
    ResetOptions optVout
End Sub

' Resets the option buttons.
Public Sub ResetOptions(optCaller As OptionButton)
    optVin.Value = (optCaller.Name = "optVin")
    optVgnd.Value = (optCaller.Name = "optVgnd")
    optR1.Value = (optCaller.Name = "optR1")
    optVout.Value = (optCaller.Name = "optVout")
    optR2.Value = (optCaller.Name = "optR2")
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
