VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Electronics Engineer Toolkit"
   ClientHeight    =   11775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11775
   ScaleWidth      =   17865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox boxContainers 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   1
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   7695
      TabIndex        =   2
      Top             =   5640
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   1935
         Left            =   2400
         TabIndex        =   4
         Top             =   1440
         Width           =   2895
      End
   End
   Begin VB.PictureBox boxContainers 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   0
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   7695
      TabIndex        =   1
      Top             =   360
      Width           =   7695
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   1455
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
   End
   Begin ComctlLib.TabStrip tbsMain 
      Height          =   5415
      Left            =   10
      TabIndex        =   0
      Top             =   10
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9551
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Basics"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmMain
''' EEKit application launcher window.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_objTabHandler As TabStripHandler

' Form just showed up.
Private Sub Form_Load()
    ' Initialize and setup the Tab Strip handler.
    Set m_objTabHandler = New TabStripHandler
    m_objTabHandler.ParentForm = Me
    m_objTabHandler.TStrip = tbsMain
    m_objTabHandler.Containers = boxContainers
    
    ' Position all the controls appropriately.
    m_objTabHandler.ResetPositions
End Sub

' TabStrip tab selection changed.
Private Sub tbsMain_Click()
    m_objTabHandler.HandleClick
End Sub
