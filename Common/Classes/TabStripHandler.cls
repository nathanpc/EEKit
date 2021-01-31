VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabStripHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' TabStripHandler
''' A simple class to help us deal with TabStrips.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_frmParent As Form
Private m_tabStrip As tabStrip
Private m_arrContainers As Object

' Resets the position of all of the controls that we deal with.
Public Sub ResetPositions()
    Dim intIndex As Integer
    
    ' Resize window to fit the TabStrip.
    ParentForm.Width = TStrip.Width + 120
    ParentForm.Height = TStrip.Height + 470
    
    ' Reset the positions of the containers.
    For intIndex = 1 To TStrip.Tabs.Count
        m_arrContainers(intIndex - 1).Visible = False
        m_arrContainers(intIndex - 1).Move TStrip.ClientLeft, _
            TStrip.ClientTop, TStrip.ClientWidth, TStrip.ClientHeight
    Next intIndex
    
    ' Select the first tab.
    TStrip.SelectedItem = TStrip.Tabs(1)
End Sub

' Handles the Click event of the TabStrip.
Public Sub HandleClick()
    Dim intIndex As Integer
    
    ' Show the desired container.
    For intIndex = 0 To (TStrip.Tabs.Count - 1)
        m_arrContainers(intIndex).Visible = _
            ((TStrip.SelectedItem.Index - 1) = intIndex)
    Next intIndex
End Sub

' Parent Form getter.
Public Property Get ParentForm() As Form
    Set ParentForm = m_frmParent
End Property

' Parent Form setter.
Public Property Let ParentForm(frmParent As Form)
    Set m_frmParent = frmParent
End Property

' Tab Strip getter.
Public Property Get TStrip() As tabStrip
    Set TStrip = m_tabStrip
End Property

' Tab Strip setter.
Public Property Let TStrip(tbsTabStrip As tabStrip)
    Set m_tabStrip = tbsTabStrip
End Property

' Containers array getter.
Public Property Get Containers() As Object
    Set Containers = m_arrContainers
End Property

' Containers array setter.
Public Property Let Containers(arrContainers As Object)
    Set m_arrContainers = arrContainers
End Property

