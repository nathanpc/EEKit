Attribute VB_Name = "modMain"
''' modMain
''' Application main entry module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Import some API calls.
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type
Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean

' Application main entry point.
Private Sub Main()
    Dim iccex As InitCommonControlsExStruct, hMod As Long
    Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all known values
    ' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx

    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_ALL_CLASSES    ' you really should customize this value from the available constants
    End With
    On Error Resume Next ' error? Requires IEv3 or above
    hMod = LoadLibrary("shell32.dll")
    InitCommonControlsEx iccex
    If Err Then
        InitCommonControls ' try Win9x version
        Err.Clear
    End If
    On Error GoTo 0
    frmMain.Show
    
    If hMod Then FreeLibrary hMod

'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.
End Sub

