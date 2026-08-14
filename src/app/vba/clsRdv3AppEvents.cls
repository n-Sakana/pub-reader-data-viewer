VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRdv3AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsRdv3AppEvents -- FE-local application events, used for ONE thing: the
' pump watchdog. If the OnTime chain was
' lost while work remains -- a tick error, a canceled schedule -- the next
' local sheet change or activation re-arms it. Duplicate ticks are harmless:
' every channel record dedups by version.
'
' This is NOT a result-notification path. BE results arrive only through the
' file channel, pulled by the pump.
'==============================================================================
Option Explicit

Public WithEvents App As Excel.Application

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Sh.Parent.FullName <> ThisWorkbook.FullName Then Exit Sub
    modRdv3App.Rdv3PumpEnsureArmed
    On Error GoTo 0
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Parent.FullName <> ThisWorkbook.FullName Then Exit Sub
    modRdv3App.Rdv3PumpEnsureArmed
    On Error GoTo 0
End Sub
