VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerlVBExm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-------------------------------------------
'|  Automating VB Code With PERL           |
'|      By: Trenton Miller                 |
'-------------------------------------------
'|  Please don't stick your name on my Code|
'| you will need a PERL interpreter or this|
'| to work.                                |
'-------------------------------------------
'I would do more explaining but i'm lazy

Public NewValue As Single
Private sngInternalSafeValue As Single
Private Sub Class_Initialize()
    NewValue = 1
End Sub
Public Property Get SafeValue() As Single
    SafeValue = sngInternalSafeValue
End Property
Public Property Let SafeValue(ByVal vNewValue As Single)
     If vNewValue > 0 Then
        sngInternalSafeValue = vNewValue
    End If
End Property
Public Function ReturnSafeValue() As Single
    ReturnSafeValue = sngInternalSafeValue
End Function
Public Function Addem(Operand As Integer, Operand2 As Integer) As Integer
    Addem = Operand + Operand2
End Function
Public Sub ShowForm()
    Form1.Show
End Sub
