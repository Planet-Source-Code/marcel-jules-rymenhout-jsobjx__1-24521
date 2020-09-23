VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XTMouseEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API-Declarations
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'Control for the Extended Mouse Events
Private mControl As Control

'Shows if the Mouse has entered the Control
Private mboolEntered As Boolean

'Shows if the Mouse has left the Control
Private mboolLeft As Boolean

'Mouse enter a Control
Public Event MouseEnter()
'Mouse leave a Control
Public Event MouseLeave()

'Set or get the Control
Public Property Get Control() As Control
    Set Control = mControl
End Property

Public Property Set Control(ByVal ctlNewValue As Control)
    Set mControl = ctlNewValue
End Property

Public Property Get Captured() As Boolean
    If Not mControl Is Nothing Then
        Captured = (GetCapture = mControl.hwnd)
    End If
End Property

Public Property Let Captured(ByVal boolNewValue As Boolean)
    If Not mControl Is Nothing Then
        If boolNewValue Then
            SetCapture mControl.hwnd
        Else
            ReleaseCapture
        End If
    End If
End Property

'Shows if the Mouse has entered the Control
Public Property Get Entered() As Boolean
    Entered = mboolEntered
End Property

'Show if the Mouse has left the Control
Public Property Get Left() As Boolean
    Left = mboolLeft
End Property

'Shows if the Mouse is over the Control
Public Function IsInside() As Boolean

    Dim pointCurPos As POINTAPI
    Dim rectCtl As RECT

    If Not mControl Is Nothing Then
        GetCursorPos pointCurPos
        GetWindowRect mControl.hwnd, rectCtl
        With pointCurPos
            'Proof, if the Mouse is over the Control
            IsInside = (.X >= rectCtl.Left) And (.X <= rectCtl.Right) _
                And (.Y >= rectCtl.Top) And (.Y <= rectCtl.Bottom)
        End With
    End If

End Function

'Call this Sub in the MouseMove-Event of the Control
Public Sub OnMouseMove()
    If IsInside() Then
        Captured = True
        If Not mboolEntered Then
        'Mouse enter the Control
        mboolEntered = True
        'Mouse is over the Control
        mboolLeft = False
        'Fire the MouseEnter-Event
        RaiseEvent MouseEnter
        End If
    Else
        If Not mboolLeft Then
        Captured = False
        'Mouse has not entered the Control
        mboolEntered = False
        'Mouse has left the Control
        mboolLeft = True
        'Fire the MouseLeave-Event
        RaiseEvent MouseLeave
        End If
    End If
End Sub
