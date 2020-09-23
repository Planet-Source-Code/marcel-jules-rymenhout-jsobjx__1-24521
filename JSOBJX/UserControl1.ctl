VERSION 5.00
Begin VB.UserControl JSOBJX 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   ScaleHeight     =   2265
   ScaleWidth      =   2805
   Begin VB.PictureBox PICCTL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "JSOBJX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )

Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest

Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Dim mypos As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim M_SPEED As Long
Dim M_FRAMES As Integer
Dim M_PATH As String
Dim M_CLICK As String
Dim M_SPRITE As String
Dim M_SKIN As String
Dim M_PREVIEW As String
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

 Private WithEvents cmdTestXTEvents As XTMouseEvents
Attribute cmdTestXTEvents.VB_VarHelpID = -1

Dim myx As clsBitmap
Dim myy As clsBitmap
Dim MYP As clsBitmap



Private Sub doframe()
On Error Resume Next
BitBlt PICCTL.hDC, 0, 0, (myy.Width / M_FRAMES), (myy.Height) / 2, myy.hDC, (myy.Width / M_FRAMES) * mypos, (myy.Height / 2), SRCAND
PICCTL.Refresh
BitBlt PICCTL.hDC, _
        0, 0, _
        (myy.Width / M_FRAMES), _
        myy.Height / 2, _
        myy.hDC, _
        (myy.Width / M_FRAMES) * mypos, 0, SRCPAINT
PICCTL.Refresh
End Sub


Private Sub doPREVIEW()
On Error Resume Next
BitBlt PICCTL.hDC, 0, 0, (MYP.Width), (MYP.Height) / 2, MYP.hDC, 0, (MYP.Height / 2), SRCAND
PICCTL.Refresh
BitBlt PICCTL.hDC, _
        0, 0, _
        (MYP.Width), _
        MYP.Height / 2, _
        MYP.hDC, _
        0, 0, SRCPAINT
PICCTL.Refresh
End Sub



Private Sub doclick()
On Error Resume Next
Set myx = New clsBitmap
myx.LoadFile M_PATH & M_CLICK

'------------------------
'Do Bitblt function using picMask.hDC
'Note- VbSrcAnd combines the pixels using the And operator
'------------------------

BitBlt PICCTL.hDC, 0, 0, (myx.Width / 3), (myx.Height) / 2, myx.hDC, (myx.Width / 3) * mypos, (myx.Height / 2), SRCAND
        
'------------------------
'Refresh the destination hDC
'------------------------
PICCTL.Refresh
'------------------------
'Do Bitblt function using picSprite.hDC
'Note- VbSrcPaint combines the pixels using the XOR operator
'------------------------
BitBlt PICCTL.hDC, _
        0, 0, _
        (myx.Width / 3), _
        myx.Height / 2, _
        myx.hDC, _
        (myx.Width / 3) * mypos, 0, SRCPAINT
        
'------------------------
'Refresh the destination hDC
'------------------------
PICCTL.Refresh
Set myx = Nothing
End Sub
Public Property Get SKIN() As String
SKIN = M_SKIN
End Property

Public Property Let SKIN(NEWSKIN As String)
If Right(NEWSKIN, 1) <> "\" Then
    NEWSKIN = NEWSKIN & "\"
End If



M_SKIN = NEWSKIN

Dim xpath As String
xpath = Dir(NEWSKIN)
If xpath <> "" Then

    LOADDAT
    Set MYP = New clsBitmap
    MYP.LoadFile M_PATH & M_PREVIEW
   
    doPREVIEW
    PropertyChanged
End If
UserControl_Resize

End Property

Private Sub LOADDAT()
Dim xpath As String
xpath = Dir(M_SKIN)
If xpath <> "" Then
On Error Resume Next
Dim fs, a, retstring
Dim MYSTRING
Dim textline As String


Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.opentextfile(M_SKIN & "Jsx.obj")
Do While a.atendofstream <> True
 textline = a.readline
 MYSTRING = Split(textline, "=")
  '  If MYSTRING(0) = "PATH" Then
  '     M_PATH = MYSTRING(1)
  '  End If
    If MYSTRING(0) = "SPRITE" Then
       M_SPRITE = MYSTRING(1)
    End If
    If MYSTRING(0) = "CLICK" Then
       M_CLICK = MYSTRING(1)
    End If
    If MYSTRING(0) = "SPEED" Then
       M_SPEED = CLng(MYSTRING(1))
    End If
    If MYSTRING(0) = "FRAMES" Then
       M_FRAMES = CInt(MYSTRING(1))
    End If
     If MYSTRING(0) = "PREVIEW" Then
       M_PREVIEW = MYSTRING(1)
    End If

Loop
M_PATH = M_SKIN
a.Close
End If




End Sub





Private Sub cmdTestXTEvents_MouseEnter()
Set myy = New clsBitmap
myy.LoadFile M_PATH & M_SPRITE
For i = 0 To M_FRAMES - 1
    mypos = i
    PICCTL.Cls
    
    doframe
    Sleep M_SPEED
Next i
End Sub


Private Sub cmdTestXTEvents_MouseLeave()
For i = 0 To M_FRAMES - 1
    mypos = (M_FRAMES - 1) - i
    PICCTL.Cls
    
    doframe
    Sleep M_SPEED
Next i
Set myy = Nothing
PICCTL.Cls
doPREVIEW


End Sub


Private Sub PICCTL_Click()
RaiseEvent Click
End Sub

Private Sub PICCTL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

For i = 0 To 2
    mypos = i
    PICCTL.Cls
    doclick
    Sleep M_SPEED
Next i
End Sub


Private Sub PICCTL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdTestXTEvents.OnMouseMove
End Sub


Private Sub PICCTL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


For i = 0 To 2
    mypos = (3 - 1) - i
    PICCTL.Cls
    doclick
    Sleep M_SPEED
Next i
End Sub




Private Sub UserControl_AmbientChanged(PropertyName As String)
PICCTL.BackColor = Ambient.BackColor
End Sub

Private Sub UserControl_Initialize()
Set cmdTestXTEvents = New XTMouseEvents
Set cmdTestXTEvents.Control = PICCTL




End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
M_SKIN = PropBag.ReadProperty("SKINz", "")
M_FRAMES = CInt(PropBag.ReadProperty("FRAMEz", 0))
M_SPEED = CLng(PropBag.ReadProperty("SPEEDz", 0))
M_SPRITE = PropBag.ReadProperty("SPRITEz", "")
M_CLICK = PropBag.ReadProperty("CLICKz", "")
M_PATH = PropBag.ReadProperty("PATHz", "")
M_PREVIEW = PropBag.ReadProperty("PREVIEWz", "")


  Set MYP = New clsBitmap
   MYP.LoadFile M_PATH & M_PREVIEW
   doPREVIEW
  
     
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

UserControl.Width = (MYP.Width) * Screen.TwipsPerPixelX
UserControl.Height = (MYP.Height / 2) * Screen.TwipsPerPixelY
PICCTL.BackColor = Ambient.BackColor
PICCTL.Move 0, 0, UserControl.Width, UserControl.Height
doPREVIEW
End Sub


Private Sub UserControl_Terminate()
Set myy = Nothing
Set myx = Nothing
Set MYP = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     PropBag.WriteProperty "SKINz", M_SKIN, ""
     PropBag.WriteProperty "FRAMEz", M_FRAMES, ""
     PropBag.WriteProperty "SPEEDz", M_SPEED, ""
     PropBag.WriteProperty "SPRITEz", M_SPRITE, ""
     PropBag.WriteProperty "CLICKz", M_CLICK, ""
     PropBag.WriteProperty "PATHz", M_PATH, ""
     PropBag.WriteProperty "PREVIEWz", M_PREVIEW, ""
     
   
End Sub










