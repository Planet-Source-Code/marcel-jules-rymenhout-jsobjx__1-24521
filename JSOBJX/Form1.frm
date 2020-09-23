VERSION 5.00
Object = "{EA600A57-4EAC-11D5-BA87-0060085F3BD0}#4.0#0"; "Vbobjx.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00D67563&
   Caption         =   "JSOBJX Sample"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin JSOBJ.JSOBJX JSOBJX1 
      Height          =   0
      Left            =   3270
      TabIndex        =   0
      Top             =   600
      Width           =   0
      _ExtentX        =   0
      _ExtentY        =   0
      FRAMEz          =   0
      SPEEDz          =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Try changing forms background color"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
With List1
.AddItem "Calculater"
.AddItem "Access"
.AddItem "Word"
.AddItem "Excel"
.AddItem "Outlook"
.AddItem "Desktop"
.AddItem "Notepad"
.AddItem "folder"
.AddItem "hd"
.AddItem "my computer"
.AddItem "cdrom"
.AddItem "createcd"
End With
End Sub


Private Sub List1_Click()
Me.JSOBJX1.SKIN = App.Path & "\skins\" & List1.Text & "\"

End Sub


