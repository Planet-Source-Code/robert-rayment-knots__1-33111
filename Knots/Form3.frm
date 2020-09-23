VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2610
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3450
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000080FF&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1035
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Accept"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   45
      TabIndex        =   2
      Top             =   390
      Width           =   2490
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   1
      Top             =   60
      Width           =   2505
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   45
      Pattern         =   "*.knt"
      TabIndex        =   0
      Top             =   1395
      Width           =   2490
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form3.frm  Load/Save

Option Base 1
DefLng A-W
DefSng X-Z
' Public LoadOrSave, FileSpec$


Private Sub Form_Load()

LoadOrSave = True

If LoadSave = 0 Then
   Caption = Space$(10) & "Load knots (*.knt)"
Else
   Caption = Space$(10) & "Save knots (*.knt)"
End If

Left = Form1.Left + Form1.Width
Top = Form1.Top
Height = Form1.Height

Text1.Text = ""
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo NoDrive
Dir1.Path = Drive1.Drive
Exit Sub
'==========
NoDrive:
Beep
Exit Sub
Resume
End Sub

Private Sub File1_Click()
Text1.Text = File1.FileName
End Sub

Private Sub cmdAccept_Click()
' Load or Save knots
   
FName$ = Text1.Text

If FName$ <> "" Then
   FPath$ = File1.Path
   If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ & "\"
   FileSpec$ = FPath$ & FName$
   
   If LoadSave = 1 Then FixFileExtension FileSpec$, "knt"

Else
   FileSpec$ = ""
End If

LoadOrSave = False
Form1.LabRope.Caption = ""

Unload Form3

End Sub

Private Sub cmdCancel_Click()

FileSpec$ = ""
LoadOrSave = False
Form1.LabRope.Caption = ""

Unload Form3

End Sub

Private Sub FixFileExtension(FSpec$, Ext$)

E$ = "." + Ext$
pdot = InStr(1, FSpec$, ".")
If pdot = 0 Then
   FSpec$ = FSpec$ + E$
Else
   Ext$ = LCase$(Mid$(FSpec$, pdot))
   If Ext$ <> E$ Then
      FSpec$ = Left$(FSpec$, pdot - 1) + E$
   End If
End If

End Sub

