VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2400
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "  Grid"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   150
      TabIndex        =   9
      Top             =   5115
      Width           =   870
   End
   Begin VB.CheckBox chkHairs 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "+ Hairs"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   150
      TabIndex        =   8
      Top             =   5340
      Width           =   870
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000080FF&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1275
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5175
      Width           =   840
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0E0FF&
      Height          =   2790
      Left            =   150
      TabIndex        =   6
      Top             =   1800
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Number of ropes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   330
      TabIndex        =   3
      Top             =   735
      Width           =   1695
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Two ropes"
         Height          =   315
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   570
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "One rope"
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   225
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   405
      TabIndex        =   2
      Text            =   "KnotName"
      Top             =   330
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddKnot 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Add knot to list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   375
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4635
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter knot name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   390
      TabIndex        =   1
      Top             =   75
      Width           =   1545
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form2.frm   Make a knot

Option Base 1
DefLng A-W
DefSng X-Z

Dim RopesToDo

Private Sub Form_Load()

MakingKnot = True

Caption = Space$(15) & "Make a knot"
Left = Form1.Left + Form1.Width
Top = Form1.Top
Height = Form1.Height

List1.AddItem " Click on display:"
List1.AddItem " "
List1.AddItem " LButton          - point o"
List1.AddItem " Ctrl + LButton - point u"
List1.AddItem " "
List1.AddItem " RButton - swap o and u"
List1.AddItem " Ctrl + RButton - delete"
List1.AddItem "         to end of rope"
List1.AddItem " "
List1.AddItem " Points o overdraw"
List1.AddItem " Points u underdraw"
List1.AddItem " "
List1.AddItem " Ropes will be drawn"
List1.AddItem " in order of entry."

Option1(0).Value = True
RopesToDo = 1
Rope = 1
NPts(1) = 0

Form1.LabRope.Caption = "ROPE" & Str$(Rope)
End Sub

Private Sub cmdAddKnot_Click()
' Add knot to list

If NPts(1) = 0 Then
   Beep
   GoTo BackToForm1
End If

If RopesToDo = 2 Then
   If Rope = 1 Then
      Rope = Rope + 1
      Form1.LabRope.Caption = "ROPE" & Str$(Rope)
      NPts(2) = 0
      cmdAddKnot.Caption = "Add knot to list"
      Exit Sub
   End If
Else
   Rope = 1
End If

If RopesToDo = 2 And NPts(2) = 0 Then
   NPts(1) = 0
   Beep
   GoTo BackToForm1
End If

NumOfKnots = NumOfKnots + 1
KName$(NumOfKnots) = Text1.Text
Form1.List1.AddItem KName$(NumOfKnots)

NR(NumOfKnots) = Rope
For rp = 1 To Rope
   NP(NumOfKnots, rp) = NPts(rp)
   For i = 1 To NPts(rp)
      xr(NumOfKnots, rp, i) = xp(rp, i)
      yr(NumOfKnots, rp, i) = yp(rp, i)
      zr(NumOfKnots, rp, i) = zp(rp, i)
   Next i
Next rp

'==============
BackToForm1:

MakingKnot = False
Hairs = False
Grid = False

Form1.PIC.Cls
Form1.LabRope.Caption = ""

' Reset PIC cursor
SetClassLong Form1.PIC.hwnd, GCL_HCURSOR, PICCurHandle
DoEvents    ' seems necessary?

Unload Form2

End Sub

Private Sub cmdCancel_Click()
' Cancel

MakingKnot = False
Hairs = False
Grid = False

Form1.PIC.Cls
Form1.LabRope.Caption = ""

' Reset PIC cursor
SetClassLong Form1.PIC.hwnd, GCL_HCURSOR, PICCurHandle
DoEvents

Unload Form2

End Sub

Private Sub Option1_Click(Index As Integer)

If Option1(0).Value Then RopesToDo = 1 Else RopesToDo = 2
DoEvents

If RopesToDo = 1 Then
   cmdAddKnot.Caption = "Add knot to list"
Else
   cmdAddKnot.Caption = "Add rope to list"
End If

End Sub

Private Sub chkHairs_Click()

Hairs = Not Hairs

End Sub

Private Sub chkGrid_Click()

Grid = Not Grid

End Sub

