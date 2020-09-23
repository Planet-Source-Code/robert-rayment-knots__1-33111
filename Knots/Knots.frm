VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Knots  by Robert Rayment"
   ClientHeight    =   5670
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   9105
   ControlBox      =   0   'False
   Icon            =   "KNOTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   607
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5160
      Left            =   1545
      ScaleHeight     =   340
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   9
      Top             =   330
      Width           =   7455
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   265
         X2              =   265
         Y1              =   62
         Y2              =   99
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   245
         X2              =   245
         Y1              =   52
         Y2              =   89
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   225
         X2              =   225
         Y1              =   44
         Y2              =   81
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   201
         X2              =   201
         Y1              =   35
         Y2              =   72
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   183
         X2              =   183
         Y1              =   32
         Y2              =   69
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   169
         X2              =   169
         Y1              =   29
         Y2              =   66
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   116
         X2              =   149
         Y1              =   199
         Y2              =   201
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   115
         X2              =   148
         Y1              =   179
         Y2              =   181
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   119
         X2              =   152
         Y1              =   158
         Y2              =   160
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   115
         X2              =   148
         Y1              =   138
         Y2              =   140
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   116
         X2              =   149
         Y1              =   114
         Y2              =   116
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   114
         X2              =   147
         Y1              =   93
         Y2              =   95
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   108
         X2              =   141
         Y1              =   79
         Y2              =   81
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   111
         X2              =   144
         Y1              =   59
         Y2              =   61
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   153
         X2              =   153
         Y1              =   26
         Y2              =   63
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   109
         X2              =   142
         Y1              =   43
         Y2              =   45
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   4  'Dash-Dot
         X1              =   14
         X2              =   68
         Y1              =   49
         Y2              =   49
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   4  'Dash-Dot
         X1              =   33
         X2              =   33
         Y1              =   21
         Y2              =   82
      End
   End
   Begin VB.CheckBox chkNew 
      BackColor       =   &H000080FF&
      Caption         =   "New"
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   15
      Width           =   495
   End
   Begin VB.CheckBox chkSaveKnots 
      Caption         =   "Save knots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   15
      Width           =   1200
   End
   Begin VB.CheckBox chkLoadKnot 
      Caption         =   "Load knots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   15
      Width           =   1200
   End
   Begin VB.CheckBox chkMakeKnot 
      Caption         =   "Make a knot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   15
      Width           =   1395
   End
   Begin VB.CheckBox chkExit 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   8655
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   15
      Width           =   450
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4545
      Left            =   15
      TabIndex        =   1
      ToolTipText     =   "LB to select. Ctrl+LB to delete."
      Top             =   300
      Width           =   1500
   End
   Begin VB.Label LabXY 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Index           =   1
      Left            =   15
      TabIndex        =   3
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label LabRope 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   4905
      TabIndex        =   2
      Top             =   45
      Width           =   180
   End
   Begin VB.Label LabXY 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      ForeColor       =   &H0080FFFF&
      Height          =   210
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   4875
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Knots.frm

' Knots  by  Robert Rayment   March 2002

Option Base 1
DefLng A-W
DefSng X-Z

Private Sub Form_Load()

Caption = Space$(80) & "Knots  by  Robert Rayment"
Left = 100: Top = 300

With PIC
   .Width = 500
   .Height = 350
End With

PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

MousePointer = vbDefault
DoEvents

' Load colored cursor from file for PIC
FilCurHandle = LoadCursorFromFile(PathSpec$ & "RRHand.cur")

INIT

SetKnotElement ' Rope element shape & color


End Sub

Private Sub INIT()

' CAPTIONS & BACKCOLORS
LabXY(0).Caption = ""
LabXY(1).Caption = ""
LabRope.Caption = ""
chkLoadKnot.BackColor = RGB(136, 92, 40)
chkSaveKnots.BackColor = RGB(136, 92, 40)
chkMakeKnot.BackColor = RGB(136, 92, 40)
chkNew.BackColor = RGB(136, 92, 40)
chkExit.BackColor = RGB(136, 92, 40)

'---------------------------------------------------------------------
' FOR KNOTS' DATA.  Max = 50 knots
ReDim KName$(50)          ' Knot names
ReDim NR(50)              ' Num of ropes in knot
ReDim NP(50, 2)           ' Num points in rope. Max = 200 points
ReDim xr(50, 2, 200), yr(50, 2, 200), zr(50, 2, 200)
                          ' Coarse coords of points, x,y,z (z=0/1)
'---------------------------------------------------------------------
' FOR MAKING A KNOT
Rope = 1
NumOfKnots = 0
ReDim NPts(2)
ReDim xp(2, 200), yp(2, 200), zp(2, 200)
MakingKnot = False
Drawing = False
'---------------------------------------------------------------------
' Cross-hairs
Line1.X1 = 0
Line1.X2 = 0
Line1.Y1 = 0
Line1.Y2 = 0
Line1.Visible = False
Line2.X1 = 0
Line2.X2 = 0
Line2.Y1 = 0
Line2.Y2 = 0
Line2.Visible = False
Hairs = False
'---------------------------------------------------------------------
' GRID
Line3(0).X1 = 50
Line3(0).X2 = 50
Line3(0).Y1 = 0
Line3(0).Y2 = PIC.Height
Line3(0).BorderStyle = vbBSDot
Line3(0).Visible = False
For i = 1 To 8
   Line3(i).X1 = 50 * (i + 1)
   Line3(i).X2 = 50 * (i + 1)
   Line3(i).Y1 = 0
   Line3(i).Y2 = PIC.Height
   Line3(i).BorderStyle = Line3(0).BorderStyle
   Line3(i).Visible = False
Next i

Line4(0).X1 = 0
Line4(0).X2 = PIC.Width
Line4(0).Y1 = 50
Line4(0).Y2 = 50
Line4(0).BorderStyle = vbBSDot
Line4(0).Visible = False
For i = 1 To 6
   Line4(i).X1 = 0
   Line4(i).X2 = PIC.Width
   Line4(i).Y1 = 50 * (i + 1)
   Line4(i).Y2 = 50 * (i + 1)
   Line4(i).BorderStyle = Line4(0).BorderStyle
   Line4(i).Visible = False
Next i
Grid = False
'---------------------------------------------------------------------
' OTHER INITS
LoadOrSave = False
List1.Clear
PIC.Cls
'---------------------------------------------------------------------

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Form As Form

' Remove loaded cursor
DestroyCursor FilCurHandle

' Mke sure all forms cleared
For Each Form In Forms
   Unload Form
   Set Form = Nothing
Next Form
End

End Sub

Private Sub chkLoadKnot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' LOAD KNOTS

chkLoadKnot.Value = Unchecked
If MakingKnot = True Then Exit Sub
If LoadOrSave = True Then Exit Sub
If Drawing = True Then Exit Sub

LoadSave = 0
PIC.Cls
Form3.Show vbModal

If FileSpec$ <> "" Then
   ReadKnots FileSpec$
   List1.Clear
   For i = 1 To NumOfKnots
      List1.AddItem KName$(i)
   Next i
End If

PIC.SetFocus

End Sub

Private Sub chkSaveKnots_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' SAVE KNOTS

chkSaveKnots.Value = Unchecked
If MakingKnot = True Then Exit Sub
If LoadOrSave = True Then Exit Sub
If NumOfKnots = 0 Then
   Beep
   Exit Sub
End If
If Drawing = True Then Exit Sub

LoadSave = 1
PIC.Cls
Form3.Show vbModal

If FileSpec$ <> "" Then
   SaveKnots FileSpec$
End If

PIC.SetFocus

End Sub

Private Sub chkMakeKnot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Make a knot

chkMakeKnot.Value = Unchecked

If LoadOrSave = True Then Exit Sub
If MakingKnot = True Then Exit Sub
If Drawing = True Then Exit Sub

PIC.Cls

' Set colored cursor for PIC
PICCurHandle = SetClassLong(PIC.hwnd, GCL_HCURSOR, FilCurHandle)

Form2.Show

PIC.SetFocus

End Sub

Private Sub chkNew_Click()

chkNew.Value = Unchecked
If MakingKnot = True Then Exit Sub
If LoadOrSave = True Then Exit Sub
If Drawing = True Then Exit Sub

INIT

End Sub

Private Sub chkExit_Click()
Form_Unload 0
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Select knot

If MakingKnot = True Then Exit Sub
If LoadOrSave = True Then Exit Sub
If NumOfKnots = 0 Then Exit Sub
If Drawing = True Then Exit Sub

PIC.Cls
LabRope.Caption = ""
DoEvents

kn = List1.ListIndex + 1  'KnotNumber
If kn < 1 Then Exit Sub

If Shift = 0 And Button = 1 Then ' LB  Select knot
   
   Delayer 50
   LabRope.Caption = KName$(kn)
   DrawRope kn

ElseIf Shift = 2 And Button = 1 Then ' Ctrl+LB  Delete knot entry
'---------------------------------------------------------------------
' KNOTS' DATA.  Max = 50 knots
' ReDim KName$(50)          ' Knot names
' ReDim NR(50)              ' Num of ropes in knot
' ReDim NP(50, 2)           ' Num points in rope. Max = 100 points
' ReDim xr(50, 2, 100), yr(50, 2, 100), zr(50, 2, 100)
                          ' Coarse coords of points, x,y,z (z=0/1)
'---------------------------------------------------------------------
   If kn < NumOfKnots Then

      For i = kn To NumOfKnots - 1
         KName$(i) = KName$(i + 1)
         NR(i) = NR(i + 1)
         NP(i, 1) = NP(i + 1, 1)
         NP(i, 2) = NP(i + 1, 2)
         For j = 1 To 2
            For k = 1 To 200
               xr(i, j, k) = xr(i + 1, j, k)
               yr(i, j, k) = yr(i + 1, j, k)
               zr(i, j, k) = zr(i + 1, j, k)
            Next k
         Next j
      Next i
   
   End If
   
   NumOfKnots = NumOfKnots - 1
   List1.Clear
   If NumOfKnots > 0 Then
      For i = 1 To NumOfKnots
         List1.AddItem KName$(i)
      Next i
   End If
   
End If

End Sub

Private Sub DrawRope(kn)
' kn = KnotNumber
' NR(kn) = RopeNumber =  rp
' p = NP(kn,rp) = Number of starting points in rope

' Original points:-
' xr(kn, rp, p), yr(kn, rp, p), zr(kn, rp, p)

' Method
' 1) Develop a new set of smoothed points (xbb(),ybb()) from these
'    using a Bezier-like interpolation
' 2) Apply formula for transferring zr() 0 & 1 points to zbb()
' 3) Calculate evenly spaced points between xbb(),ybb() Bezier points
' 4) Rotate rope element along direction of drawing
' 5) Draw the rope elements a pixel at a time and if in underdraw mode
'    only plot points on a black background otherwise just plot points.
' 6) Use a time delay between drawing each element so that it can be seen
'    how the knot is made.
' 7) Draw a small circle at the end of the rope


Drawing = True

For rp = 1 To NR(kn)   ' For 1 or 2 ropes

   p = NP(kn, rp)
   
   If p > 1 Then
   
      ReDim xbb(p), ybb(p), zbb(p)
      For N = 1 To p
         xbb(N) = xr(kn, rp, N): ybb(N) = yr(kn, rp, N): zbb(N) = zr(kn, rp, N)
      Next N
      newpts = 2
         
      '-----------------------------------------------------
      If p > 2 Then
         
         '-----------------------------------------------------
         ' Develop Bezier-like points
         xfrac = 0.25
         SUP = 3
         oldpts = p
         For s = 1 To SUP
            ReDim xaa(oldpts), yaa(oldpts), zaa(oldpts)
            For i = 1 To oldpts
               xaa(i) = xbb(i): yaa(i) = ybb(i): zaa(i) = zbb(i)
            Next i
            newpts = 2 * oldpts - 2
            ReDim xbb(newpts), ybb(newpts), zbb(newpts)
            xbb(1) = xaa(1): ybb(1) = yaa(1): zbb(1) = zaa(1)
            For i = 2 To oldpts - 1
               xdx = xaa(i) - xaa(i - 1)
               xbb(2 * i - 2) = xaa(i) - xfrac * xdx
               ydy = yaa(i) - yaa(i - 1)
               ybb(2 * i - 2) = yaa(i) - xfrac * ydy
               xdx = xaa(i + 1) - xaa(i)
               xbb(2 * i - 1) = xaa(i) + xfrac * xdx
               ydy = yaa(i + 1) - yaa(i)
               ybb(2 * i - 1) = yaa(i) + xfrac * ydy
            Next i
            xbb(newpts) = xaa(oldpts): ybb(newpts) = yaa(oldpts)
            oldpts = newpts
         Next s
         '-----------------------------------------------------
      
         ' Fix over-underdraw in new array of points
         For i = 1 To 4
            zbb(i) = zr(kn, rp, 1)
         Next i
         iL = 5
         For N = 2 To p - 1
            iUP = iL + 7
            If iUP > newpts Then iUP = newpts
            For i = iL To iUP
               zbb(i) = zr(kn, rp, N)
            Next i
            iL = iL + 8
         Next N
      
      End If   ' for p > 2
      '-----------------------------------------------------
      
      ' Plot at evenly spaced points between Bezier points  (//)
      zL = 2   ' Distance between plotted points
      For i = 1 To newpts - 1
         
         zy = ybb(i + 1) - ybb(i)
         zx = xbb(i + 1) - xbb(i)
         
         zang = zATan2(zy, zx)
         zcos = Cos(zang): zsin = Sin(zang)
         
         '-----------------------------------------------------
         ' Rotate element along direction
         '  ==>> xbb(i),ybb(i) ==>> xbb(i+1),ybb(i+1)
         For iyd = -4 To 4
         For ixd = -4 To 4
            ixs = ixd * zcos + iyd * zsin
            iys = iyd * zcos - ixd * zsin
            If (ixs >= -4 And ixs <= 4) And (iys >= -4 And iys <= 4) Then
               KED(ixd, iyd) = KES(ixs, iys)
            End If
         Next ixd
         Next iyd
         '-----------------------------------------------------
         
         zdis = Sqr(zy ^ 2 + zx ^ 2)   ' xbb(i),ybb(i) ==>> xbb(i+1),ybb(i+1)
         
         If zdis > zL Then
            nsteps = zdis / zL
            zLsina = zL * zsin: zLcosa = zL * zcos
         Else
            nsteps = 1
            zLsina = 0: zLcosa = 0
         End If
         
         For zj = 1 To nsteps
            
            xx = xbb(i) + (zj - 1) * zLcosa
            yy = ybb(i) + (zj - 1) * zLsina
            
            If zbb(i) = 0 Then      ' OVERDRAW
               For iy = -4 To 4
               For ix = -4 To 4
                  If KED(ix, iy) <> 0 Then PIC.PSet (xx + ix, yy + iy), cul(rp, KED(ix, iy))
               Next ix
               Next iy
               
               Delayer 8
            
            Else     ' UNDERDRAW
               For iy = -4 To 4
               For ix = -4 To 4
                  If PIC.Point(xx + ix, yy + iy) = 0 Then   ' Only plot on a black background
                     If KED(ix, iy) <> 0 Then PIC.PSet (xx + ix, yy + iy), cul(rp, KED(ix, iy))
                  End If
               Next ix
               Next iy
               
               Delayer 24
            
            End If
         
         Next zj
         
         'xx,yy just short of xybb(i+1)
      
      Next i
      '-----------------------------------------------------
      
      If rp = 1 Then
         PIC.Circle (xbb(i), ybb(i)), 2, RGB(200, 200, 0)
      Else
         PIC.Circle (xbb(i), ybb(i)), 2, RGB(0, 200, 0)
      End If
      
   End If  ' for p > 1

Next rp  ' Next rope

Erase xaa, yaa, zaa, xbb, ybb, zbb

Drawing = False

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

LabXY(0).Caption = Str$(X) & Str$(Y)

Line1.Visible = Hairs
Line2.Visible = Hairs

If Hairs And MakingKnot Then
   Line1.X1 = X
   Line1.X2 = X
   Line1.Y1 = 0
   Line1.Y2 = PIC.Height
   Line2.X1 = 0
   Line2.X2 = PIC.Width
   Line2.Y1 = Y
   Line2.Y2 = Y
End If

If Grid And MakingKnot Then
   For i = 0 To 8
      Line3(i).Visible = True
   Next i
   For i = 0 To 6
      Line4(i).Visible = True
   Next i
ElseIf Not Grid Then
   For i = 0 To 8
      Line3(i).Visible = False
   Next i
   For i = 0 To 6
      Line4(i).Visible = False
   Next i
End If

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If MakingKnot = False Then Exit Sub
If LoadOrSave = True Then Exit Sub
If Drawing = True Then Exit Sub

If Button = 1 Then   ' LB
   
   If NPts(Rope) > 199 Then NPts(Rope) = 199
   
   NPts(Rope) = NPts(Rope) + 1
   xp(Rope, NPts(Rope)) = X: yp(Rope, NPts(Rope)) = Y
   zp(Rope, NPts(Rope)) = 0
   
   If Shift = 2 Then zp(Rope, NPts(Rope)) = 1   ' Ctrl LB u
   
   If Rope = 1 Then
      PIC.ForeColor = QBColor(14)   ' Yellow
   Else
      PIC.ForeColor = QBColor(10)   ' Green
   End If
   
   PIC.CurrentX = X - 3
   PIC.CurrentY = Y - 8
   a$ = "o": If zp(Rope, NPts(Rope)) = 1 Then a$ = "u"
   PIC.Print a$;
   
   ' Show rope coords & Number of points
   c$ = Str$(xp(Rope, NPts(Rope))) & _
   Str$(yp(Rope, NPts(Rope))) & _
   Str$(zp(Rope, NPts(Rope))) & _
   Str$(NPts(Rope))
   LabXY(1).Caption = c$
   
   If NPts(Rope) > 180 Then
      MsgBox ("Near points limit, maybe better to cancel and start again!")
   End If
   
   
ElseIf Button = 2 Then  ' RB
   
   For i = 1 To NPts(Rope)
      zdis = Sqr((X - xp(Rope, i)) ^ 2 + (Y - yp(Rope, i)) ^ 2)
      If zdis <= 4 Then Exit For    ' Point found, i th
   Next i
   
   If i <= NPts(Rope) Then ' Point found, i th
   
      If Shift = 0 Then        ' RB swap o <<>> u
         zp(Rope, i) = 1 - zp(Rope, i)
      ElseIf Shift = 2 Then    ' Ctrl RB delete to end of rope
         NPts(Rope) = i - 1
         If Rope = 1 Then LabXY(1).Caption = ""
      End If
         
         PIC.Cls     ' Clear all displayed ropes
         DoEvents
         ' Redraw ropes
         For rp = 1 To Rope
            If rp = 1 Then
               PIC.ForeColor = QBColor(14)   ' Yellow
            Else
               PIC.ForeColor = QBColor(10)   ' Green
            End If
            
            For j = 1 To NPts(rp)
               PIC.CurrentX = xp(rp, j) - 3
               PIC.CurrentY = yp(rp, j) - 8
               a$ = "o": If zp(rp, j) = 1 Then a$ = "u"
               PIC.Print a$;
            Next j
         Next rp
   
   End If

End If
   
End Sub

Private Sub Delayer(TDelay)

T = timeGetTime
Do: DoEvents
Loop Until (timeGetTime() - T >= TDelay Or (timeGetTime() - T) < 0)

End Sub

