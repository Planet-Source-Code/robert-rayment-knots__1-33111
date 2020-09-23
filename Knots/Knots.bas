Attribute VB_Name = "Module1"
'Knots.bas

Option Base 1
DefLng A-W
DefSng X-Z

' Knot data
Public Drawing As Boolean
Public NumOfKnots
Public KName$()         ' Knot names
Public NR()             ' Num of ropes in knot
Public NP()             ' Num points in rope
Public xr(), yr(), zr() ' Saved knot's rope coarse coords, x,y,z (z=0/1)
                        
Public cul()      ' Rope colors
Public KES() As Byte, KED() As Byte 'KnotElement Source & Destination arrays

' For making a knot
Public MakingKnot As Boolean
Public NPts(), Rope
Public xp(), yp(), zp()    ' Rope clicked coarse coords, x,y,z (o/u)
Public Hairs As Boolean   ' + hairs
Public Grid As Boolean    ' Grid

Public LoadOrSave As Boolean
Public LoadSave  ' 0 LOAD, 1 SAVE
Public FileSpec$, PathSpec$

' --------------------------------------------------------------

' Windows API - For timing
Public Declare Function timeGetTime& Lib "winmm.dll" ()
' --------------------------------------------------------------

' Windows APIs - For colored cursor

Public Declare Function LoadCursorFromFile Lib "user32" Alias _
"LoadCursorFromFileA" (ByVal lpFileName As String) As Long

Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GCL_HCURSOR = (-12)

Public Declare Function DestroyCursor Lib "user32" _
(ByVal hCursor As Long) As Long

Public FilCurHandle, PICCurHandle

' --------------------------------------------------------------

Public Const pi# = 3.1415926535898

Public Sub ReadKnots(FileName$)

On Error GoTo LoadError

Open FileName$ For Input As #1

Input #1, NumOfKnots
If NumOfKnots = 0 Then Close: Exit Sub

For N = 1 To NumOfKnots

   Line Input #1, KName$(N)   ' Knot name
   Input #1, NR(N)          ' Num of ropes in knot
   For r = 1 To NR(N)
      Input #1, NP(N, r)  ' Num pts in rope
      For p = 1 To NP(N, r)
         Input #1, xr(N, r, p), yr(N, r, p), zr(N, r, p)
      Next p
   Next r

Next N
Close
Exit Sub
'=================
LoadError:
Close
NumOfKnots = 0
Beep
Exit Sub
End Sub

Public Sub SaveKnots(FileName$)

If NumOfKnots = 0 Then
   Beep
   Exit Sub
End If

On Error GoTo SaveError

Open FileName$ For Output As #1

Print #1, NumOfKnots

For i = 1 To NumOfKnots
   Print #1, KName$(i)
   Print #1, NR(i)  ' Number of ropes
   For j = 1 To NR(i)
      Print #1, NP(i, j) ' Number of pts in rope
      For k = 1 To NP(i, j)
         Print #1, xr(i, j, k); ","; yr(i, j, k); ","; zr(i, j, k)
      Next k
   Next j
Next i

Close
Exit Sub
'=================
SaveError:
Close
Beep
Exit Sub

End Sub

Public Sub SetKnotElement()

ReDim cul(2, 4)
Dim a$(-4 To 4)
ReDim KES(-4 To 4, -4 To 4)
ReDim KED(-4 To 4, -4 To 4)

' Knot element colors
' Rope 1
cul(1, 1) = RGB(255, 255, 0)
cul(1, 2) = RGB(200, 200, 0)
cul(1, 3) = RGB(150, 150, 0)
cul(1, 4) = RGB(100, 100, 0)
' 2nd rope
cul(2, 1) = RGB(0, 255, 0)
cul(2, 2) = RGB(0, 200, 0)
cul(2, 3) = RGB(0, 150, 0)
cul(2, 4) = RGB(0, 100, 0)

' Knot element shape, numbers giving colors from cul()
' Both ropes same shape but diffeent colors
'a$(-4) = "000000000"
'a$(-3) = "000000000"
'a$(-2) = "004444400"
'a$(-1) = "003333300"
' a$(0) = "003333300"
' a$(1) = "002222200"
' a$(2) = "001111100"
' a$(3) = "000000000"
' a$(4) = "000000000"

a$(-4) = "000000000"
a$(-3) = "000000000"
a$(-2) = "000444000"
a$(-1) = "003333300"
 a$(0) = "003333300"
 a$(1) = "002222200"
 a$(2) = "000111000"
 a$(3) = "000000000"
 a$(4) = "000000000"

' Transfer shape and color numbers to KES() array
For iy = -4 To 4
For ix = -4 To 4
   KES(ix, iy) = Val(Mid$(a$(iy), ix + 5, 1))
Next ix
Next iy

Erase a$

End Sub

Public Function zATan2(ByVal zy, ByVal zx)
' Find angle Atan from -pi# to +pi#
' Public Const pi# = 3.1415926535898

' USED HERE TO ROTATE ROPE ELEMENT & DIRECTION OF DRAWING

If zx <> 0 Then
   zATan2 = Atn(zy / zx)
   If (zx < 0) Then
     If (zy < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
   End If
Else  ' zx=0
   If Abs(zy) > Abs(zx) Then   'Must be an overflow
      If zy > 0 Then zATan2 = pi# / 2 Else zATan2 = -pi# / 2
   Else
      zATan2 = 0   'Must be an underflow
   End If
End If
End Function

