VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TOPSIS()
x = InputBox("Alternatif Sayýsý Giriniz")
y = InputBox("Kriter Sayýsý Giriniz")
Dim etop(100) As Double
Dim renk(100, 100) As Double
Dim w(100) As Double
Dim trr(100, 100) As Double
Dim ttop(100) As Double
Dim normal(100, 100) As Double
Dim ext(100) As Integer
Dim pideal(100) As Double
Dim nideal(100) As Double
Dim pstop(100) As Double
Dim nstop(100) As Double
Dim rcis(100) As Double
Dim AA(100) As Double
Dim BB(100) As Double
Dim yedek(100) As Double
Dim son(100) As Double

Dim r As Long
Dim c As Long




For c = 2 To y + 1
r = 2
AA(c - 1) = Cells(r, c)
BB(c - 1) = Cells(r, c)
For r = 2 To x + 1
If Cells(r, c) > AA(c - 1) Then
AA(c - 1) = Cells(r, c)
End If
If Cells(r, c) < BB(c - 1) Then
BB(c - 1) = Cells(r, c)
End If
Next r
Next c


For c = 2 To y + 1
For r = 2 To x + 1
If Cells(r, c) < 0 Then
Cells(r, c) = (AA(c - 1) - BB(c - 1)) * (Cells(r, c) - Int(BB(c - 1))) / ((Sgn(AA(c - 1)) * Int(Abs(AA(c - 1)))) - (Sgn(BB(c - 1)) * Int(Abs(BB(c - 1)))))

End If

Next r
Next c
For c = 2 To y + 1
For r = 2 To x + 1
If Cells(r, c) = 0 Then
Cells(r, c) = 0.000001
End If
Next r
Next c



r = 2
 c = 1
For r = 2 To x + 1

Cells(r, c).Select
Cells(r, c).Value = "A" & CStr(r - 1)

Next r
r = 1
For c = 2 To y + 1

Cells(r, c).Select
Cells(r, c).Value = "C" & CStr(c - 1)
ext(c - 1) = InputBox("Bu kriter için ideal deðer minimum ise 1 maximum ise 2 giriniz.")
Next c
Cells(x + 2, 1).Value = "W"

For c = 2 To y + 1
etop(c - 1) = 0
ttop(c - 1) = 0
For r = 2 To x + 1
s = Cells(r, c)
ttop(c - 1) = ttop(c - 1) + s ^ 2

etop(c - 1) = etop(c - 1) + s
Next r
ttop(c - 1) = ttop(c - 1) ^ (1 / 2)


Next c

For c = 2 To y + 1

For r = 2 To x + 1

renk(r - 1, c - 1) = Cells(r, c) / etop(c - 1)
normal(r, c) = Cells(r, c) / ttop(c - 1)



Next r


Next c

p = 0
For c = 2 To y + 1
t = 0
For r = 2 To x + 1


k = -renk(r - 1, c - 1) * Log(renk(r - 1, c - 1)) / Log(x)
t = t + k
Next r
p = p + 1 - t
w(c - 1) = t



Next c



For c = 2 To y + 1
w(c - 1) = (1 - w(c - 1)) / p

Cells(x + 2, c).Value = w(c - 1)

Next c

Cells(x + 6, 1).Value = "WNDM"
For c = 2 To y + 1
For r = 2 To x + 1

normal(r, c) = normal(r, c) * w(c - 1)
Cells(r + x + 4, c).Value = normal(r, c)


Next r
Next c

Cells(x + 3, 1).Value = "V+"
Cells(x + 4, 1).Value = "V-"
For c = 2 To y + 1
r = 2
mini = normal(r, c)
maxi = normal(r, c)

For r = 2 To x + 1

If normal(r, c) > maxi Then
maxi = normal(r, c)
End If

If normal(r, c) < mini Then
mini = normal(r, c)
End If


Next r
If ext(c - 1) = 1 Then
pideal(c - 1) = mini
nideal(c - 1) = maxi
End If
If ext(c - 1) = 2 Then
pideal(c - 1) = maxi
nideal(c - 1) = mini
End If
Cells(x + 3, c).Value = pideal(c - 1)
Cells(x + 4, c).Value = nideal(c - 1)
Next c



For r = 2 To x + 1
pstop(r - 1) = 0
nstop(r - 1) = 0
For c = 2 To y + 1
pstop(r - 1) = pstop(r - 1) + (normal(r, c) - pideal(c - 1)) ^ 2
nstop(r - 1) = nstop(r - 1) + (normal(r, c) - nideal(c - 1)) ^ 2
Next c
pstop(r - 1) = pstop(r - 1) ^ (1 / 2)
nstop(r - 1) = nstop(r - 1) ^ (1 / 2)
Cells(r, y + 3).Value = pstop(r - 1)
Cells(r, y + 4).Value = nstop(r - 1)
Next r
For r = 2 To x + 1
rcis(r - 1) = nstop(r - 1) / (nstop(r - 1) + pstop(r - 1))

Next r
Cells(1, y + 3).Value = "S+"

Cells(1, y + 4).Value = "S-"

For r = 2 To x + 1
Cells(r, y + 7) = rcis(r - 1)
Next r
Cells(1, y + 7).Value = "C+"




End Sub


