Attribute VB_Name = "Functions"
Sub Arc_Radius_3_Points()
Dim a() As Double, b() As Double, c() As Double
Dim area As Double, radius As Double, dist_a As Double, dist_b As Double, dist_c As Double

ReDim a(1)
ReDim b(1)
ReDim c(1)

'R = 2,520
a(0) = 2.8
a(1) = -7.911
b(0) = 3.555
b(1) = -7.78
c(0) = 4.288
c(1) = -7.389

area = Abs((a(0) * (b(1) - c(1)) + b(0) * (c(1) - a(1)) + c(0) * (a(1) - b(1))) / 2)
Debug.Print "area = " & Format(area, "0.000")

dist_a = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
dist_b = Sqr((c(0) - b(0)) ^ 2 + (c(1) - b(1)) ^ 2)
dist_c = Sqr((c(0) - a(0)) ^ 2 + (c(1) - a(1)) ^ 2)

Debug.Print "Distance a = " & Format(dist_a, "0.000")
Debug.Print "Distance b = " & Format(dist_b, "0.000")
Debug.Print "Distance c = " & Format(dist_c, "0.000")


radius = (dist_a * dist_b * dist_c) / (4 * area)
Debug.Print "Radius = " & Format(radius, "0.000")
End Sub

