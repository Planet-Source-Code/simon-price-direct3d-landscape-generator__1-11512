Attribute VB_Name = "MOD_GEO"
' how smooth the engine makes it circles
Const CIRCLE_SMOOTHNESS = 20

' sets a vertex without a normal
Sub SetVertexWON(v As D3DVERTEX, x As Single, y As Single, Z As Single, tu As Single, tv As Single)
With v
   .x = x
   .y = y
   .Z = Z
   .tu = tu
   .tv = tv
End With
End Sub

' makes a triangle
Public Sub MakeTri(ThePrim As tPrim, Pos As D3DVECTOR, w As Single, h As Single)
With ThePrim
    .PrimType = D3DPT_TRIANGLESTRIP
    ReDim .v(0 To 2)
    SetVertexWON .v(0), Pos.x - w / 2, Pos.y - h / 2, Pos.Z, 0, 0
    SetVertexWON .v(1), Pos.x, Pos.y + h / 2, Pos.Z, 0.5, 1
    SetVertexWON .v(2), Pos.x + w / 2, Pos.y - h / 2, Pos.Z, 1, 0
End With
End Sub

' makes a right-angle triangle
Public Sub MakeRightTri(ThePrim As tPrim, Pos As D3DVECTOR, w As Single, h As Single)
With ThePrim
    .PrimType = D3DPT_TRIANGLESTRIP
    ReDim .v(0 To 2)
    SetVertexWON .v(0), Pos.x, Pos.y, Pos.Z, 0, 0
    SetVertexWON .v(1), Pos.x, Pos.y + h, Pos.Z, 0, 1
    SetVertexWON .v(2), Pos.x + w, Pos.y, Pos.Z, 1, 0
End With
End Sub

' makes a rectangle
Sub MakeRectangle(ThePrim As tPrim, Pos As D3DVECTOR, Width As Single, Height As Single)
w = Width / 2
h = Height / 2
With ThePrim
    .PrimType = D3DPT_TRIANGLESTRIP
    ReDim .v(0 To 3)
    SetVertexWON .v(0), Pos.x - w, Pos.y + h, Pos.Z, 0, 1
    SetVertexWON .v(1), Pos.x + w, Pos.y + h, Pos.Z, 1, 1
    SetVertexWON .v(2), Pos.x - w, Pos.y - h, Pos.Z, 0, 0
    SetVertexWON .v(3), Pos.x + w, Pos.y - h, Pos.Z, 1, 0
End With
End Sub

' makes a circle
Sub MakeCircle(ThePrim As tPrim, Pos As D3DVECTOR, w As Single, h As Single)
MakeRegularPoly ThePrim, Pos, w, h, CIRCLE_SMOOTHNESS
End Sub

' makes a polygon of any no. of sides
Public Sub MakeRegularPoly(ThePrim As tPrim, Pos As D3DVECTOR, w As Single, h As Single, NumSides As Byte)
Dim i As Byte
Dim DiffAngle As Double
Dim CurAngle As Double
Dim DiffTu As Double
Dim CurTu As Single
w = w / 2
h = h / 2
With ThePrim
    .PrimType = D3DPT_TRIANGLEFAN
    ReDim .v(0 To NumSides + 1)
    SetVertexWON .v(0), Pos.x, Pos.y, Pos.Z, 0, 0
    DiffAngle = 2 * PI / NumSides
    DiffTu = 1 / NumSides
    For i = 1 To NumSides
         SetVertexWON .v(i), Pos.x + Sin(CurAngle) * w, Pos.y + Cos(CurAngle) * h, Pos.Z, CurTu, 1
         CurAngle = CurAngle + DiffAngle
         CurTu = CurTu + DiffTu
    Next
    CurAngle = 2 * PI
    CurTu = 1
    SetVertexWON .v(i), Pos.x + Sin(CurAngle) * w, Pos.y + Cos(CurAngle) * h, Pos.Z, CurTu, 1
End With
End Sub

' rotates the shape so it ends up facing the right way
Sub RotatePrim(ThePrim As tPrim, N As D3DVECTOR)
Dim C As D3DVECTOR
Dim C2 As D3DVECTOR
Dim vec As D3DVECTOR
Dim i As Byte
With ThePrim
For i = 0 To UBound(.v)
    C.x = C.x + .v(i).x
    C.y = C.y + .v(i).y
    C.Z = C.Z + .v(i).Z
Next
C.x = C.x / (UBound(.v) + 1)
C.y = C.y / (UBound(.v) + 1)
C.Z = C.Z / (UBound(.v) + 1)
For i = 0 To UBound(.v)
    MOD_MATH.CopyVertex2Vec .v(i), vec
    vec = MOD_MATH.RotateYVectorAroundVector(vec, N.x, C)
    vec = MOD_MATH.RotateXVectorAroundVector(vec, N.y, C)
    vec = MOD_MATH.RotateZVectorAroundVector(vec, N.Z, C)
    MOD_MATH.CopyVec2Vertex vec, .v(i)
Next
End With
End Sub

Sub CalculateNormals(ThePrim As tPrim)
On Error Resume Next
Dim i As Byte
Dim i2 As Byte
Dim vec As D3DVECTOR
Dim vec2 As D3DVECTOR
Dim vec3 As D3DVECTOR
Dim vec4 As D3DVECTOR
With ThePrim
    Select Case .PrimType
        Case D3DPT_TRIANGLELIST
            For i = 0 To UBound(.v) Step 3
                vec = MOD_MATH.NormalOfPlane(.v(i), .v(i + 1), .v(i + 2))
                MOD_MATH.CopyVec2VertexNormal vec, .v(i)
                MOD_MATH.CopyVec2VertexNormal vec, .v(i + 1)
                MOD_MATH.CopyVec2VertexNormal vec, .v(i + 2)
            Next
        Case D3DPT_TRIANGLESTRIP
            For i = 0 To UBound(.v) Step 2
                vec4 = MOD_MATH.NormalOfPlane(.v(i), .v(i + 1), .v(i + 2))
                If i > 0 Then
                   vec2 = vec4
                   MOD_MATH.CopyVertexNormal2Vec .v(i), vec3
                   vec = MOD_MATH.AverageOf2Vectors(vec2, vec3)
                Else
                   vec = vec4
                End If
                MOD_MATH.CopyVec2VertexNormal vec, .v(i)
                If i > 2 Then
                   vec2 = vec4
                   MOD_MATH.CopyVertexNormal2Vec .v(i + 1), vec3
                   vec = MOD_MATH.AverageOf2Vectors(vec2, vec3)
                Else
                   vec = vec4
                End If
                MOD_MATH.CopyVec2VertexNormal vec, .v(i + 1)
                If i > 2 Then
                   vec2 = vec4
                   MOD_MATH.CopyVertexNormal2Vec .v(i + 1), vec3
                   vec = MOD_MATH.AverageOf2Vectors(vec2, vec3)
                Else
                   vec = vec4
                End If
                MOD_MATH.CopyVec2VertexNormal vec, .v(i + 2)
            Next
        Case D3DPT_TRIANGLEFAN
            For i = 1 To UBound(.v)
                vec4 = MOD_MATH.NormalOfPlane(.v(0), .v(i), .v(i + 1))
                If i > 1 Then
                   vec2 = vec4
                   MOD_MATH.CopyVertexNormal2Vec .v(0), vec3
                   vec = MOD_MATH.AverageOf2Vectors(vec2, vec3)
                Else
                   vec = vec4
                End If
                MOD_MATH.CopyVec2VertexNormal vec, .v(0)
                If i > 1 Then
                   vec2 = vec4
                   MOD_MATH.CopyVertexNormal2Vec .v(i), vec3
                   vec = MOD_MATH.AverageOf2Vectors(vec2, vec3)
                Else
                   vec = vec4
                End If
                MOD_MATH.CopyVec2VertexNormal vec, .v(i)
                MOD_MATH.CopyVec2VertexNormal vec, .v(i + 1)
            Next
    End Select
End With
End Sub

' makes a cube from rectangles
Public Sub MakeCube(ThePrims() As tPrim, Pos As D3DVECTOR, Size As D3DVECTOR)
ReDim ThePrims(0 To 5)
MakeRectangle ThePrims(0), MakeVector(Pos.x, Pos.y, Pos.Z - Size.Z / 2), Size.x, Size.y
MakeRectangle ThePrims(1), MakeVector(Pos.x, Pos.y, Pos.Z + Size.Z / 2), Size.x, Size.y
RotatePrim ThePrims(1), MakeVector(Deg2Rad(180), 0, 0)
MakeRectangle ThePrims(2), MakeVector(Pos.x - Size.x / 2, Pos.y, Pos.Z), Size.Z, Size.y
RotatePrim ThePrims(2), MakeVector(Deg2Rad(-90), 0, 0)
MakeRectangle ThePrims(3), MakeVector(Pos.x + Size.x / 2, Pos.y, Pos.Z), Size.Z, Size.y
RotatePrim ThePrims(3), MakeVector(Deg2Rad(90), 0, 0)
MakeRectangle ThePrims(4), MakeVector(Pos.x, Pos.y + Size.y / 2, Pos.Z), Size.x, Size.Z
RotatePrim ThePrims(4), MakeVector(0, Deg2Rad(-90), 0)
MakeRectangle ThePrims(5), MakeVector(Pos.x, Pos.y - Size.y / 2, Pos.Z), Size.x, Size.Z
RotatePrim ThePrims(5), MakeVector(0, Deg2Rad(90), 0)
End Sub

' makes a triangle based pyrimid from triangles
Public Sub MakeTriBasePyrimid(ThePrims() As tPrim, Pos As D3DVECTOR, Size As D3DVECTOR)
Dim vec() As D3DVECTOR
ReDim vec(0 To 3)
vec(0) = MakeVector(Pos.x, Pos.y + Size.y, Pos.Z)
vec(1) = MakeVector(Pos.x, Pos.y, Pos.Z - Size.Z / 2)
vec(2) = MakeVector(Pos.x - Size.x / 2, Pos.y, Pos.Z + Size.Z / 2)
vec(3) = MakeVector(Pos.x + Size.x / 2, Pos.y, Pos.Z + Size.Z / 2)
ReDim ThePrims(0 To 3)
ThePrims(0) = MakeTriangle(vec(3), vec(2), vec(1))
ThePrims(1) = MakeTriangle(vec(2), vec(0), vec(1))
ThePrims(2) = MakeTriangle(vec(1), vec(0), vec(3))
ThePrims(3) = MakeTriangle(vec(3), vec(0), vec(2))
End Sub

' makes a square based pyrimid from triangles + 1 square
Public Sub MakeSquareBasePyrimid(ThePrims() As tPrim, Pos As D3DVECTOR, Size As D3DVECTOR)
Dim vec() As D3DVECTOR
ReDim vec(0 To 4)
vec(0) = MakeVector(Pos.x, Pos.y + Size.y, Pos.Z)
vec(1) = MakeVector(Pos.x - Size.x / 2, Pos.y, Pos.Z - Size.Z / 2)
vec(2) = MakeVector(Pos.x + Size.x / 2, Pos.y, Pos.Z - Size.Z / 2)
vec(3) = MakeVector(Pos.x - Size.x / 2, Pos.y, Pos.Z + Size.Z / 2)
vec(4) = MakeVector(Pos.x + Size.x / 2, Pos.y, Pos.Z + Size.Z / 2)
ReDim ThePrims(0 To 4)
ThePrims(0) = MakeTriangle(vec(0), vec(2), vec(1))
ThePrims(1) = MakeTriangle(vec(0), vec(3), vec(4))
ThePrims(2) = MakeTriangle(vec(0), vec(1), vec(3))
ThePrims(3) = MakeTriangle(vec(0), vec(4), vec(2))
ThePrims(4) = MakeQuadrilateral(vec(1), vec(2), vec(3), vec(4))
End Sub

' makes a custom shape triangle
Function MakeTriangle(v1 As D3DVECTOR, v2 As D3DVECTOR, v3 As D3DVECTOR) As tPrim
With MakeTriangle
    .PrimType = D3DPT_TRIANGLELIST
    ReDim .v(0 To 2)
    CopyVec2Vertex v1, .v(0)
    CopyVec2Vertex v2, .v(1)
    CopyVec2Vertex v3, .v(2)
End With
End Function

' makes a custom shape quadrilateral
Function MakeQuadrilateral(v1 As D3DVECTOR, v2 As D3DVECTOR, v3 As D3DVECTOR, v4 As D3DVECTOR) As tPrim
With MakeQuadrilateral
    .PrimType = D3DPT_TRIANGLESTRIP
    ReDim .v(0 To 3)
    CopyVec2Vertex v1, .v(0)
    CopyVec2Vertex v2, .v(1)
    CopyVec2Vertex v3, .v(2)
    CopyVec2Vertex v4, .v(3)
End With
End Function

' makes a cylinder from the primitives passed to it
Sub MakeCylinder(ThePrims() As tPrim, Pos As D3DVECTOR, Size As D3DVECTOR)
On Error Resume Next
Dim i As Byte
Dim i2 As Byte
Dim Size2 As D3DVECTOR
Size2 = Size
ReDim ThePrims(0 To 2)
MakeCircle ThePrims(0), MakeVector(Pos.x, Pos.y + Size.y / 2, Pos.Z), Size.x, Size.Z
RotatePrim ThePrims(0), MakeVector(0, Deg2Rad(-90), 0)
ThePrims(1) = ThePrims(0)
For i = 0 To UBound(ThePrims(1).v)
    ThePrims(1).v(i).y = ThePrims(1).v(i).y - Size2.y
Next
ReDim ThePrims(2).v(0 To UBound(ThePrims(0).v) * 2)
With ThePrims(2)
.PrimType = D3DPT_TRIANGLESTRIP
For i = 1 To UBound(ThePrims(0).v)
    .v((i - 1) * 2) = ThePrims(0).v(i)
    .v((i - 1) * 2 + 1) = ThePrims(1).v(i)
Next
.v(UBound(.v)) = ThePrims(0).v(1)
End With
Dim tempPrim As tPrim
tempPrim = ThePrims(1)
With ThePrims(1)
i2 = UBound(.v)
For i = 1 To UBound(.v)
   .v(i) = tempPrim.v(i2)
   i2 = i2 - 1
Next
End With
End Sub

' makes a sphere from the primitives passed to it
Sub MakeSphere(ThePrims() As tPrim, Pos As D3DVECTOR, Size As D3DVECTOR)
'On Error Resume Next
Dim i As Byte
Dim i2 As Byte
Dim ti As Integer
Dim ti2 As Integer
Dim NumSides As Byte
Dim Angle As Single
Dim stepAngle As Single
Dim r As Single
Dim N As D3DVECTOR
Dim tempPrim() As tPrim
NumSides = CIRCLE_SMOOTHNESS - 1
ReDim tempPrim(0 To NumSides)
stepAngle = 2 * PI / NumSides
For i = 0 To UBound(tempPrim)
    r = Size.x
    MakeCircle tempPrim(i), Pos, r, r
    N.x = Angle
    RotatePrim tempPrim(i), N
    Angle = Angle + stepAngle
Next
ReDim ThePrims(0 To NumSides)
For i = 0 To NumSides
    With ThePrims(i)
    If i = NumSides Then
        ti = -1
    Else
        ti = i
    End If
    ReDim .v(0 To NumSides + 1)
    For i2 = 0 To NumSides \ 2
        .PrimType = D3DPT_TRIANGLESTRIP
        .v(i2 * 2) = tempPrim(i).v(i2 + 1)
        .v(i2 * 2 + 1) = tempPrim(ti + 1).v(i2 + 1)
    Next
    .v(NumSides + 1) = tempPrim(0).v(NumSides \ 2 + 2)
    End With
Next
End Sub

' makes a cone from some primitives
Sub MakeCone(ThePrims() As tPrim, Pos As D3DVECTOR, Size As D3DVECTOR)
Dim i As Byte
Dim NumVerts As Byte
ReDim ThePrims(0 To 1)
MakeCircle ThePrims(0), Pos, Size.x, Size.y
RotatePrim ThePrims(0), MakeVector(0, Deg2Rad(90), 0)
NumVerts = UBound(ThePrims(0).v)
ReDim ThePrims(1).v(0 To NumVerts)
For i = 1 To NumVerts
   ThePrims(1).v(i) = ThePrims(0).v(NumVerts - i + 1)
Next
With ThePrims(1)
   .v(0) = ThePrims(0).v(0)
   .v(0).y = Pos.y + Size.y
   .PrimType = D3DPT_TRIANGLEFAN
End With

End Sub
