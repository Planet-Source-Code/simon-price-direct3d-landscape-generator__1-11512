Attribute VB_Name = "MOD_MATH"
'''''''''''''''''''''''''''''''''''''''''''''''''
'
'         MOD_MATH.BAS - BY SIMON PRICE
'
'           LOADS OF MATHS FUNCTIONS
'
'''''''''''''''''''''''''''''''''''''''''''''''''


' The number Pi
Public Const PI As Single = 3.141592
Public Const PIdiv180 = PI / 180





'**************************************************
'                   RECTANGLE STUFF
'**************************************************

' Returns a RECT
Function MakeRect(x As Integer, y As Integer, Width As Integer, Height As Integer) As RECT
MakeRect.Left = x
MakeRect.Top = y
MakeRect.Right = x + Width
MakeRect.Bottom = y + Height
End Function

' Sets the values of a RECT
Sub SetRect(TheRect As RECT, x As Integer, y As Integer, Width As Integer, Height As Integer)
On Error Resume Next
TheRect.Left = x
TheRect.Top = y
TheRect.Right = x + Width
TheRect.Bottom = y + Height
End Sub

'' Returns a RECT
'Function MakeRect2(Left As Integer, Top As Integer, right As Integer, Left As Integer) As RECT
'On Error Resume Next
'MakeRect.Left = Left
'MakeRect.Top = Top
'MakeRect.right = right
'MakeRect.Bottom = Bottom
'End Function
'
'' Sets the values of a RECT
'Sub SetRect2(TheRect As RECT, Left As Integer, Top As Integer, right As Integer, Left As Integer)
'On Error Resume Next
'TheRect.Left = Left
'TheRect.Top = Top
'TheRect.right = right
'TheRect.Bottom = Bottom
'End Sub





'**************************************************
'                   ANGLE STUFF
'**************************************************

Function Deg2Rad(DegAngle As Single) As Single
Deg2Rad = DegAngle * PIdiv180
End Function

Function Rad2Deg(RadAngle As Single) As Single
Rad2Deg = RadAngle / PIdiv180
End Function






'**************************************************
'              VECTOR + VERTEX STUFF
'**************************************************

' moves a vertex by adding a vector
Sub MoveVertex(v As D3DVERTEX, T As D3DVECTOR)
On Error Resume Next
   v.x = v.x + T.x
   v.y = v.y + T.y
   v.Z = v.Z + T.Z
End Sub

' finds the average of two vectors
Function AverageOf2Vectors(vec1 As D3DVECTOR, vec2 As D3DVECTOR) As D3DVECTOR
AverageOf2Vectors = AddVector(vec1, vec2)
AverageOf2Vectors.x = AverageOf2Vectors.x / 2
AverageOf2Vectors.y = AverageOf2Vectors.y / 2
AverageOf2Vectors.Z = AverageOf2Vectors.Z / 2
End Function

' finds the vector average of two vectices
Function AverageOf2Vertices(v1 As D3DVERTEX, v2 As D3DVERTEX) As D3DVECTOR
Dim vec1 As D3DVECTOR
Dim vec2 As D3DVECTOR
CopyVertex2Vec v1, vec1
CopyVertex2Vec v2, vec2
AverageOf2Vertices = AddVector(vec1, vec2)
AverageOf2Vertices.x = AverageOf2Vertices.x / 2
AverageOf2Vertices.y = AverageOf2Vertices.y / 2
AverageOf2Vertices.Z = AverageOf2Vertices.Z / 2
End Function

' copies vector info into a vertex
Sub CopyVec2Vertex(srcVec As D3DVECTOR, destVert As D3DVERTEX)
destVert.x = srcVec.x
destVert.y = srcVec.y
destVert.Z = srcVec.Z
End Sub

' copies vertex info into a vector
Sub CopyVertex2Vec(srcVert As D3DVERTEX, destVec As D3DVECTOR)
destVec.x = srcVert.x
destVec.y = srcVert.y
destVec.Z = srcVert.Z
End Sub

' copies vector info into a vertex normal
Sub CopyVec2VertexNormal(srcVec As D3DVECTOR, destVert As D3DVERTEX)
destVert.nx = srcVec.x
destVert.ny = srcVec.y
destVert.nz = srcVec.Z
End Sub

' copies vertex normal info into a vector
Sub CopyVertexNormal2Vec(srcVert As D3DVERTEX, destVec As D3DVECTOR)
destVec.x = srcVert.nx
destVec.y = srcVert.ny
destVec.Z = srcVert.nz
End Sub

' Returns a vector
Function MakeVector(x As Single, y As Single, Z As Single) As D3DVECTOR
Dim Vector As D3DVECTOR
With Vector
    .x = x
    .y = y
    .Z = Z
End With
MakeVector = Vector
End Function

' returns a vertex
Function MakeVertex2(x As Single, y As Single, Z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As D3DVERTEX
With makevertex
    .x = x
    .y = y
    .Z = Z
    .nx = nx
    .ny = ny
    .nz = nz
    .tu = tu
    .tv = tv
End With
End Function

Sub DebugVertex(v As D3DVERTEX)
With v
    Debug.Print .x, .y, .Z, .nx, .ny, .nz, .tu, .tv
End With
End Sub

' puts the cross product of two vectors in a destination vector
Sub VectorCrossProduct(dest As D3DVECTOR, a As D3DVECTOR, b As D3DVECTOR)
   dest.x = a.y * b.Z - a.Z * b.y
   dest.y = a.Z * b.x - a.x * b.Z
   dest.Z = a.x * b.y - a.y * b.x
End Sub

' returns the dot product of two vectors
Function VectorDotProduct(a As D3DVECTOR, b As D3DVECTOR) As Single
  VectorDotProduct = a.x * b.x + a.y * b.y + a.Z * b.Z
End Function

' returns the addition of 2 vectors
Function AddVector(a As D3DVECTOR, b As D3DVECTOR) As D3DVECTOR
AddVector.x = a.x + b.x
AddVector.y = a.y + b.y
AddVector.Z = a.Z + b.Z
End Function

' returns the subraction of one vector from another
Function SubtractVector(a As D3DVECTOR, b As D3DVECTOR) As D3DVECTOR
SubtractVector.x = a.x - b.x
SubtractVector.y = a.y - b.y
SubtractVector.Z = a.Z - b.Z
End Function

' puts the subtraction of two vectors into a destination vector
Sub VectorSubtract(dest As D3DVECTOR, a As D3DVECTOR, b As D3DVECTOR)
  dest.x = a.x - b.x
  dest.y = a.y - b.y
  dest.Z = a.Z - b.Z
End Sub

' finds the centre of a list of vertices
Function CentreOfVertices(v() As D3DVERTEX) As D3DVECTOR
Dim i As Integer
For i = 0 To UBound(v)
   CentreOfVerts.x = CentreOfVerts.x + v(i).x
   CentreOfVerts.y = CentreOfVerts.y + v(i).y
   CentreOfVerts.Z = CentreOfVerts.Z + v(i).Z
Next
CentreOfVerts.x = CentreOfVerts.x / UBound(v)
CentreOfVerts.y = CentreOfVerts.y / UBound(v)
CentreOfVerts.Z = CentreOfVerts.Z / UBound(v)
End Function

Function ScaleVectorFromVector(v As D3DVECTOR, S As D3DVECTOR, C As D3DVECTOR) As D3DVECTOR
Dim d As D3DVECTOR
Dim matRot As D3DMATRIX
Dim matMove As D3DMATRIX
d = SubtractVector(v, C)
d.x = d.x * S.x
d.y = d.y * S.y
d.Z = d.Z * S.Z
ScaleVectorFromVector = AddVector(d, C)
End Function

' rotates a point around an origin
Function RotateXVectorAroundVector(v As D3DVECTOR, Rot As Single, C As D3DVECTOR) As D3DVECTOR
Dim d As D3DVECTOR
Dim matRot As D3DMATRIX
Dim matMove As D3DMATRIX
d = SubtractVector(v, C)
DX.IdentityMatrix matRot
DX.IdentityMatrix matMove
matMove.rc41 = d.x
matMove.rc42 = d.y
matMove.rc43 = d.Z
DX.RotateXMatrix matRot, Rot
DX.MatrixMultiply matMove, matMove, matRot
d.x = matMove.rc41
d.y = matMove.rc42
d.Z = matMove.rc43
RotateXVectorAroundVector = AddVector(d, C)
End Function

' rotates a point around an origin
Function RotateYVectorAroundVector(v As D3DVECTOR, Rot As Single, C As D3DVECTOR) As D3DVECTOR
Dim d As D3DVECTOR
Dim matRot As D3DMATRIX
Dim matMove As D3DMATRIX
d = SubtractVector(v, C)
DX.IdentityMatrix matRot
DX.IdentityMatrix matMove
matMove.rc41 = d.x
matMove.rc42 = d.y
matMove.rc43 = d.Z
DX.RotateYMatrix matRot, Rot
DX.MatrixMultiply matMove, matMove, matRot
d.x = matMove.rc41
d.y = matMove.rc42
d.Z = matMove.rc43
RotateYVectorAroundVector = AddVector(d, C)
End Function

' rotates a point around an origin
Function RotateZVectorAroundVector(v As D3DVECTOR, Rot As Single, C As D3DVECTOR) As D3DVECTOR
Dim d As D3DVECTOR
Dim matRot As D3DMATRIX
Dim matMove As D3DMATRIX
d = SubtractVector(v, C)
DX.IdentityMatrix matRot
DX.IdentityMatrix matMove
matMove.rc41 = d.x
matMove.rc42 = d.y
matMove.rc43 = d.Z
DX.RotateZMatrix matRot, Rot
DX.MatrixMultiply matMove, matMove, matRot
d.x = matMove.rc41
d.y = matMove.rc42
d.Z = matMove.rc43
RotateZVectorAroundVector = AddVector(d, C)
End Function

' returns the normal of a plane defined by 3 points
Function NormalOfPlane(v1 As D3DVERTEX, v2 As D3DVERTEX, v3 As D3DVERTEX) As D3DVECTOR
Dim vec1 As D3DVECTOR
Dim vec2 As D3DVECTOR
Dim vec3 As D3DVECTOR
Dim vec4 As D3DVECTOR
Dim vec5 As D3DVECTOR
CopyVertex2Vec v1, vec1
CopyVertex2Vec v2, vec2
CopyVertex2Vec v3, vec3
VectorSubtract vec4, vec2, vec1
VectorSubtract vec5, vec3, vec2
VectorCrossProduct NormalOfPlane, vec4, vec5
End Function

' returns if 2 vertices are in the same position, to a degree of accurary
Function SimilarVertices(v1 As D3DVERTEX, v2 As D3DVERTEX, GiveOrTake As Single) As Boolean
If Abs(v1.x - v2.x) > GiveOrTake Then Exit Function
If Abs(v1.y - v2.y) > GiveOrTake Then Exit Function
If Abs(v1.Z - v2.Z) > GiveOrTake Then Exit Function
SimilarVertices = True
End Function







'**************************************************
'                  MATRIX STUFF
'**************************************************

Sub PostionMatrix(Mat As D3DMATRIX, vec As D3DVECTOR)
With Mat
    .rc41 = vec.x
    .rc42 = vec.y
    .rc43 = vec.Z
End With
End Sub






'**************************************************
'                  COLOUR STUFF
'**************************************************

Function MakeD3DCOLORVALUE(a As Single, r As Single, g As Single, b As Single) As D3DCOLORVALUE
With MakeD3DCOLORVALUE
  .a = a
  .r = r
  .g = g
  .b = b
End With
End Function

Sub Long2RGB(LongCol As Long, r As Single, g As Single, b As Single)
r = LongCol And 255
g = (LongCol And 65280) \ 256&
b = (LongCol And 16711680) \ 65535
End Sub

Function ContrastColor(LongCol As Long) As Long
Dim r As Byte, g As Byte, b As Byte
r = LongCol And 255
g = (LongCol And 65280) \ 256&
b = (LongCol And 16711680) \ 65535
r = 255 - r
g = 255 - g
b = 255 - b
ContrastColor = RGB(r, g, b)
End Function

Function GreyScale(LongCol As Long) As Single
Dim r As Single
Dim g As Single
Dim b As Single
Long2RGB LongCol, r, g, b
GreyScale = (r + b + g) / 765
End Function
