Attribute VB_Name = "MOD_LAND"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

' loads a landscape
Function LoadLandscape(hdcLand As Long, picWidth As Integer, Optional MeshSize As Single = 10, Optional MeshDetail As Integer = 10, Optional Height As Single = 10, Optional refMtrl As Byte = 0, Optional refTex As Byte = 0) As Boolean
Dim GridSize As Single
Dim x As Integer
Dim y As Integer
Dim col As Long
Dim xx As Single
Dim yy As Single
Dim Poly As tPrim
On Error GoTo LoadWentWrong
MOD_SI3D.Restart
GridSize = picWidth / MeshSize
xx = 0
For x = 1 To MeshSize
xx = xx + GridSize
yy = 0
For y = 1 To MeshSize
GetPixel hdcLand, xx, yy
yy = yy + GridSize
    With Poly
        .PrimType = D3DPT_TRIANGLESTRIP
        ReDim .v(0 To 3)
        With .v(0)
             .x = x
             .Z = y + 1
             col = GetPixel(hdcLand, xx, yy + GridSize)
             If col = -1 Then
                .y = 0
             Else
                .y = Height * MOD_MATH.GreyScale(col)
             End If
        End With
        With .v(1)
             .x = x + 1
             .Z = y + 1
             col = GetPixel(hdcLand, xx + GridSize, yy + GridSize)
             If col = -1 Then
                .y = 0
             Else
                .y = Height * MOD_MATH.GreyScale(col)
             End If
        End With
        With .v(2)
             .x = x
             .Z = y
             col = GetPixel(hdcLand, xx, yy)
             If col = -1 Then
                .y = 0
             Else
                .y = Height * MOD_MATH.GreyScale(col)
             End If
        End With
        With .v(3)
             .x = x + 1
             .Z = y
             col = GetPixel(hdcLand, xx + GridSize, yy)
             If col = -1 Then
                .y = 0
             Else
                .y = Height * MOD_MATH.GreyScale(col)
             End If
        End With
        .refMtrl = refMtrl
        .refTex = refTex
    End With
    MOD_GEO.CalculateNormals Poly
    MOD_SI3D.AddPrim Poly
Next
Next
LoadLandscape = True
Exit Function
LoadWentWrong:
LoadLandscape = False
End Function


