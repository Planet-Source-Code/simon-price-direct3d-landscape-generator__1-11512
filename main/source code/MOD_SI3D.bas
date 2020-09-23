Attribute VB_Name = "MOD_SI3D"

' the polygons use this data format
Public Type tPrim
    v() As D3DVERTEX
    PrimType As CONST_D3DPRIMITIVETYPE
    refMtrl As Byte
    refTex As Byte
End Type
Public Prim() As tPrim

' remembers a list of polygons
Public Type tPrimGroup
    refPrim() As Integer
    Tag As String
End Type
Public Group() As tPrimGroup

' stores a material
Public Type tMtrl
    Mtrl As D3DMATERIAL7
    Tag As String
End Type
Public Mtrl() As tMtrl
Public DefaultMtrl As tMtrl

' stores a texture
Public Type tTex
    Tex As DirectDrawSurface7
    Trans As Long
    Tag As String
End Type
Public Tex() As tTex

' stores a light
Public Type tLight
    Light As D3DLIGHT7
    Tag As String
End Type
Public Light() As tLight

' stores a camera
Public Type tCamera
    vec As D3DVECTOR
    N As D3DVECTOR
    Tag As String
End Type
Public Camera() As tCamera

' possible si3d errors
Public Enum SI3D_ERR
     SI3D_OK = 0
     InvalidFormat = 1
     OldVersion = 2
     NewVersion = 3
     MissingMtrlFile = 4
     MissingTextureFile = 5
     MissingLightFile = 6
     Unknown = 255
End Enum

' translations
Public Enum SI3D_TRANSFORMATION
     SI3D_TRANS_XPLUS = 1
     SI3D_TRANS_YPLUS = 2
     SI3D_TRANS_ZPLUS = 3
     SI3D_TRANS_XMINUS = 4
     SI3D_TRANS_YMINUS = 5
     SI3D_TRANS_ZMINUS = 6
     SI3D_ROT_XPLUS = 7
     SI3D_ROT_YPLUS = 8
     SI3D_ROT_ZPLUS = 9
     SI3D_ROT_XMINUS = 10
     SI3D_ROT_YMINUS = 11
     SI3D_ROT_ZMINUS = 12
     SI3D_SCALE_XPLUS = 13
     SI3D_SCALE_YPLUS = 14
     SI3D_SCALE_ZPLUS = 15
     SI3D_SCALE_XMINUS = 16
     SI3D_SCALE_YMINUS = 17
     SI3D_SCALE_ZMINUS = 18
End Enum

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long



     


'***************************************************
'
'                      START UP
'
'***************************************************

' call to start engine
Sub StartUp()
On Error Resume Next
ZeroArrays
With DefaultMtrl
   .Tag = "Default Mtrl"
   .Mtrl = MakeMtrl()
End With
D3DDEV.SetMaterial DefaultMtrl.Mtrl
End Sub

' clears arrays
Sub ZeroArrays()
On Error Resume Next
ReDim Prim(0)
ReDim Group(0)
ReDim Mtrl(0)
ReDim Tex(0)
ReDim Light(0)
ReDim Camera(0)
End Sub

' emptys engine contents
Sub Restart()
On Error Resume Next
ZeroArrays
End Sub






'+++++++++++++++++++++++++++++++++++++++++++++++++++
'
'                    MATERIALS
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++

' makes a material structure from the components given
Function MakeMtrl(Optional aa As Single = 1, Optional ar As Single = 1, Optional ag As Single = 1, Optional ab As Single = 1, Optional da As Single = 1, Optional dr As Single = 1, Optional dg As Single = 1, Optional db As Single = 1, Optional ea As Single = 0, Optional er As Single = 0, Optional eg As Single = 0, Optional eb As Single = 0, Optional sa As Single = 1, Optional sr As Single = 1, Optional sg As Single = 1, Optional sb As Single = 1, Optional p As Single = 5) As D3DMATERIAL7
With MakeMtrl
    With .Ambient
        .a = aa
        .r = ar
        .g = ag
        .b = ab
    End With
    With .Diffuse
        .a = da
        .r = dr
        .g = dg
        .b = db
    End With
    With .emissive
        .a = ea
        .r = er
        .g = eg
        .b = eb
    End With
    With .Specular
        .a = sa
        .r = sr
        .g = sg
        .b = sb
    End With
    .power = p
End With
End Function

' adds a new material the the engine
Function AddMtrl(NewMtrl As D3DMATERIAL7, Tag As String) As Boolean
ReDim Preserve Mtrl(0 To UBound(Mtrl) + 1)
With Mtrl(UBound(Mtrl))
    .Tag = Tag
    .Mtrl = NewMtrl
End With
End Function

' saves a material in the .mtrl file format
Function SaveMtrl(Filename As String, TheMtrl As D3DMATERIAL7) As Boolean
On Error GoTo FileMessedUp
Open Filename & ".mtrl" For Random As #3 Len = 4
With TheMtrl
    With .Ambient
        Put #3, 1, .a
        Put #3, 2, .r
        Put #3, 3, .g
        Put #3, 4, .b
    End With
    With .Diffuse
        Put #3, 5, .a
        Put #3, 6, .r
        Put #3, 7, .g
        Put #3, 8, .b
    End With
    With .emissive
        Put #3, 9, .a
        Put #3, 10, .r
        Put #3, 11, .g
        Put #3, 12, .b
    End With
    With .Specular
        Put #3, 13, .a
        Put #3, 14, .r
        Put #3, 15, .g
        Put #3, 16, .b
    End With
    Put #3, 17, .power
End With
Close #3
SaveMtrl = True
Exit Function
FileMessedUp:
SaveMtrl = False
End Function

' loads a material in the .mtrl file format
Function LoadMtrl(Filename As String, TheMtrl As D3DMATERIAL7) As Boolean
On Error GoTo FileMessedUp
Open Filename & ".mtrl" For Random As #3 Len = 4
With TheMtrl
    With .Ambient
        Get #3, 1, .a
        Get #3, 2, .r
        Get #3, 3, .g
        Get #3, 4, .b
    End With
    With .Diffuse
        Get #3, 5, .a
        Get #3, 6, .r
        Get #3, 7, .g
        Get #3, 8, .b
    End With
    With .emissive
        Get #3, 9, .a
        Get #3, 10, .r
        Get #3, 11, .g
        Get #3, 12, .b
    End With
    With .Specular
        Get #3, 13, .a
        Get #3, 14, .r
        Get #3, 15, .g
        Get #3, 16, .b
    End With
    Get #3, 17, .power
End With
Close #3
LoadMtrl = True
Exit Function
FileMessedUp:
Close #3
LoadMtrl = False
End Function







'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'                     TEXTURES
'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Sub AddTex(Filename As String, TheTex As tTex)
On Error Resume Next
Dim i As Long
Dim IsFound As Boolean
Dim ddsd As DDSURFACEDESC2
ReDim Preserve Tex(0 To UBound(Tex) + 1)
'Set Tex(UBound(Tex)).Tex = MOD_D3D.CreateTextureSurface(Filename & ".bmp")

ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_PIXELFORMAT Or DDSD_TEXTURESTAGE
Dim TextureEnum As Direct3DEnumPixelFormats
Set TextureEnum = D3DDEV.GetTextureFormatsEnum()
For i = 1 To TextureEnum.GetCount()
    IsFound = True
    TextureEnum.GetItem i, ddsd.ddpfPixelFormat
    With ddsd.ddpfPixelFormat
        If .lFlags And (DDPF_LUMINANCE Or DDPF_BUMPLUMINANCE Or DDPF_BUMPDUDV) Then IsFound = False
        If .lFourCC <> 0 Then IsFound = False
        If .lFlags And DDPF_ALPHAPIXELS Then IsFound = False
        If .lRGBBitCount <> 16 Then IsFound = False
    End With
    If IsFound Then Exit For
Next i
If Not IsFound Then
    MsgBox "Unable to locate 16-bit surface support on your hardware."
    End
End If
ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
ddsd.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
ddsd.lTextureStage = 0
Set Tex(UBound(Tex)).Tex = DDRAW.CreateSurfaceFromFile(Filename & ".bmp", ddsd)
Tex(UBound(Tex)).Tag = TheTex.Tag
Tex(UBound(Tex)).Trans = TheTex.Trans
End Sub

Function SaveTex(Filename As String, TheTex As tTex, Width As Integer, Height As Integer) As Boolean
With TheTex
On Error GoTo FileMuffUp
Open Filename & ".tex" For Output As #4
Write #4, .Tag, .Trans, Width, Height
Close #4
SaveTex = True
Exit Function
FileMuffUp:
SaveTex = False
End With
End Function

Function LoadTex(Filename As String, TheTex As tTex, Width As Integer, Height As Integer) As Boolean
With TheTex
On Error GoTo FileMuffUp
Open Filename & ".tex" For Input As #4
Input #4, .Tag, .Trans, Width, Height
Close #4
LoadTex = True
Exit Function
FileMuffUp:
LoadTex = False
End With
End Function








'|||||||||||||||||||||||||||||||||||||||||||||||||||
'
'                      GROUPS
'
'|||||||||||||||||||||||||||||||||||||||||||||||||||

' adds a group to the engine
Sub AddGroup(NewGroup As tPrimGroup)
ReDim Preserve Group(0 To UBound(Group) + 1)
Group(UBound(Group)) = NewGroup
End Sub

' deletes a group
Sub DeleteGroup(refGroup As Integer)
On Error Resume Next
Group(refGroup) = Group(UBound(Group))
ReDim Preserve Group(0 To UBound(Group) - 1)
End Sub

' returns a group
Function MakeGroup(refPrim() As Integer, Tag As String) As tPrimGroup
MakeGroup.refPrim = refPrim()
MakeGroup.Tag = Tag
End Function

' adds a new polygon to a group
Sub AddPrim2Group(refGroup As Integer, refPrim As Integer)
With Group(refGroup)
ReDim Preserve .refPrim(0 To UBound(.refPrim) + 1)
.refPrim(UBound(.refPrim)) = refPrim
End With
End Sub

' deletes a polygon from a group
Sub DeletePrimFromGroup(refGroup As Integer, refPrim As Integer)
With Group(refGroup)
.refPrim(refPrim) = .refPrim(UBound(.refPrim))
ReDim Preserve .refPrim(0 To UBound(.refPrim) + 1)
End With
End Sub









'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'
'                    PRIMITIVES
'
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

' adds a new primitive
Sub AddPrim(NewPrim As tPrim)
ReDim Preserve Prim(0 To UBound(Prim) + 1)
Prim(UBound(Prim)) = NewPrim
End Sub

' deletes a primitive
Sub DeletePrim(refPrim As Integer)
On Error Resume Next
Prim(refPrim) = Prim(UBound(Prim))
ReDim Preserve Prim(0 To UBound(Prim) - 1)
End Sub

' returns a primtive
Function MakePrim(v() As D3DVERTEX, PrimType As CONST_D3DPRIMITIVETYPE, refMtrl As Byte, refTex As Byte) As tPrim
With MakePrim
    .v = v()
    .PrimType = PrimType
    .refMtrl = refMtrl
    .refTex = refTex
End With
End Function

Sub DebugPrim(refPrim As Integer)
Dim i As Byte
With Prim(refPrim)
   Debug.Print "Debug Info on Primitive ref no. " & refPrim
   Debug.Print .PrimType
   Debug.Print .refMtrl
   Debug.Print .refTex
   For i = 0 To UBound(.v)
       MOD_MATH.DebugVertex .v(i)
   Next
End With
End Sub






'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                     VERTICES
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

' moves a vertex
Sub TranslateVertex(refPrim As Integer, refVertex As Byte, x As Single, y As Single, Z As Single)
With Prim(refPrim).v(refVertex)
    .x = .x + x
    .y = .y + y
    .Z = .Z + Z
End With
End Sub






'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'
'                      CAMERAS
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

' adds a camera
Sub AddCamera(NewCamera As tCamera)
ReDim Preserve Camera(0 To UBound(Camera) + 1)
Camera(UBound(Camera)) = NewCamera
End Sub

' returns a camera
Function MakeCamera(vec As D3DVECTOR, N As D3DVECTOR, Tag As String) As tCamera
With MakeCamera
    .vec = vec
    .N = N
    .Tag = Tag
End With
End Function

Sub TransformCamera(refCamera As Byte, WhatWay As SI3D_TRANSFORMATION, HowMuch As Single)
Dim S As Single
Dim matMove As D3DMATRIX
Dim matRot As D3DMATRIX
Const MULT = 180 / PI
DX.IdentityMatrix matMove
S = HowMuch
With Camera(refCamera)
Select Case WhatWay
    Case SI3D_TRANS_XMINUS
        MOD_MATH.PostionMatrix matMove, MOD_MATH.MakeVector(0, 0, -1)
        DX.RotateXMatrix matRot, .N.x
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateYMatrix matRot, .N.y + Deg2Rad(90)
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateZMatrix matRot, .N.Z
        DX.MatrixMultiply matMove, matMove, matRot
        .vec.x = .vec.x + matMove.rc41 * S
        .vec.y = .vec.y + matMove.rc42 * S
        .vec.Z = .vec.Z + matMove.rc43 * S
    Case SI3D_TRANS_XPLUS
        MOD_MATH.PostionMatrix matMove, MOD_MATH.MakeVector(0, 0, -1)
        DX.RotateXMatrix matRot, .N.x
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateYMatrix matRot, .N.y + Deg2Rad(90)
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateZMatrix matRot, .N.Z
        DX.MatrixMultiply matMove, matMove, matRot
        .vec.x = .vec.x - matMove.rc41 * S
        .vec.y = .vec.y - matMove.rc42 * S
        .vec.Z = .vec.Z - matMove.rc43 * S
    Case SI3D_TRANS_YMINUS
        MOD_MATH.PostionMatrix matMove, MOD_MATH.MakeVector(0, 0, -1)
        DX.RotateXMatrix matRot, .N.x + Deg2Rad(90)
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateYMatrix matRot, .N.y
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateZMatrix matRot, .N.Z
        DX.MatrixMultiply matMove, matMove, matRot
        .vec.x = .vec.x - matMove.rc41 * S
        .vec.y = .vec.y - matMove.rc42 * S
        .vec.Z = .vec.Z - matMove.rc43 * S
    Case SI3D_TRANS_YPLUS
        MOD_MATH.PostionMatrix matMove, MOD_MATH.MakeVector(0, 0, -1)
        DX.RotateXMatrix matRot, .N.x + Deg2Rad(90)
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateYMatrix matRot, .N.y
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateZMatrix matRot, .N.Z
        DX.MatrixMultiply matMove, matMove, matRot
        .vec.x = .vec.x + matMove.rc41 * S
        .vec.y = .vec.y + matMove.rc42 * S
        .vec.Z = .vec.Z + matMove.rc43 * S
    Case SI3D_TRANS_ZMINUS
        MOD_MATH.PostionMatrix matMove, MOD_MATH.MakeVector(0, 0, -1)
        DX.RotateXMatrix matRot, .N.x
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateYMatrix matRot, .N.y
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateZMatrix matRot, .N.Z
        DX.MatrixMultiply matMove, matMove, matRot
        .vec.x = .vec.x + matMove.rc41 * S
        .vec.y = .vec.y + matMove.rc42 * S
        .vec.Z = .vec.Z + matMove.rc43 * S
    Case SI3D_TRANS_ZPLUS
        MOD_MATH.PostionMatrix matMove, MOD_MATH.MakeVector(0, 0, -1)
        DX.RotateXMatrix matRot, .N.x
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateYMatrix matRot, .N.y
        DX.MatrixMultiply matMove, matMove, matRot
        DX.RotateZMatrix matRot, .N.Z
        DX.MatrixMultiply matMove, matMove, matRot
        .vec.x = .vec.x - matMove.rc41 * S
        .vec.y = .vec.y - matMove.rc42 * S
        .vec.Z = .vec.Z - matMove.rc43 * S
'    Case SI3D_TRANS_YPLUS
'        .vec.y = .vec.y - S
'    Case SI3D_TRANS_YMINUS
'        .vec.y = .vec.y + S
'    Case SI3D_TRANS_XMINUS
'        .vec.x = .vec.x + S
'    Case SI3D_TRANS_XPLUS
'        .vec.x = .vec.x - S
'    Case SI3D_TRANS_ZMINUS
'        .vec.z = .vec.z - S
'    Case SI3D_TRANS_ZPLUS
'        .vec.z = .vec.z + S
    Case SI3D_ROT_YPLUS
        .N.y = .N.y - S
    Case SI3D_ROT_YMINUS
        .N.y = .N.y + S
    Case SI3D_ROT_XMINUS
        .N.x = .N.x + S
    Case SI3D_ROT_XPLUS
        .N.x = .N.x - S
    Case SI3D_ROT_ZMINUS
        .N.Z = .N.Z - S
    Case SI3D_ROT_ZPLUS
        .N.Z = .N.Z + S
End Select
End With
End Sub

' sets the current camera
Sub SetCamera(ByVal refCamera As Byte)
Dim matView As D3DMATRIX
Dim matTemp As D3DMATRIX
Dim matRot As D3DMATRIX
With Camera(refCamera)
    DX.IdentityMatrix matView
    matView.rc41 = -.vec.x
    matView.rc42 = -.vec.y
    matView.rc43 = -.vec.Z
    DX.IdentityMatrix matRot
    ' pitch
    DX.RotateXMatrix matTemp, -.N.x
    DX.MatrixMultiply matRot, matRot, matTemp
    ' yaw
    DX.RotateYMatrix matTemp, -.N.y
    DX.MatrixMultiply matRot, matRot, matTemp
    ' roll
    DX.RotateZMatrix matTemp, -.N.Z
    DX.MatrixMultiply matRot, matRot, matTemp
End With

DX.MatrixMultiply matView, matView, matRot
D3DDEV.SetTransform D3DTRANSFORMSTATE_VIEW, matView
End Sub
 
'' moves the entire engines contents
'Sub MoveWorld(Camera As tCamera)
'Dim matWorld As D3DMATRIX
'Dim matTemp As D3DMATRIX
'Dim matRot As D3DMATRIX
'
'DX.IdentityMatrix matWorld
'matWorld.rc41 = Camera.Pos.x
'matWorld.rc42 = Camera.Pos.y
'matWorld.rc43 = Camera.Pos.z
'DX.IdentityMatrix matRot
'' pitch
'DX.RotateXMatrix matTemp, Camera.Dir.x
'DX.MatrixMultiply matRot, matRot, matTemp
'' yaw
'DX.RotateYMatrix matTemp, Camera.Dir.y
'DX.MatrixMultiply matRot, matRot, matTemp
'' roll
'DX.RotateZMatrix matTemp, Camera.Dir.z
'DX.MatrixMultiply matRot, matRot, matTemp
'
'DX.MatrixMultiply matWorld, matRot, matWorld
'D3DDEV.SetTransform D3DTRANSFORMSTATE_WORLD, matWorld
'End Sub







'===================================================
'
'                    RENDERING
'
'===================================================

' renders the scene
Function RenderScene() As Boolean
On Error Resume Next
Dim i As Integer
Dim LastMtrl As Byte, LastTex As Byte

D3DDEV.BeginScene
D3DDEV.SetMaterial DefaultMtrl.Mtrl
D3DDEV.SetTexture 0, Nothing

For i = 1 To UBound(Prim)
    With Prim(i)
        If LastMtrl <> .refMtrl Then
             D3DDEV.SetMaterial Mtrl(.refMtrl).Mtrl
             LastMtrl = .refMtrl
        End If
        If LastTex <> .refTex Then
             If .refTex = 0 Then
                  D3DDEV.SetTexture 0, Nothing
             Else
                  D3DDEV.SetTexture 0, Tex(.refTex).Tex
             End If
             LastTex = .refTex
        End If
        D3DDEV.DrawPrimitive .PrimType, D3DFVF_VERTEX, .v(0), UBound(.v) + 1, D3DDP_DEFAULT
    End With
Next

D3DDEV.EndScene

If Err.Number = DD_OK Then RenderScene = True
End Function

' clears the scene
Function Clear(Flags As CONST_D3DCLEARFLAGS, Optional Color As Long = vbBlack, Optional Z As Single = 1, Optional refStencil As Single = 0) As Boolean
D3DDEV.Clear 1, Viewport(), Flags, Color, Z, refStencil
If Err.Number = DD_OK Then Clear = True
End Function






'###################################################
'
'                      LIGHTS
'
'###################################################

' adds a light to the engine
Sub AddLight(NewLight As D3DLIGHT7, Tag As String, Enabled As Boolean)
On Error Resume Next
Dim UBL As Long
UBL = UBound(Light) + 1
ReDim Preserve Light(0 To UBL)
Light(UBL).Light = NewLight
Light(UBL).Tag = Tag
D3DDEV.SetLight UBL, Light(UBL).Light
D3DDEV.LightEnable UBL, Enabled
End Sub

' deletes a light from the engine
Sub DeleteLight(refLight As Byte)
On Error Resume Next
D3DDEV.LightEnable refLight, False
Light(refLight) = Light(UBound(Light))
ReDim Preserve Light(0 To UBound(Light) - 1)
End Sub

Sub EnableLight(refLight As Byte, Enabled As Boolean)
On Error Resume Next
D3DDEV.LightEnable refLight, Enabled
End Sub

Sub EnableAllLights(Enabled As Boolean)
On Error Resume Next
Dim i As Byte
For i = 1 To UBound(Light)
    D3DDEV.LightEnable i, Enabled
Next
End Sub

' returns a light
Function MakeLight(LightType As CONST_D3DLIGHTTYPE, Ambient As D3DCOLORVALUE, Diffuse As D3DCOLORVALUE, Specular As D3DCOLORVALUE, Pos As D3DVECTOR, Dir As D3DVECTOR, Optional attenuation0 As Single = 0, Optional attenuation1 As Single = 0, Optional attenuation2 As Single = 0, Optional FallOff As Single = 0, Optional phi As Single = 0, Optional theta As Single = 0, Optional Range As Single = 0) As D3DLIGHT7
With MakeLight
  .Ambient = Ambient
  .attenuation0 = attenuation0
  .attenuation1 = attenuation1
  .attenuation2 = attenuation2
  .Diffuse = Diffuse
  .direction = Dir
  .dltType = LightType
  .FallOff = FallOff
  .phi = phi
  .position = Pos
  .Range = Range
  .Specular = Specular
  .theta = theta
End With
End Function

' make ambient light white, brightness is a % value
Function SetAmbientLight(Brightness As Byte) As Boolean
Dim C As Single
C = Brightness / 100
D3DDEV.SetRenderState D3DRENDERSTATE_AMBIENT, DX.CreateColorRGBA(C, C, C, C)
End Function

' translates a light
Sub MoveLight(refLight As Byte, T As D3DVECTOR)
With Light(refLight).Light.position
     .x = .x + T.x
     .y = .y + T.y
     .Z = .Z + T.Z
End With
End Sub

' loads a light from a .lgt file
Function LoadLight(Filename As String, TheLight As D3DLIGHT7) As Boolean
On Error GoTo FileMuffUp
Open Filename For Random As #5 Len = 4
With TheLight
    With .Ambient
        Get #5, 1, .a
        Get #5, 2, .r
        Get #5, 3, .g
        Get #5, 4, .b
    End With
        Get #5, 5, .attenuation0
        Get #5, 6, .attenuation1
        Get #5, 7, .attenuation2
    With .Diffuse
        Get #5, 8, .a
        Get #5, 9, .r
        Get #5, 10, .g
        Get #5, 11, .b
    End With
    With .direction
        Get #5, 12, .x
        Get #5, 13, .y
        Get #5, 14, .Z
    End With
        Get #5, 15, .dltType
        Get #5, 16, .FallOff
    With .position
        Get #5, 17, .x
        Get #5, 18, .y
        Get #5, 19, .Z
    End With
        Get #5, 20, .phi
        Get #5, 21, .Range
    With .Specular
        Get #5, 22, .a
        Get #5, 23, .r
        Get #5, 24, .g
        Get #5, 25, .b
    End With
        Get #5, 26, .theta
End With
Close #5
LoadLight = True
Exit Function
FileMuffUp:
On Error Resume Next
Close #5
LoadLight = False
End Function







'###################################################
'
'                    MODEL FILES
'
'###################################################

' loads a .si3dm file
Function LoadModelFile(Filename As String, VerLBound As Byte, VerUbound As Byte) As SI3D_ERR
On Error Resume Next
Restart
DeleteFile App.Path & "\log files\file error log.txt"
Open App.Path & "\log files\file error log.txt" For Output As #255
Print #255, "FILE ERROR LOG FOR SI3D MODELLER V1.0"
Print #255, ""
Print #255, "Attempting to load si3D model file : " & Filename & ".si3dm"
Print #255, ""
On Error GoTo FileMessedUp
Open Filename & ".si3dm" For Input As #1
' check file format
    Dim FileType As String, Major As Byte, Minor As Byte
    Input #1, FileType, Major, Minor
    If FileType <> "SI3D_MODEL_FILE" Then
         LoadModelFile = InvalidFormat
         Print #255, "FATAL LOADING ERROR : The file was of an invalid format, it is either corrupted or has an incorrect file extension!"
         GoTo CloseUp
    End If
' check file version
    If Major < VerLBound Then
         LoadModelFile = OldVersion
         Print #255, "FATAL LOADING ERROR : The file in an old .si3dm file format, it was designed for use with SI3D version " & Major & "." & Minor & "!"
         GoTo CloseUp
    End If
    If Major > VerUbound Then
         LoadModelFile = NewVersion
         Print #255, "FATAL LOADING ERROR : The file in a newer .si3dm file format, it is designed for use with SI3D version " & Major & "." & Minor & "! Try to get hold of the latest version, visit www.VBgames.co.uk for an update!"
         GoTo CloseUp
    End If
' read each keyword in file until the end
    Do
         Dim KeyWord As String
         Input #1, KeyWord
         Select Case KeyWord
              Case "MATERIALS"
                   ReadMaterials
              Case "TEXTURES"
                   ReadTextures
              Case "LIGHTS"
                   ReadLights
              Case "GROUPS"
                   ReadGroups
              Case "PRIMITIVES"
                   ReadPrimitives
                   GoTo FileDone
         End Select
         Print #255, ""
    Loop Until EOF(1)
FileDone:
Close #1
' report loading errors
LoadModelFile = SI3D_OK
Print #255, ""
Print #255, "FILE LOAD COMPLETED SUCCESSFULLY!"
Close #255
Exit Function
FileMessedUp:
On Error Resume Next
LoadModelFile = Unknown
Print #255, ""
Print #255, "FILE LOAD FAILED DUE TO SUDDEN UNKNOWN ERROR!"
Close #255
Exit Function
CloseUp:
On Error Resume Next
Print #255, ""
Print #255, "FILE LOAD FAILED!"
Close #255
Close #1
End Function

' reads the materials in a .si3dm file
Sub ReadMaterials()
On Error Resume Next
Dim i As Byte
Dim Filename As String
Dim NewMtrl As D3DMATERIAL7
Print #255, "Reading Materials..."
Input #1, i
For i = 1 To i
    Input #1, Filename
    If LoadMtrl(App.Path & "\materials\mtrl files\" & Filename, NewMtrl) = False Then
        Print #255, "MATERIAL LOAD FAILED! - check that the material " & Filename & " is available!"
    Else
        Print #255, "Material " & Filename & " loaded OK"
    End If
    AddMtrl NewMtrl, Filename
Next
End Sub

' reads the textures in a .si3dm file
Sub ReadTextures()
On Error Resume Next
Dim i As Byte
Dim Filename As String
Dim NewTex As tTex
Dim w As Integer, h As Integer
Print #255, "Reading Textures..."
Input #1, i
For i = 1 To i
    Input #1, Filename
    If LoadTex(App.Path & "\textures\tex files\" & Filename, NewTex, w, h) = False Then
        Print #255, "TEXTURE LOAD FAILED! - check that the texture " & Filename & " is available!"
    Else
        Print #255, "Texture " & Filename & " loaded OK"
    End If
    AddTex App.Path & "\textures\tex files\bitmaps\" & NewTex.Tag, NewTex
Next
End Sub

' reads the lights in a .si3dm file
Sub ReadLights()
'On Error Resume Next
Dim i As Byte
Dim Filename As String
Dim NewLight As D3DLIGHT7
Print #255, "Reading Lights..."
Input #1, i
For i = 1 To i
    Input #1, Filename
    If LoadLight(App.Path & "\lights\lgt files\" & Filename & ".lgt", NewLight) = False Then
        Print #255, "LIGHT LOAD FAILED! - check that the light " & Filename & " is available!"
    Else
        Print #255, "Light " & Filename & " loaded OK"
    End If
    AddLight NewLight, Filename, True
Next
End Sub

' reads the groups in a .si3dm files
Sub ReadGroups()
On Error Resume Next
Dim refs As Byte
Dim num As Byte
Dim Tag As String
Dim i As Byte
Dim i2 As Byte
Dim refPrim() As Integer
Print #255, "Reading Polygon Group Data..."
Input #1, num
For i = 1 To num
    Input #1, Tag, refs
    ReDim refPrim(0 To refs)
    For i2 = 0 To refs
        Input #1, refPrim(i2)
    Next
    AddGroup MakeGroup(refPrim, Tag)
Next
End Sub

' reads the primitives in a .si3dm file
Sub ReadPrimitives()
On Error Resume Next
Dim strPrimType As String, Vertices As Byte
Dim PrimType As CONST_D3DPRIMITIVETYPE, refMtrl As Byte, refTex As Byte, v() As D3DVERTEX
Dim i As Byte
Print #255, "Reading Polygon Data..."
Do
    Input #1, strPrimType, Vertices, refMtrl, refTex
    If EOF(1) Then Exit Do
    ReDim v(0 To Vertices - 1)
    Select Case strPrimType
        Case "PL"
            PrimType = D3DPT_POINTLIST
        Case "LL"
            PrimType = D3DPT_LINELIST
        Case "LS"
            PrimType = D3DPT_LINESTRIP
        Case "TL"
            PrimType = D3DPT_TRIANGLELIST
        Case "TS"
            PrimType = D3DPT_TRIANGLESTRIP
        Case "TF"
            PrimType = D3DPT_TRIANGLEFAN
    End Select
    For i = 0 To UBound(v)
         With v(i)
             Input #1, .x, .y, .Z, .nx, ny, .nz, .tu, .tv
         End With
    Next
    AddPrim MakePrim(v(), PrimType, refMtrl, refTex)
Loop Until EOF(1)
Print #255, "Polygons Loaded OK, End Of File Found"
End Sub

' saves the clipboard data
Function SaveModelFile(Filename As String, VerMajor As Byte, VerMinor As Byte) As SI3D_ERR
Dim i As Integer
Dim i2 As Integer
On Error Resume Next
Open App.Path & "\log files\file error log.txt" For Output As #255
Print #255, "FILE ERROR LOG FOR SI3D MODELLER V1.0"
Print #255, ""
Print #255, "Attempting to save SI3D model file : " & Filename & ".si3dm"
Print #255, ""
On Error GoTo FileMessedUp
Open Filename & ".si3dm" For Output As #1
    ' print version info
    Write #1, "SI3D_MODEL_FILE", VerMajor, VerMinor
    ' save materials
    Print #255, "Saving Materials..."
    Write #1, "MATERIALS"
    Write #1, UBound(Mtrl)
    For i = 1 To UBound(Mtrl)
        Write #1, Mtrl(i).Tag
    Next
    ' save textures
    Print #255, "Saving Textures..."
    Write #1, "TEXTURES"
    Write #1, UBound(Tex)
    For i = 1 To UBound(Tex)
        Write #1, Tex(i).Tag
    Next
    'save lights
    Print #255, "Saving Lights..."
    Write #1, "LIGHTS"
    Write #1, UBound(Light)
    For i = 1 To UBound(Light)
        Write #1, Light(i).Tag
    Next
    ' save groups
    Print #255, "Saving Polygon Grouping..."
    Write #1, "GROUPS"
    Write #1, UBound(Group)
    For i = 1 To UBound(Group)
        With Group(i)
        Write #1, .Tag, UBound(.refPrim)
        For i2 = 0 To UBound(.refPrim)
            Write #1, .refPrim(i2)
        Next
        End With
    Next
    ' save primitives
    Print #255, "Saving Polygon Data..."
    Write #1, "PRIMITIVES"
    For i = 1 To UBound(Prim)
        With Prim(i)
            Select Case .PrimType
                Case D3DPT_POINTLIST
                   Write #1, "PL", UBound(.v) + 1, .refMtrl, .refTex
                Case D3DPT_LINELIST
                   Write #1, "LL", UBound(.v) + 1, .refMtrl, .refTex
                Case D3DPT_LINESTRIP
                   Write #1, "LS", UBound(.v) + 1, .refMtrl, .refTex
                Case D3DPT_TRIANGLELIST
                   Write #1, "TL", UBound(.v) + 1, .refMtrl, .refTex
                Case D3DPT_TRIANGLESTRIP
                   Write #1, "TS", UBound(.v) + 1, .refMtrl, .refTex
                Case D3DPT_TRIANGLEFAN
                   Write #1, "TF", UBound(.v) + 1, .refMtrl, .refTex
             End Select
         For i2 = 0 To UBound(.v)
         With .v(i2)
             Write #1, .x, .y, .Z, .nx, .ny, .nz, .tu, .tv
         End With
         Next
         End With
    Next
FileDone:
Close #1
On Error Resume Next
' report saving errors
Print #255, "FILE SAVE COMPLETED SUCCESSFULLY!"
Close #255
SaveModelFile = SI3D_OK
Exit Function
FileMessedUp:
On Error Resume Next
Print #255, "FILE SAVE FAILED DUE TO UNKNOWN ERROR!"
Close #255
Close #1
SaveModelFile = Unknown
End Function



