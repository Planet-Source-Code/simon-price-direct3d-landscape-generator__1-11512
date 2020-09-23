Attribute VB_Name = "MOD_D3D"


' the main direct3D object
Public D3D As Direct3D7
' the 3D rendering device
Public D3DDEV As Direct3DDevice7
' remembers what 3D stuff we can do
Public D3Denum As Direct3DEnumDevices
Public DeviceDesc As D3DDEVICEDESC7
'Public HardwareDesc As D3DDEVICEDESC7
'Public EmulationDesc As D3DDEVICEDESC7
' viewport
Public Viewport(0) As D3DRECT
Public VPDesc As D3DVIEWPORT7
' the z-buffer
Public Zbuff As DirectDrawSurface7

Function EnumerateDevices(Guid As String, DeviceDescription As String, DeviceName As String, BPP As Byte, Optional NeedGouraud As Boolean = False) As Boolean
On Error Resume Next
Dim iDevice As Integer
Dim DevDesc As D3DDEVICEDESC7
Dim IsHardware As Boolean
For iDevice = 1 To D3Denum.GetCount
    Dim CheckDesc As D3DDEVICEDESC7
    D3Denum.GetDesc iDevice, DevDesc
    IsHardware = (DevDesc.lDevCaps Or D3DDEVCAPS_HWRASTERIZATION)
    If CheckDesc.lDeviceRenderBitDepth And BPP Then
        If NeedGouraud Then
            If (Not CheckDesc.dpcTriCaps.lShadeCaps And D3DPSHADECAPS_COLORGOURAUDRGB) Then GoTo Next_Device
        End If
        Guid = D3Denum.GetGuid(iDevice)
        DeviceDescription = D3Denum.GetDescription(iDevice)
        DeviceName = D3Denum.GetName(iDevice)
        DeviceDesc = DevDesc
        If IsHardware = True Then Exit For
    End If
Next_Device:
Next
If Err.Number = DD_OK And IsHardware Then EnumerateDevices = True
End Function

' creates direct3d
Function StartUp() As Boolean
On Error Resume Next
Set D3D = DDRAW.GetDirect3D
Set D3Denum = D3D.GetDevicesEnum
If Err.Number = DD_OK Then StartUp = True
End Function

' creates the Direct3D device, the viewport rectangle and projection matrix
Function StartUpDevice(Guid As String, Surf As DirectDrawSurface7, Width As Integer, Height As Integer, Optional MinZ As Single = 0, Optional MaxZ As Single = 1) As Boolean
Set D3DDEV = D3D.CreateDevice(Guid, Surf)

VPDesc.lWidth = Width
VPDesc.lHeight = Height
VPDesc.MinZ = MinZ
VPDesc.MaxZ = MaxZ
D3DDEV.SetViewport VPDesc
With Viewport(0)
    .X1 = 0: .Y1 = 0
    .X2 = Width
    .Y2 = Height
End With

Dim matProj As D3DMATRIX
DX.IdentityMatrix matProj
DX.ProjectionMatrix matProj, 1, 1000, PI / 3
D3DDEV.SetTransform D3DTRANSFORMSTATE_PROJECTION, matProj

If Err.Number = DD_OK Then StartUpDevice = True
End Function

' creates a z-buffer and attach it to a surface
Function AttachZbuffer(Surf As DirectDrawSurface7, BPP As Byte, Guid As String, Width As Integer, Height As Integer, Optional UseVideoMem As Boolean = False) As Boolean
On Error Resume Next
Dim ddpfZBuffer As DDPIXELFORMAT
Dim d3dEnumPFs As Direct3DEnumPixelFormats
Dim i As Long
Dim SurfDesc As DDSURFACEDESC2

Set d3dEnumPFs = D3D.GetEnumZBufferFormats(Guid)
For i = 1 To d3dEnumPFs.GetCount()
    d3dEnumPFs.GetItem i, ddpfZBuffer
'    If ddpfZBuffer.lFlags = DDPF_ZBUFFER Then
'    If ddpfZBuffer.lZBufferBitDepth = BPP Then Exit For
'    End If
    If ddpfZBuffer.lFlags = DDPF_ZBUFFER Then Exit For
Next i
If Err.Number <> DD_OK Then Exit Function

SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_PIXELFORMAT
If UseVideoMem = False Then
    SurfDesc.ddsCaps.lCaps = DDSCAPS_ZBUFFER Or DDSCAPS_SYSTEMMEMORY
Else
    SurfDesc.ddsCaps.lCaps = DDSCAPS_ZBUFFER Or DDSCAPS_VIDEOMEMORY
End If
SurfDesc.lWidth = Width
SurfDesc.lHeight = Height
SurfDesc.ddpfPixelFormat = ddpfZBuffer
'SurfDesc.ddsCaps.lCaps = SurfDesc.ddsCaps.lCaps Or DDSCAPS_SYSTEMMEMORY

Set Zbuff = DDRAW.CreateSurface(SurfDesc)
If Err.Number <> DD_OK Then Exit Function
Surf.AddAttachedSurface Zbuff
If Err.Number = DD_OK Then AttachZbuffer = True
End Function

Public Function CreateTextureSurface(File As String) As DirectDrawSurface7
Dim i As Long
Dim IsFound As Boolean
Dim ddsd As DDSURFACEDESC2

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
Set CreateTextureSurface = DDRAW.CreateSurfaceFromFile(File, ddsd)
End Function

Function EndIt() As Boolean
On Error Resume Next
Set D3D = Nothing
If Err.Number = DD_OK Then EndIt = True
End Function

