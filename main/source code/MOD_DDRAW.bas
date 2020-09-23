Attribute VB_Name = "MOD_DDRAW"
' the main directdraw object thingy
Public DDRAW As DirectDraw7
' the surfaces
Public Primary As DirectDrawSurface7
Public BackBuffer As DirectDrawSurface7
'Public Surface() As DirectDrawSurface7
' describes the surfaces
Public PrimaryDesc As DDSURFACEDESC2
Public BackBufferDesc As DDSURFACEDESC2
'Public SurfaceDesc() As DDSURFACEDESC2
' what the hardware / emulator can do
Public HardWareCaps As DDCAPS
Public EmulationCaps As DDCAPS
' what the screen can do
Public EnumDisplayModes As DirectDrawEnumModes
' remembers if we're in exclusive mode
Private InExMode As Boolean
' used for showing / hiding the cursor
Private Cur As Long

' used to show / hide the cursor
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


' start directdraw
Function StartUp() As Boolean
On Error Resume Next
Set DDRAW = DX.DirectDrawCreate("")
DDRAW.GetCaps HardWareCaps, EmulationCaps
If Err.Number = DD_OK Then StartUp = True
End Function

' set how important the program is
Function SetCoop(Hwnd As Long, FullScreen As Boolean, Reboot As Boolean) As Boolean
On Error Resume Next
Dim Flags As Long
If FullScreen Then
   Flags = DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWMODEX
   If AllowReboot Then Flags = Flags & DDSCL_ALLOWREBOOT
Else
   Flags = DDSCL_NORMAL
End If
DDRAW.SetCooperativeLevel Hwnd, Flags
If Err.Number = DD_OK Or DDERR_EXCLUSIVEMODEALREADYSET Then
   InExMode = FullScreen
   SetCoop = True
End If
End Function

' set screen resolution and color depth
Function SetDisplay(Width As Integer, Height As Integer, BPP As Byte) As Boolean
On Error Resume Next
If MOD_DDRAW.CheckDisplayMode(Width, Height, BPP) = False Then Exit Function
DDRAW.SetDisplayMode Width, Height, BPP, 0, DDSDM_DEFAULT
If Err.Number = DD_OK Then SetDisplay = True
End Function

' put screen back to normal
Function RestoreDisplay() As Boolean
On Error Resume Next
DDRAW.RestoreDisplayMode
If Err.Number = DD_OK Then RestoreDisplay = True
End Function

' hides the cursor
Sub HideTheCursor()
On Error Resume Next
Cur = ShowCursor(0)
End Sub

' shows the cursor
Sub ShowTheCursor()
On Error Resume Next
If Cur Then ShowCursor Cur
End Sub

' checks if a display mode is available
Function CheckDisplayMode(Width As Integer, Height As Integer, BPP As Byte) As Boolean
On Error Resume Next
Dim SurfDesc As DDSURFACEDESC2
Set EnumDisplayModes = DDRAW.GetDisplayModesEnum(0, SurfDesc)
For i = 1 To EnumDisplayModes.GetCount()
    EnumDisplayModes.GetItem i, SurfDesc
    If SurfDesc.lWidth = Width Then
       If SurfDesc.lHeight = Height Then
          If SurfDesc.ddpfPixelFormat.lRGBBitCount = BPP Then
             CheckDisplayMode = True
          End If
       End If
    End If
Next
End Function

' creates primary and backbuffer surfaces
Function CreatePrimaryAndBackBuffer(Hwnd As Long, Flippable As Boolean, Buffers As Byte, Cap3D As Boolean, CapZbuffer As Boolean) As Boolean
On Error Resume Next
Dim SurfDesc As DDSURFACEDESC2
Dim Flags As CONST_DDSURFACEDESCFLAGS
Dim Caps As CONST_DDSURFACECAPSFLAGS
Dim r As RECT
Set Primary = Nothing
Set BackBuffer = Nothing

' primary
Caps = DDSCAPS_PRIMARYSURFACE
If Flippable Then
    Flags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    Caps = Caps Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    SurfDesc.lBackBufferCount = Buffers
End If
SurfDesc.lFlags = Flags
SurfDesc.ddsCaps.lCaps = Caps
Set Primary = DDRAW.CreateSurface(SurfDesc)
If Err.Number Then Exit Function

' backbuffer
If Flippable Then
    Caps = DDSCAPS_BACKBUFFER Or DDSCAPS_FLIP
Else
    Flags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
    Caps = DDSCAPS_OFFSCREENPLAIN
    DX.GetWindowRect Hwnd, r
    SurfDesc.lWidth = r.Right - r.Left
    SurfDesc.lHeight = r.Bottom - r.Top
End If
If Cap3D Then Caps = Caps Or DDSCAPS_3DDEVICE
If CapZbuffer Then Caps = Caps Or DDSCAPS_ZBUFFER
Set BackBuffer = DDRAW.CreateSurface(SurfDesc)
If Err.Number = DD_OK Then CreatePrimaryAndBackBuffer = True
End Function

' creates a surface from a file
Function CreateSurfFromFile(Surf As DirectDrawSurface7, FileName As String, Width As Integer, Height As Integer, Optional CK_low As Long = -1, Optional CK_high As Long = -1) As Boolean
On Error Resume Next
Dim SurfDesc As DDSURFACEDESC2
Dim ColorKey As DDCOLORKEY

With SurfDesc
   .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
   .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
   .lWidth = Width
   .lHeight = Height
End With
Set Surf = DDRAW.CreateSurfaceFromFile(FileName, SurfDesc)

If CK_low + CK_high > -1 Then
   ColorKey.low = CK_low
   ColorKey.high = CK_high
   Surf.SetColorKey DDCKEY_SRCBLT, ColorKey
End If
If Err.Number = DD_OK Then CreateSurfFromFile = True
End Function

' sets a clipper for a surface
Function CreateClipper(Surf As DirectDrawSurface7, x As Integer, y As Integer, Width As Integer, Height As Integer) As Boolean
On Error Resume Next
Dim Clipper As DirectDrawClipper
Set Clipper = DDRAW.CreateClipper(0)
Dim r(0) As RECT
r(0).Left = x
r(0).Top = y
r(0).Right = x + Width
r(0).Bottom = y + Height
Clipper.SetClipList 0, r()
Surf.SetClipper Clipper
If Err.Number = DD_OK Then CreateClipper = True
End Function

' sets a clipper for a surface from a hwnd
Function SetClipperFromHwnd(Surf As DirectDrawSurface7, Hwnd As Long) As Boolean
On Error Resume Next
Dim Clipper As DirectDrawClipper
Set Clipper = DDRAW.CreateClipper(0)
Clipper.SetHWnd Hwnd
Surf.SetClipper Clipper
If Err.Number = DD_OK Then CreateClipperfromhwnd = True
End Function

' closes down exclusive mode and unloads surfaces
Function EndIt(Hwnd As Long) As Boolean
On Error Resume Next
Set BackBuffer = Nothing
Set Primary = Nothing
If InExMode Then SetCoop Hwnd, False, True
If Err.Number = DD_OK Then EndIt = True
End Function

' turns a jpeg into a bitmap and saves it
Function JPEG2BMP(FileName As String, LoadPB As PictureBox, SavePB As PictureBox) As Boolean
On Error GoTo FileMuffUp
LoadPB = LoadPicture(FileName & ".jpg")
SavePB = LoadPB
SavePicture SavePB.Picture, FileName & ".bmp"
JPEG2BMP = True
Exit Function
FileMuffUp:
JPEG2BMP = False
End Function

