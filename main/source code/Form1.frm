VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOADING... PLEASE WAIT"
   ClientHeight    =   5520
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   720
      ScaleHeight     =   204
      ScaleWidth      =   1284
      TabIndex        =   9
      Top             =   1080
      Width           =   1332
   End
   Begin VB.TextBox txtMeshHeight 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   288
      Left            =   720
      TabIndex        =   7
      Text            =   "10"
      Top             =   720
      Width           =   1332
   End
   Begin VB.TextBox txtMeshSize 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   288
      Left            =   720
      TabIndex        =   5
      Text            =   "100"
      Top             =   360
      Width           =   1332
   End
   Begin VB.CommandButton cmdLoadBitmap 
      Caption         =   "Load Bitmap ..."
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   2052
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   240
      Top             =   4440
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   "*.bmp, *.jpg"
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   2280
      ScaleHeight     =   440
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   440
      TabIndex        =   0
      Top             =   120
      Width           =   5280
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2172
      Left            =   0
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   1
      Top             =   2880
      Width           =   2292
      Begin VB.Shape MapCursor 
         BackColor       =   &H000000FF&
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   240
         Left            =   600
         Shape           =   3  'Circle
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.Label lblColor 
      Caption         =   "Color :"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2052
   End
   Begin VB.Label Label3 
      Caption         =   "Height :"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2052
   End
   Begin VB.Label lblMeshSize 
      Caption         =   "Size :"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2052
   End
   Begin VB.Label lblMesh 
      Alignment       =   2  'Center
      Caption         =   "Mesh Properties"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2052
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadBitmap 
         Caption         =   "&Load Bitmap..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRender 
      Caption         =   "&Render"
      Begin VB.Menu mnuUseSoftware 
         Caption         =   "Software Enumation"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUseHardware 
         Caption         =   "Hardware Rendering"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPoints 
         Caption         =   "&Points"
      End
      Begin VB.Menu mnuWireframe 
         Caption         =   "&Wireframe"
      End
      Begin VB.Menu mnuSolid 
         Caption         =   "&Solid"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Keys(0 To 255) As Boolean

Dim MeshSize As Single
Dim MeshDetail As Integer
Dim MeshHeight As Single

Dim LastColor As Long

Const WALKSPEED = 1
Const TURNSPEED = 0.1

Private Sub cmdLoadBitmap_Click()
On Error Resume Next
MeshHeight = Val(txtMeshHeight)
If MeshHeight < 0.1 Then
    MsgBox "Please choose a higher value for mesh height.", vbInformation, "Invalid Variable Size"
    Exit Sub
End If
If MeshHeight > 10000 Then
    MsgBox "Please choose a lower value for mesh height.", vbInformation, "Invalid Variable Size"
    Exit Sub
End If
ComDialog.ShowOpen
If ComDialog.FileName = "" Then Exit Sub
picMap = LoadPicture(ComDialog.FileName)
picMap.Left = 95 - picMap.Width / 2
MeshDetail = picMap.Width
If MeshDetail > 181 Then
    MsgBox "Please select a bitmap smaller than 181 pixels wide!", vbInformation, "Bitmap too large!"
    Exit Sub
End If
MeshSize = MeshDetail
txtMeshSize = MeshSize
MousePointer = vbHourglass
If MOD_LAND.LoadLandscape(picMap.hdc, picMap.Width, MeshSize, MeshDetail, MeshHeight, 1) = False Then
    MousePointer = vbDefault
    MsgBox "ERROR : Could not load landscape!"
Else
    MousePointer = vbDefault
    CreateMaterials
    CreateLights ' lights,
    CreateCameras ' camera,
    MainLoop ' action!
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Keys(KeyCode) = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Keys(KeyCode) = False
End Sub

Private Sub Form_Load()
Loading
ComDialog.InitDir = App.Path & "\main\bitmaps\"
Caption = "SI3D Landscape Engine - By Simon Price - visit www.VBgames.co.uk for more cool VB programs!"
End Sub

Sub Loading()
On Error Resume Next
Dim i As Integer
' show form
Show
DoEvents
' open log file
DeleteFile App.Path & "\main\log files\loading error log.txt"
Open App.Path & "\main\log files\loading error log.txt" For Output As #101
Print #101, "SI3D LANDSCAPE V1.0 LOADING ERROR LOG FILE"
Print #101, ""
' show splash screen
SplashForm.Show 0, Me
DoEvents
' load DirectDraw
lblStatus = "Loading DirectDraw 7..."
Print #101, "Loading DirectDraw 7..."
DoEvents
If MOD_DDRAW.StartUp = False Then
    Print #101, "FATAL ERROR : Could not start DirectDraw 7! Make sure DirectX 7 or higher is installed!"
    Close #101
    FatalError "Could not start DirectDraw 7! Make sure DirectX 7 or higher is installed!"
End If
DoEvents
If MOD_DDRAW.SetCoop(Hwnd, False, True) = False Then
    Print #101, "FATAL ERROR : Could not set DirectDraw cooperative level! Make sure no other programs are interfering!"
    Close #101
    FatalError "Could not set DirectDraw cooperative level! Make sure no other programs are interfering!"
End If
Print #101, "DirectDraw 7 loaded successfully"
DoEvents
' create surfaces
CreateSurfaces
DoEvents
' load Direct3D
lblStatus = "Loading Direct3D 7..."
Print #101, "Loading Direct3D 7..."
DoEvents
If MOD_D3D.StartUp = False Then
    Print #101, "FATAL ERROR : Could not start Direct3D! Make sure DirectX 7 or higher is installed!"
    Close #101
    FatalError "Could not start Direct3D! Make sure DirectX 7 or higher is installed!"
End If
DoEvents
' load Z-buffer
lblStatus = "Loading Z-Buffer..."
Print #101, "Loading Z-Buffer..."
If MOD_D3D.AttachZbuffer(BackBuffer, 16, "IID_IDirect3DRGBDevice", 450, 450) = False Then
    Print #101, "FATAL ERROR : Could not create Z-Buffer!"
    Close #101
    FatalError "Could not create Z-Buffer! Your computer does not support the required mode or bit depth!"
End If
Print #101, "Z-Buffer Loaded Successfully"
DoEvents
If MOD_D3D.StartUpDevice("IID_IDirect3DRGBDevice", BackBuffer, 450, 450) = False Then
    Print #101, "FATAL ERROR : Could not start Direct3D software rendering!"
    Close #101
    FatalError "Could not start Direct3D software rendering!"
End If
Print #101, "Direct3D 7 Loaded Successfully"
' load si3d
Print #101, "Loading SI3D Graphics Engine..."
DoEvents
MOD_SI3D.StartUp
D3DDEV.SetRenderState D3DRENDERSTATE_ZENABLE, D3DZB_TRUE
D3DDEV.SetRenderState D3DRENDERSTATE_SHADEMODE, D3DSHADE_FLAT
DoEvents
RefreshTimer.Enabled = True
Print #101, "Loading Complete..."
Close #101
DoEvents
MousePointer = vbDefault
End Sub

Sub Loading2()
On Error Resume Next
Dim i As Integer
' show form
Show
DoEvents
' open log file
DeleteFile App.Path & "\main\log files\loading error log.txt"
Open App.Path & "\main\log files\loading error log.txt" For Output As #101
Print #101, "SI3D LANDSCAPE V1.0 LOADING ERROR LOG FILE"
Print #101, ""
' show splash screen
SplashForm.Show 0, Me
DoEvents
' load DirectDraw
lblStatus = "Loading DirectDraw 7..."
Print #101, "Loading DirectDraw 7..."
DoEvents
If MOD_DDRAW.StartUp = False Then
    Print #101, "FATAL ERROR : Could not start DirectDraw 7! Make sure DirectX 7 or higher is installed!"
    Close #101
    FatalError "Could not start DirectDraw 7! Make sure DirectX 7 or higher is installed!"
End If
DoEvents
If MOD_DDRAW.SetCoop(Hwnd, False, True) = False Then
    Print #101, "FATAL ERROR : Could not set DirectDraw cooperative level! Make sure no other programs are interfering!"
    Close #101
    FatalError "Could not set DirectDraw cooperative level! Make sure no other programs are interfering!"
End If
Print #101, "DirectDraw 7 loaded successfully"
DoEvents
' create surfaces
CreateSurfaces2
DoEvents
' load Direct3D
lblStatus = "Loading Direct3D 7..."
Print #101, "Loading Direct3D 7..."
DoEvents
If MOD_D3D.StartUp = False Then
    Print #101, "FATAL ERROR : Could not start Direct3D! Make sure DirectX 7 or higher is installed!"
    Close #101
    FatalError "Could not start Direct3D! Make sure DirectX 7 or higher is installed!"
End If
DoEvents
' load Z-buffer
lblStatus = "Loading Z-Buffer..."
Print #101, "Loading Z-Buffer..."
If MOD_D3D.AttachZbuffer(BackBuffer, 16, "IID_IDirect3DRGBDevice", 450, 450, True) = False Then
    Print #101, "FATAL ERROR : Could not create Z-Buffer!"
    Close #101
    FatalError "Could not create Z-Buffer! Your computer does not support the required mode or bit depth!"
End If
Print #101, "Z-Buffer Loaded Successfully"
DoEvents
If MOD_D3D.StartUpDevice("IID_IDirect3DHALDevice", BackBuffer, 450, 450) = False Then
    Print #101, "FATAL ERROR : Could not start Direct3D hardware rendering!"
    Close #101
    FatalError "Could not start Direct3D hardware rendering!"
End If
Print #101, "Direct3D 7 Loaded Successfully"
' load si3d
Print #101, "Loading SI3D Graphics Engine..."
DoEvents
MOD_SI3D.StartUp
D3DDEV.SetRenderState D3DRENDERSTATE_ZENABLE, D3DZB_TRUE
D3DDEV.SetRenderState D3DRENDERSTATE_SHADEMODE, D3DSHADE_FLAT
DoEvents
RefreshTimer.Enabled = True
Print #101, "Loading Complete..."
Close #101
DoEvents
MousePointer = vbDefault
End Sub

Sub CreateLights()
Dim White As D3DCOLORVALUE
Dim TheSun As D3DLIGHT7
White = MOD_MATH.MakeD3DCOLORVALUE(10, 10, 10, 10)
TheSun = MOD_SI3D.MakeLight(D3DLIGHT_POINT, White, White, White, MakeVector(MeshSize / 2, MeshHeight * 5, MeshSize / 2), MakeVector(0, 0, 0), 0, 1, 0, 0, 0, 0, 1000)
MOD_SI3D.AddLight TheSun, "Sun", True
End Sub

Sub CreateCameras()
MOD_SI3D.AddCamera MOD_SI3D.MakeCamera(MOD_MATH.MakeVector(MeshSize / 2, MeshHeight + 3, MeshSize / 2), MOD_MATH.MakeVector(0, 0, 0), "Main")
MOD_SI3D.SetCamera 1
End Sub

Sub CreateMaterials()
Dim r As Single, g As Single, b As Single
MOD_MATH.Long2RGB picColor.BackColor, r, g, b
r = r / 255
g = g / 255
b = b / 255
MOD_SI3D.AddMtrl MOD_SI3D.MakeMtrl(1, r, g, b, 1, r, g, b), "Land"
End Sub

' creates the surfaces needed by the program
Sub CreateSurfaces2()
On Error Resume Next
Print #101, "Loading Primary And Backbuffer Surfaces..."
With PrimaryDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With
With BackBufferDesc
    .lFlags = DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_3DDEVICE Or DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    .lWidth = 450
    .lHeight = .lWidth
End With
DoEvents
Set MOD_DDRAW.Primary = DDRAW.CreateSurface(PrimaryDesc)
MOD_DDRAW.SetClipperFromHwnd Primary, picMain.Hwnd
DoEvents
Set MOD_DDRAW.BackBuffer = DDRAW.CreateSurface(BackBufferDesc)
Print #101, "Surfaces Loaded OK"
MousePointer = vbDefault
End Sub

' creates the surfaces needed by the program
Sub CreateSurfaces()
On Error Resume Next
Print #101, "Loading Primary And Backbuffer Surfaces..."
With PrimaryDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With
With BackBufferDesc
    .lFlags = DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_3DDEVICE Or DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    .lWidth = 450
    .lHeight = .lWidth
End With
DoEvents
Set MOD_DDRAW.Primary = DDRAW.CreateSurface(PrimaryDesc)
MOD_DDRAW.SetClipperFromHwnd Primary, picMain.Hwnd
DoEvents
Set MOD_DDRAW.BackBuffer = DDRAW.CreateSurface(BackBufferDesc)
Print #101, "Surfaces Loaded OK"
MousePointer = vbDefault
End Sub

' emergency unload, shows message and ends program
Sub FatalError(msg As String)
UnloadDX
MsgBox "ERROR : " & msg, vbCritical, "FATAL ERROR!"
Unload Me
End Sub

' unloads dx stuff
Sub UnloadDX()
On Error Resume Next
MOD_D3D.EndIt
MOD_DDRAW.RestoreDisplay
MOD_DDRAW.SetCoop Hwnd, False, True
MOD_DDRAW.EndIt Hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
UnloadDX
MsgBox "WAS THAT COOL OR WAS THAT COOL? WHATEVER YOU THINK, PLEASE REMEMBER TO VOTE FOR ME ON WWW.PLANET-SOURCE-CODE.COM !!!", vbInformation, "Thanks for trying the SI3D landscape engine!"
End
End Sub

' draws the scene in the main picturebox
Sub RenderIt()
On Error Resume Next
MOD_SI3D.Clear D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER
MOD_SI3D.RenderScene
Blt2screen
End Sub

' copies backbuffer picture to picturebox
Sub Blt2screen()
Dim DestRect As RECT
Dim SrcRect As RECT
DX.GetWindowRect picMain.Hwnd, DestRect
SrcRect = MOD_MATH.MakeRect(0, 0, 450, 450)
Primary.Blt DestRect, BackBuffer, SrcRect, DDBLT_WAIT
End Sub

Sub MainLoop()
Walk 0
Do
    If Keys(vbKeyUp) Then Walk WALKSPEED
    If Keys(vbKeyDown) Then Walk -WALKSPEED
    If Keys(vbKeyRight) Then Turn TURNSPEED
    If Keys(vbKeyLeft) Then Turn -TURNSPEED
    MOD_SI3D.SetCamera 1
    If LastColor <> picColor.BackColor Then
        ReDim MOD_SI3D.Mtrl(0)
        CreateMaterials
    End If
    LastColor = picColor.BackColor
    RenderIt
    DoEvents
Loop Until Keys(vbKeyEscape) = True
Unload Me
End Sub

Sub Walk(Speed As Single)
MOD_SI3D.TransformCamera 1, SI3D_TRANS_ZPLUS, Speed
With MOD_SI3D.Camera(1).vec
   .y = MeshHeight * MOD_MATH.GreyScale(GetPixel(picMap.hdc, .x, .Z)) + (MeshHeight * 0.15)
   MapCursor.Move .x - 10, .Z - 10
End With
End Sub

Sub Turn(Speed As Single)
MOD_SI3D.TransformCamera 1, SI3D_ROT_YPLUS, Speed
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuLoadBitmap_Click()
cmdLoadBitmap_Click
End Sub

Private Sub mnuPoints_Click()
mnuPoints.Checked = True
mnuWireframe.Checked = False
mnuSolid.Checked = False
D3DDEV.SetRenderState D3DRENDERSTATE_FILLMODE, D3DFILL_POINT
End Sub

Private Sub mnuSolid_Click()
mnuPoints.Checked = False
mnuWireframe.Checked = False
mnuSolid.Checked = True
D3DDEV.SetRenderState D3DRENDERSTATE_FILLMODE, D3DFILL_SOLID
End Sub

Private Sub mnuUseHardware_Click()
mnuUseSoftware.Checked = False
mnuUseHardware.Checked = True
UnloadDX
MousePointer = vbHourglass
Loading2
MousePointer = vbDefault
picMain.Cls
End Sub

Private Sub mnuUseSoftware_Click()
mnuUseSoftware.Checked = True
mnuUseHardware.Checked = False
UnloadDX
MousePointer = vbHourglass
Loading
MousePointer = vbDefault
picMain.Cls
End Sub

Private Sub mnuWireframe_Click()
mnuPoints.Checked = False
mnuWireframe.Checked = True
mnuSolid.Checked = False
D3DDEV.SetRenderState D3DRENDERSTATE_FILLMODE, D3DFILL_WIREFRAME
End Sub

Private Sub picColor_Click()
ComDialog.Color = picColor.BackColor
ComDialog.ShowColor
picColor.BackColor = ComDialog.Color
End Sub
