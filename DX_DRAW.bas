Attribute VB_Name = "MOD_DX_DRAW"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'            MOD_DX_DRAW.BAS - BY SIMON PRICE
'
'          LOADS OF HANDY DIRECT DRAW FUNCTIONS
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' the great grand daddy of them all
Public DX As New DirectX7
' the direct draw object, direct access to video card = cool!
Public DX_DRAW As DirectDraw7
' have we started or not?
Private InExclusiveMode As Boolean

' surfaces
Public BackBuffer As DirectDrawSurface7
Public View As DirectDrawSurface7
Public Background As DirectDrawSurface7
Public Scene As DirectDrawSurface7
Public Const NUM_TEX = 7
Public Tex(NUM_TEX) As DirectDrawSurface7
Public Const TEX_WALL = 0
Public Const TEX_SIDEWALL = 1
Public Const TEX_GRASS = 2
Public Const TEX_WATER = 3
Public Const TEX_FENCE = 4
Public Const TEX_ROOF = 5
Public Const TEX_PLANE_BOAT = 6
Public Const TEX_TREE = 7

' surface descriptions
Public SurfDesc As DDSURFACEDESC2

' back buffer capabilaties
Public BackBufferCaps As DDSCAPS2

' colour key for masking
Public ColorKey As DDCOLORKEY

' rects
Public SrcRect As RECT
Public DestRect As RECT

Public Declare Function GetInputState Lib "user32" () As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Cur As Long

Sub CrankItUp(hwnd As Long, FullScreen As Boolean)
' this sub gets it all going but creating Direct Draw
On Error GoTo TheCrappyThingDidNotEvenStartUp

' if we've already started, don't bother starting again
If InExclusiveMode Then Exit Sub

' create direct draw
Set DX_DRAW = DX.DirectDrawCreate("")

If FullScreen Then
    ' give us all the screen and all the power, yes!
    DX_DRAW.SetCooperativeLevel hwnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE 'Or DDSCL_ALLOWREBOOT
    InExclusiveMode = True
Else
    ' use normal mode
    DX_DRAW.SetCooperativeLevel hwnd, DDSCL_NORMAL
End If

Exit Sub

' if thing go pear shaped, exit
TheCrappyThingDidNotEvenStartUp:
MsgBox "Error - Cannot activate DirectX 7 - make sure you have it installed correctly!", vbExclamation, "Error!"
End
End Sub

Sub EndIt(hwnd As Long)
DX_DRAW.SetCooperativeLevel hwnd, DDSCL_NORMAL
InExMode = False
End Sub

Sub SetDisplayMode(Width As Integer, Height As Integer, Colors As Byte)
'set's the display mode to the required size and colors
 DX_DRAW.SetDisplayMode Width, Height, Colors, 0, DDSDM_DEFAULT
End Sub

Sub WaitTillOK()
Dim bRestore As Boolean

bRestore = False
Do Until ExModeActive 'short way of saying "do until it returns true"
    DoEvents 'Lets windows do other things
    bRestore = True
Loop

' if we lost and got back the surfaces, then restore them
DoEvents 'Lets windows do it's things
If bRestore Then
    bRestore = False
    DX_DRAW.RestoreAllSurfaces
    ModSurfaces.LoadAllPics ' must init the surfaces again if they we're lost. When this happens the first line of initsurfaces is important
End If
End Sub

Function ExModeActive() As Boolean
     Dim TestCoopRes As Long ' holds the return value of the test.

     TestCoopRes = DX_DRAW.TestCooperativeLevel ' Tells DDraw to do the test

     If (TestCoopRes = DD_OK) Then
         ExModeActive = True ' everything is sweet
     Else
         ExModeActive = False ' summinks gone wrong
     End If
 End Function
 
Sub CreatePrimaryWithBackBuffer()
' does what it says in the name
Set View = Nothing
Set BackBuffer = Nothing

SurfDesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
SurfDesc.lBackBufferCount = 1
Set View = DX_DRAW.CreateSurface(SurfDesc)

BackBufferCaps.lCaps = DDSCAPS_BACKBUFFER
Set BackBuffer = View.GetAttachedSurface(BackBufferCaps)
'BackBuffer.GetSurfaceDesc ViewDesc

BackBuffer.SetFontTransparency True
End Sub

Sub CreatePrimaryOnly()
' create a primary surface without a backbuffer,
' for use in normal mode
SurfDesc.lFlags = DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
Set View = DX_DRAW.CreateSurface(SurfDesc)
End Sub

Sub LoadAllSurfaces()
' loads every pic we need

If InExclusiveMode Then
    ' load primary surface and backbuffer
    CreatePrimaryWithBackBuffer
Else
    CreatePrimaryOnly
End If

'*** add app specific pics here ***

' create background
CreateSurfaceFromFile Background, SurfDesc, App.Path & "\sky.bmp", 640, 240

' set up the direct3d render target
SurfDesc.lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or DDSCAPS_SYSTEMMEMORY
' create viewport
SetRect DestRect, 0, 0, 640, 480
SurfDesc.lWidth = DestRect.Right - DestRect.LEFT
SurfDesc.lHeight = DestRect.Bottom - DestRect.Top
' create the render-target surface
Set Scene = DX_DRAW.CreateSurface(SurfDesc)
' add color key
AddColorKey Scene, vbBlack, vbBlack
' remember the dimensions of the render target
With SrcRect
    .LEFT = 0: .Top = 0
    .Bottom = SurfDesc.lHeight
    .Right = SurfDesc.lWidth
End With
'create a DirectDrawClipper and attach it to the primary surface.
'Dim Clipper As DirectDrawClipper
'Set Clipper = DX_DRAW.CreateClipper(0)
'Clipper.SetHWnd Form1.hwnd
'Scene.SetClipper Clipper

' create the z-buffer and attach to backbuffer
Dim ddpfZBuffer As DDPIXELFORMAT
Dim d3dEnumPFs As Direct3DEnumPixelFormats

Set DX_3D = DX_DRAW.GetDirect3D
Set d3dEnumPFs = DX_3D.GetEnumZBufferFormats("IID_IDirect3DRGBDevice")

Dim i As Long

For i = 1 To d3dEnumPFs.GetCount()
d3dEnumPFs.GetItem i, ddpfZBuffer
If ddpfZBuffer.lFlags = DDPF_ZBUFFER Then
  Exit For
End If
Next i

SetRect DestRect, 0, 0, 640, 480
' Prepare and create the z-buffer surface.
SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_PIXELFORMAT
SurfDesc.ddsCaps.lCaps = DDSCAPS_ZBUFFER
SurfDesc.lWidth = DestRect.Right - DestRect.LEFT
SurfDesc.lHeight = DestRect.Bottom - DestRect.Top
SurfDesc.ddpfPixelFormat = ddpfZBuffer
SurfDesc.ddsCaps.lCaps = SurfDesc.ddsCaps.lCaps Or DDSCAPS_SYSTEMMEMORY

Set ZBuff = DX_DRAW.CreateSurface(SurfDesc)

' attach the z-buffer to the back buffer
Scene.AddAttachedSurface ZBuff
End Sub

Sub UnloadSurfaces()
' remember to call this one
Set BackBuffer = Nothing
Set View = Nothing

'*** add app specific pics here ***
End Sub
 
Sub CreateSurfaceFromFile(Surface As DirectDrawSurface7, SurfDesc As DDSURFACEDESC2, FileName As String, Width As Integer, Height As Integer)
On Error GoTo LostFile
' loads a bitmap from a file and makes a pic from it
     SurfDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
     SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
     SurfDesc.lWidth = Width
     SurfDesc.lHeight = Height
     Set Surface = DX_DRAW.CreateSurfaceFromFile(FileName, SurfDesc)
Exit Sub
LostFile:
Debug.Print "File not found : " & FileName
End Sub

Sub SetRect2(Box As RECT, LEFT As Integer, Top As Integer, Right As Integer, Bottom As Integer)
' creates a rect of the required size
    Box.LEFT = LEFT
    Box.Top = Top
    Box.Right = Right
    Box.Bottom = Bottom
End Sub

Sub SetRect(Box As RECT, LEFT As Integer, Top As Integer, Width As Integer, Height As Integer)
' creates a rect of the required size
    Box.LEFT = LEFT
    Box.Top = Top
    Box.Right = LEFT + Width
    Box.Bottom = Top + Height
End Sub

Function MakeRect2(LEFT As Integer, Top As Integer, Right As Integer, Bottom As Integer) As RECT
    MakeRect2.LEFT = LEFT
    MakeRect2.Top = Top
    MakeRect2.Right = Right
    MakeRect2.Bottom = Bottom
End Function

Function MakeRect(LEFT As Integer, Top As Integer, Width As Integer, Height As Integer) As RECT
    MakeRect.LEFT = LEFT
    MakeRect.Top = Top
    MakeRect.Right = LEFT + Width
    MakeRect.Bottom = Top + Height
End Function

Sub AddColorKey(Surface As DirectDrawSurface7, low As Long, high As Long)
' for masking sprites
ColorKey.low = low
ColorKey.high = high
Surface.SetColorKey DDCKEY_SRCBLT, ColorKey
End Sub

Sub HideTheCursor()
Cur = ShowCursor(0)
End Sub

Sub ShowTheCursor()
ShowCursor Cur
End Sub

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

