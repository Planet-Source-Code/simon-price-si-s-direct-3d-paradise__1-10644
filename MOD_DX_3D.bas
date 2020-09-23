Attribute VB_Name = "MOD_DX_3D"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'            MOD_DX_3D.BAS - BY SIMON PRICE
'
'        BASICS OF USING DIRECT 3D IMMEDIATE MODE
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' the main direct 3D object
Public DX_3D As Direct3D7
' created from DX_3D
Public DX3DDEV As Direct3DDevice7
' viewport
Public Viewport(0) As D3DRECT
' the rendering target
Public D3DSurf As DirectDrawSurface7
Public D3DSurfdesc As DDSURFACEDESC2
' z buffer
Public ZBuff As DirectDrawSurface7
' enum stuff
Public ddEnum As DirectDrawEnum
Public d3dEnumDevices As Direct3DEnumDevices
Public ddEnumModes As DirectDrawEnumModes
Public DriverGUID As String
Public DeviceGUID As String
Public ddsdMode As DDSURFACEDESC2
Public UsingFullScreen As Boolean
Public Using3DHardware As Boolean
Public enumInfo As DDSURFACEDESC2

' constants
Const PI As Single = 3.141592

Public Type tCamera
   x As Single
   y As Single
   z As Single
   Pitch As Single
   Rotation As Single
   Roll As Single
End Type

Public Camera As tCamera

Sub CrankItUp()
' create direct 3d
Set DX_3D = DX_DRAW.GetDirect3D

DX_DRAW.GetDisplayMode SurfDesc
If SurfDesc.ddpfPixelFormat.lRGBBitCount <= 8 Then
    MsgBox "Stop being so cheap on the colours! I only support 16 bit colour or more!"
    End
End If

' create the direct 3D device
Set DX3DDEV = DX_3D.CreateDevice("IID_IDirect3DRGBDevice", Scene)

Dim VPDesc As D3DVIEWPORT7
VPDesc.lWidth = DestRect.Right - DestRect.LEFT
VPDesc.lHeight = DestRect.Bottom - DestRect.Top
VPDesc.minz = 0#
VPDesc.maxz = 1#
DX3DDEV.SetViewport VPDesc

' remember viewport rectangle
With Viewport(0)
    .X1 = 0: .Y1 = 0
    .X2 = VPDesc.lWidth
    .Y2 = VPDesc.lHeight
End With
End Sub

Function MakeVector(x As Double, y As Double, z As Double) As D3DVECTOR
' make a vector with 3 points
Dim Vector As D3DVECTOR
With Vector
    .x = x
    .y = y
    .z = z
End With
MakeVector = Vector
End Function

Sub EnumDrivers()
Dim i As Long
' get driver info
Set ddEnum = DX.GetDDEnum()
For i = 1 To ddEnum.GetCount()
    DriverGUID = ddEnum.GetDescription(i)
Next i
End Sub

Sub EnumDevices(cmbDevice As ComboBox)
Dim i As Long
' Get device information and place device user-friendly names in a combo box.
cmbDevice.Clear
Set d3dEnumDevices = DX_3D.GetDevicesEnum()
For i = 1 To d3dEnumDevices.GetCount()
    cmbDevice.AddItem d3dEnumDevices.GetName(i)
Next
cmbDevice.ListIndex = 0
End Sub

Sub EnumModes()
Set ddEnumModes = DX_DRAW.GetDisplayModesEnum(DDEDM_DEFAULT, enumInfo)
End Sub

Sub AttachZbuffer()
'' create the z-buffer and attach to backbuffer
'Dim ddpfZBuffer As DDPIXELFORMAT
'Dim d3dEnumPFs As Direct3DEnumPixelFormats
'
'Set DX_3D = DX_DRAW.GetDirect3D
'Set d3dEnumPFs = DX_3D.GetEnumZBufferFormats("IID_IDirect3DRGBDevice")
'
'Dim i As Long
'
'For i = 1 To d3dEnumPFs.GetCount()
'Call d3dEnumPFs.GetItem(i, ddpfZBuffer)
'If ddpfZBuffer.lFlags = DDPF_ZBUFFER Then
'  Exit For
'End If
'Next i
'
'SetRect DestRect, 0, 0, 640, 480
'' Prepare and create the z-buffer surface.
'SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_PIXELFORMAT
'SurfDesc.ddsCaps.lCaps = DDSCAPS_ZBUFFER
'SurfDesc.lWidth = DestRect.Right - DestRect.LEFT
'SurfDesc.lHeight = DestRect.Bottom - DestRect.Top
'SurfDesc.ddpfPixelFormat = ddpfZBuffer
'SurfDesc.ddsCaps.lCaps = SurfDesc.ddsCaps.lCaps Or DDSCAPS_SYSTEMMEMORY
'
'Set ZBuff = DX_DRAW.CreateSurface(SurfDesc)
'
'' attach the z-buffer to the back buffer
'Scene.AddAttachedSurface ZBuff
End Sub

Public Function CreateTextureSurface(File As String) As DirectDrawSurface7
Dim ddsTexture As DirectDrawSurface7
Dim i As Long
Dim IsFound As Boolean
Dim ddsd As DDSURFACEDESC2

ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_PIXELFORMAT Or DDSD_TEXTURESTAGE

' Enumerate the texture formats, and find a device-supported texture pixel format. This
' simple tutorial is simply looking for a 16-bit texture. Real applications may be interested in
' other formats, for alpha textures, bumpmaps, etc..
Dim TextureEnum As Direct3DEnumPixelFormats
Set TextureEnum = DX3DDEV.GetTextureFormatsEnum()

For i = 1 To TextureEnum.GetCount()
    IsFound = True
    TextureEnum.GetItem i, ddsd.ddpfPixelFormat
    With ddsd.ddpfPixelFormat
        ' Skip unusual modes.
        If .lFlags And (DDPF_LUMINANCE Or DDPF_BUMPLUMINANCE Or DDPF_BUMPDUDV) Then IsFound = False
        ' Skip any FourCC formats.
        If .lFourCC <> 0 Then IsFound = False
        'Skip alpha modes.
        If .lFlags And DDPF_ALPHAPIXELS Then IsFound = False
        'We only want 16-bit formats, so skip all others.
        If .lRGBBitCount <> 16 Then IsFound = False
    End With
    If IsFound Then Exit For
Next i
' If we did not find surface support, we should exit the application.
If Not IsFound Then
    MsgBox "Unable to locate 16-bit surface support on your hardware."
    End
End If
' Turn on texture managment for the device.
ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
ddsd.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
ddsd.lTextureStage = 0
' Create a new surface for the texture.
Set ddsTexture = DX_DRAW.CreateSurfaceFromFile(File, ddsd)
' Return the newly created texture.
Set CreateTextureSurface = ddsTexture
End Function

Public Function CreateTextureSurfaceCK(File As String) As DirectDrawSurface7
Dim ddsTexture As DirectDrawSurface7
Dim i As Long
Dim IsFound As Boolean
Dim ddsd As DDSURFACEDESC2

ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_PIXELFORMAT Or DDSD_TEXTURESTAGE Or DDSD_CKSRCBLT

' Enumerate the texture formats, and find a device-supported texture pixel format. This
' simple tutorial is simply looking for a 16-bit texture. Real applications may be interested in
' other formats, for alpha textures, bumpmaps, etc..
Dim TextureEnum As Direct3DEnumPixelFormats
Set TextureEnum = DX3DDEV.GetTextureFormatsEnum()

For i = 1 To TextureEnum.GetCount()
    IsFound = True
    TextureEnum.GetItem i, ddsd.ddpfPixelFormat
    With ddsd.ddpfPixelFormat
        ' Skip unusual modes.
        If .lFlags And (DDPF_LUMINANCE Or DDPF_BUMPLUMINANCE Or DDPF_BUMPDUDV) Then IsFound = False
        ' Skip any FourCC formats.
        If .lFourCC <> 0 Then IsFound = False
        'Skip alpha modes.
        If .lFlags And DDPF_ALPHAPIXELS Then IsFound = False
        'We only want 16-bit formats, so skip all others.
        If .lRGBBitCount <> 16 Then IsFound = False
    End With
    If IsFound Then Exit For
Next i
' If we did not find surface support, we should exit the application.
If Not IsFound Then
    MsgBox "Unable to locate 16-bit surface support on your hardware."
    End
End If
' Turn on texture managment for the device.
ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
ddsd.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
ddsd.lTextureStage = 0
' Create a new surface for the texture.
Set ddsTexture = DX_DRAW.CreateSurfaceFromFile(File, ddsd)
' Return the newly created texture.
Set CreateTextureSurfaceCK = ddsTexture
End Function

Sub LoadTextures()
' *** app specific textures here ***

Set Tex(TEX_SIDEWALL) = CreateTextureSurface(App.Path & "\sidewall.bmp")
Set Tex(TEX_WALL) = CreateTextureSurface(App.Path & "\wall.bmp")
Set Tex(TEX_GRASS) = CreateTextureSurface(App.Path & "\grass.bmp")
Set Tex(TEX_WATER) = CreateTextureSurface(App.Path & "\water.bmp")

DX3DDEV.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
Set Tex(TEX_FENCE) = CreateTextureSurfaceCK(App.Path & "\fence.bmp")
MOD_DX_DRAW.AddColorKey Tex(TEX_FENCE), vbBlack, vbBlack

Set Tex(TEX_ROOF) = CreateTextureSurface(App.Path & "\roof.bmp")
Set Tex(TEX_PLANE_BOAT) = CreateTextureSurfaceCK(App.Path & "\plane_boat.bmp")
MOD_DX_DRAW.AddColorKey Tex(TEX_PLANE_BOAT), vbBlack, vbBlack
Set Tex(TEX_TREE) = CreateTextureSurfaceCK(App.Path & "\tree.bmp")
MOD_DX_DRAW.AddColorKey Tex(TEX_TREE), vbBlack, vbBlack
End Sub

Sub LoadMaterials()
' *** app specific materials

DX3DDEV.SetMaterial MakeMaterial(1, 1, 1, 1, 1, 1, 1, 1)
End Sub

Sub LoadLighting()
' *** app specific lighting

' Enable ambient lighting
DX3DDEV.SetRenderState D3DRENDERSTATE_AMBIENT, DX.CreateColorRGBA(1, 1, 1, 1)
DX3DDEV.SetRenderState D3DRENDERSTATE_LIGHTING, False
DX3DDEV.SetRenderState D3DRENDERSTATE_SHADEMODE, D3DSHADE_FLAT
End Sub

Sub LoadMatrices()
' *** app specific matrices

' Set the projection matrix. Note that the view and world matrices are set in the
' FrameMove function, so that they can be animated each frame.
Dim matProj As D3DMATRIX
DX.IdentityMatrix matProj
DX.ProjectionMatrix matProj, 1, 1000, PI / 3
DX3DDEV.SetTransform D3DTRANSFORMSTATE_PROJECTION, matProj
End Sub

Sub Load3DScene()
' *** app specific loading ***


End Sub

Sub LoadCameras()
Camera.y = 0.7
Camera.z = 5
End Sub

Sub UnloadTextures()

' *** app specific textures here ***

Dim i As Byte
For i = 1 To NUM_TEX
   Set Tex(i) = Nothing
Next
End Sub

Function MakeMaterial(Optional aa As Byte = 0, Optional ar As Byte = 0, Optional ag As Byte = 0, Optional ab As Byte = 0, Optional da As Byte = 0, Optional dr As Byte = 0, Optional dg As Byte = 0, Optional db As Byte = 0, Optional ea As Byte = 0, Optional er As Byte = 0, Optional eg As Byte = 0, Optional eb As Byte = 0, Optional sa As Byte = 0, Optional sr As Byte = 0, Optional sg As Byte = 0, Optional sb As Byte = 0, Optional p As Byte = 0) As D3DMATERIAL7
With MakeMaterial
    With .Ambient
        .a = aa
        .r = ar
        .g = ag
        .b = ab
    End With
    With .diffuse
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
    With .specular
        .a = sa
        .r = sr
        .g = sg
        .b = sb
    End With
    .power = p
End With
End Function

Sub CopyVec2Vert(srcVec As D3DVECTOR, destVert As D3DVERTEX)
destVert.x = srcVec.x
destVert.y = srcVec.y
destVert.z = srcVec.z
End Sub

Sub CopyVert2Vec(srcVert As D3DVECTOR, destVec As D3DVERTEX)
destVec.x = srcVert.x
destVec.y = srcVert.y
destVec.z = srcVert.z
End Sub

Function MakeVertex(Vec As D3DVECTOR, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As D3DVERTEX
With MakeVertex
    .x = Vec.x
    .y = Vec.y
    .z = Vec.z
    .nx = nx
    .ny = ny
    .nz = nz
    .tu = tu
    .tv = tv
End With
End Function

