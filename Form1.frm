VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4296
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5628
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4296
   ScaleWidth      =   5628
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox SavePB 
      AutoRedraw      =   -1  'True
      Height          =   1572
      Left            =   3120
      ScaleHeight     =   1524
      ScaleWidth      =   1884
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.PictureBox LoadPB 
      AutoRedraw      =   -1  'True
      Height          =   1572
      Left            =   960
      ScaleHeight     =   1524
      ScaleWidth      =   1884
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Going As Boolean
Dim FPS As Integer

Dim Rotation As Single
Dim Pitch As Single

Dim Key As Byte

Const TURN_SPEED = 0.2
Const PITCH_SPEED = 0.05
Const WALK_SPEED = 0.4

Const BOAT_SPEED = -0.006
Const PLANE_SPEED = 0.018

Const GROUND_SIZE = 8

Dim matCamera As D3DMATRIX

Dim vNwall(4) As D3DVERTEX
Dim vSwall(4) As D3DVERTEX
Dim vEwall(5) As D3DVERTEX
Dim vWwall(5) As D3DVERTEX
Dim vRoof(6) As D3DVERTEX
Dim vGround(4) As D3DVERTEX
Dim vSea(4) As D3DVERTEX
Dim vFence(10) As D3DVERTEX
Dim vPlane(4) As D3DVERTEX
Dim vBoat(4) As D3DVERTEX
Const NUM_TREES = 4
Dim VecTree(NUM_TREES) As D3DVECTOR
Dim vTree(4) As D3DVERTEX

Dim ScrollPos As Integer
Const SCROLL_SPEED = 10

Dim IsSolid(-GROUND_SIZE - 1 To GROUND_SIZE + 1, -GROUND_SIZE - 1 To GROUND_SIZE + 1) As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Key = KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Key = 0
End Sub

Private Sub Form_Load()
Randomize Timer
MousePointer = vbHourglass
Caption = "LOADING - PLEASE WAIT..."
Show
' load everything
Convert2BMPs
MOD_DX_DRAW.CrankItUp hwnd, True
MOD_DX_DRAW.SetDisplayMode 640, 480, 16
MOD_DX_DRAW.LoadAllSurfaces
MOD_DX_3D.CrankItUp
MOD_DX_3D.AttachZbuffer
MOD_DX_3D.LoadTextures
MOD_DX_3D.LoadMaterials
MOD_DX_3D.LoadLighting
MOD_DX_3D.LoadMatrices
MOD_DX_3D.LoadCameras
' enable z-buffering
DX3DDEV.SetRenderState D3DRENDERSTATE_ZENABLE, D3DZB_TRUE
' load scene
Load3DScene
' loading complete, enter main loop
MousePointer = vbDefault
MOD_DX_DRAW.HideTheCursor
Going = True
Timer1.Enabled = True
MainLoop
' unload everything
MOD_DX_3D.UnloadTextures
MOD_DX_DRAW.UnloadSurfaces
DX_DRAW.RestoreDisplayMode
MOD_DX_DRAW.ShowTheCursor
MsgBox "Don't you reckon that was cool? If you liked this example, please visit www.planet-source-code.com and vote for Simon Price's 3D Garden!", vbInformation, "VOTING TIME!"
Unload Me
End Sub

Sub MainLoop()
Dim j As Long, x As Integer
ScrollPos = 5
Do While Going
    MoveSprites
    MoveCamera
    RenderScene
    
    x = ScrollPos
    MOD_DX_DRAW.SetRect SrcRect, x, 0, 320, 240
    MOD_DX_DRAW.SetRect DestRect, 0, 0, 640, 480
    BackBuffer.Blt DestRect, Background, SrcRect, DDBLT_WAIT
    
    MOD_DX_DRAW.SetRect SrcRect, 0, 0, 640, 480
    BackBuffer.BltFast 0, 0, Scene, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    View.Flip Nothing, DDFLIP_WAIT
    
    FPS = FPS + 1
    If GetInputState Then DoEvents
Loop
End Sub

Sub Load3DScene()
' create points to construct a house from
Const NW_LOW = 0
Const NW_HIGH = 1
Const NE_LOW = 2
Const NE_HIGH = 3
Const SW_LOW = 4
Const SW_HIGH = 5
Const SE_LOW = 6
Const SE_HIGH = 7
Const W_ROOF = 8
Const E_ROOF = 9

Const WALL_HEIGHT = 1
Const ROOF_HEIGHT = 2
Const HEIGHT_RATIO = WALL_HEIGHT / ROOF_HEIGHT
Const WALL_WIDTH = 1
Const WALL_LENGTH = 2
Const TEX_WALL_WIDTH = WALL_WIDTH * 2
Const TEX_WALL_LENGTH = WALL_LENGTH * 2

Dim Vec(10) As D3DVECTOR

Vec(NW_LOW) = MOD_DX_3D.MakeVector(-WALL_LENGTH, 0, WALL_WIDTH)
Vec(NW_HIGH) = MOD_DX_3D.MakeVector(-WALL_LENGTH, WALL_HEIGHT, WALL_WIDTH)
Vec(NE_LOW) = MOD_DX_3D.MakeVector(WALL_LENGTH, 0, WALL_WIDTH)
Vec(NE_HIGH) = MOD_DX_3D.MakeVector(WALL_LENGTH, WALL_HEIGHT, WALL_WIDTH)
Vec(SW_LOW) = MOD_DX_3D.MakeVector(-WALL_LENGTH, 0, -WALL_WIDTH)
Vec(SW_HIGH) = MOD_DX_3D.MakeVector(-WALL_LENGTH, WALL_HEIGHT, -WALL_WIDTH)
Vec(SE_LOW) = MOD_DX_3D.MakeVector(WALL_LENGTH, 0, -WALL_WIDTH)
Vec(SE_HIGH) = MOD_DX_3D.MakeVector(WALL_LENGTH, WALL_HEIGHT, -WALL_WIDTH)
Vec(W_ROOF) = MOD_DX_3D.MakeVector(-WALL_LENGTH, ROOF_HEIGHT, 0)
Vec(E_ROOF) = MOD_DX_3D.MakeVector(WALL_LENGTH, ROOF_HEIGHT, 0)

' north wall
vNwall(2) = MOD_DX_3D.MakeVertex(Vec(NE_LOW), 0, 0, 0, 0, 0)
vNwall(3) = MOD_DX_3D.MakeVertex(Vec(NW_LOW), 0, 0, 0, TEX_WALL_LENGTH, 0)
vNwall(0) = MOD_DX_3D.MakeVertex(Vec(NE_HIGH), 0, 0, 0, 0, 1)
vNwall(1) = MOD_DX_3D.MakeVertex(Vec(NW_HIGH), 0, 0, 0, TEX_WALL_LENGTH, 1)

' south wall
vSwall(0) = MOD_DX_3D.MakeVertex(Vec(SE_LOW), 0, 0, 0, 0, 0)
vSwall(1) = MOD_DX_3D.MakeVertex(Vec(SW_LOW), 0, 0, 0, TEX_WALL_LENGTH, 0)
vSwall(2) = MOD_DX_3D.MakeVertex(Vec(SE_HIGH), 0, 0, 0, 0, 1)
vSwall(3) = MOD_DX_3D.MakeVertex(Vec(SW_HIGH), 0, 0, 0, TEX_WALL_LENGTH, 1)

' east wall
vEwall(0) = MOD_DX_3D.MakeVertex(Vec(NE_LOW), 0, 0, 0, 0, 1)
vEwall(1) = MOD_DX_3D.MakeVertex(Vec(SE_LOW), 0, 0, 0, 1, 1)
vEwall(2) = MOD_DX_3D.MakeVertex(Vec(NE_HIGH), 0, 0, 0, 0, HEIGHT_RATIO)
vEwall(3) = MOD_DX_3D.MakeVertex(Vec(SE_HIGH), 0, 0, 0, 1, HEIGHT_RATIO)
vEwall(4) = MOD_DX_3D.MakeVertex(Vec(E_ROOF), 0, 0, 0, 0.5, 0)

' west wall
vWwall(4) = MOD_DX_3D.MakeVertex(Vec(NW_LOW), 0, 0, 0, 0, 1)
vWwall(3) = MOD_DX_3D.MakeVertex(Vec(SW_LOW), 0, 0, 0, 1, 1)
vWwall(2) = MOD_DX_3D.MakeVertex(Vec(NW_HIGH), 0, 0, 0, 0, HEIGHT_RATIO)
vWwall(1) = MOD_DX_3D.MakeVertex(Vec(SW_HIGH), 0, 0, 0, 1, HEIGHT_RATIO)
vWwall(0) = MOD_DX_3D.MakeVertex(Vec(W_ROOF), 0, 0, 0, 0.5, 0)

' roof
vRoof(1) = MOD_DX_3D.MakeVertex(Vec(SW_HIGH), 0, 0, 0, 0, 0)
vRoof(0) = MOD_DX_3D.MakeVertex(Vec(SE_HIGH), 0, 0, 0, TEX_WALL_LENGTH, 0)
vRoof(3) = MOD_DX_3D.MakeVertex(Vec(W_ROOF), 0, 0, 0, 0, 1)
vRoof(2) = MOD_DX_3D.MakeVertex(Vec(E_ROOF), 0, 0, 0, TEX_WALL_LENGTH, 1)
vRoof(5) = MOD_DX_3D.MakeVertex(Vec(NW_HIGH), 0, 0, 0, 0, 0)
vRoof(4) = MOD_DX_3D.MakeVertex(Vec(NE_HIGH), 0, 0, 0, TEX_WALL_LENGTH, 0)

' ground
Const NW_CORNER = 0
Const NE_CORNER = 1
Const SW_CORNER = 2
Const SE_CORNER = 3

Const GROUND_SIZE2 = GROUND_SIZE

Vec(NW_CORNER) = MOD_DX_3D.MakeVector(-GROUND_SIZE, 0, GROUND_SIZE)
Vec(NE_CORNER) = MOD_DX_3D.MakeVector(GROUND_SIZE, 0, GROUND_SIZE)
Vec(SW_CORNER) = MOD_DX_3D.MakeVector(-GROUND_SIZE, 0, -GROUND_SIZE)
Vec(SE_CORNER) = MOD_DX_3D.MakeVector(GROUND_SIZE, 0, -GROUND_SIZE)

vGround(NW_CORNER) = MOD_DX_3D.MakeVertex(Vec(NW_CORNER), 0, 0, 0, 0, 0)
vGround(NE_CORNER) = MOD_DX_3D.MakeVertex(Vec(NE_CORNER), 0, 0, 0, 4, 0)
vGround(SW_CORNER) = MOD_DX_3D.MakeVertex(Vec(SW_CORNER), 0, 0, 0, 0, 4)
vGround(SE_CORNER) = MOD_DX_3D.MakeVertex(Vec(SE_CORNER), 0, 0, 0, 4, 4)

' fence
Const FENCE_HEIGHT = 0.5
Const TEX_FENCE_WIDTH = GROUND_SIZE * 2

Const NW_LOW2 = 8
Const NW_HIGH2 = 9
Const SW_LOW2 = 6
Const SW_HIGH2 = 7
Const SE_LOW2 = 4
Const SE_HIGH2 = 5

Vec(NW_LOW) = MOD_DX_3D.MakeVector(-GROUND_SIZE, 0, GROUND_SIZE)
Vec(NE_LOW) = MOD_DX_3D.MakeVector(GROUND_SIZE, 0, GROUND_SIZE)
Vec(SW_LOW) = MOD_DX_3D.MakeVector(-GROUND_SIZE, 0, -GROUND_SIZE)
Vec(SE_LOW) = MOD_DX_3D.MakeVector(GROUND_SIZE, 0, -GROUND_SIZE)
Vec(NW_HIGH) = MOD_DX_3D.MakeVector(-GROUND_SIZE, FENCE_HEIGHT, GROUND_SIZE)
Vec(NE_HIGH) = MOD_DX_3D.MakeVector(GROUND_SIZE, FENCE_HEIGHT, GROUND_SIZE)
Vec(SW_HIGH) = MOD_DX_3D.MakeVector(-GROUND_SIZE, FENCE_HEIGHT, -GROUND_SIZE)
Vec(SE_HIGH) = MOD_DX_3D.MakeVector(GROUND_SIZE, FENCE_HEIGHT, -GROUND_SIZE)

vFence(NW_LOW) = MOD_DX_3D.MakeVertex(Vec(NW_LOW), 0, 0, 0, 0, 0)
vFence(NW_HIGH) = MOD_DX_3D.MakeVertex(Vec(NW_HIGH), 0, 0, 0, 0, 1)
vFence(NE_LOW) = MOD_DX_3D.MakeVertex(Vec(NE_LOW), 0, 0, 0, TEX_FENCE_WIDTH, 0)
vFence(NE_HIGH) = MOD_DX_3D.MakeVertex(Vec(NE_HIGH), 0, 0, 0, TEX_FENCE_WIDTH, 1)
vFence(SE_LOW2) = MOD_DX_3D.MakeVertex(Vec(SE_LOW), 0, 0, 0, 0, 0)
vFence(SE_HIGH2) = MOD_DX_3D.MakeVertex(Vec(SE_HIGH), 0, 0, 0, 0, 1)
vFence(SW_LOW2) = MOD_DX_3D.MakeVertex(Vec(SW_LOW), 0, 0, 0, TEX_FENCE_WIDTH, 0)
vFence(SW_HIGH2) = MOD_DX_3D.MakeVertex(Vec(SW_HIGH), 0, 0, 0, TEX_FENCE_WIDTH, 1)
vFence(NW_LOW2) = MOD_DX_3D.MakeVertex(Vec(NW_LOW), 0, 0, 0, 0, 0)
vFence(NW_HIGH2) = MOD_DX_3D.MakeVertex(Vec(NW_HIGH), 0, 0, 0, 0, 1)

' sea
Const SEA_SIZE = 40
Const SEA_LEVEL = -0.5

Vec(NW_CORNER) = MOD_DX_3D.MakeVector(-SEA_SIZE, SEA_LEVEL, SEA_SIZE)
Vec(NE_CORNER) = MOD_DX_3D.MakeVector(SEA_SIZE, SEA_LEVEL, SEA_SIZE)
Vec(SW_CORNER) = MOD_DX_3D.MakeVector(-SEA_SIZE, SEA_LEVEL, -SEA_SIZE)
Vec(SE_CORNER) = MOD_DX_3D.MakeVector(SEA_SIZE, SEA_LEVEL, -SEA_SIZE)

vSea(NW_CORNER) = MOD_DX_3D.MakeVertex(Vec(NW_CORNER), 0, 0, 0, 0, 0)
vSea(NE_CORNER) = MOD_DX_3D.MakeVertex(Vec(NE_CORNER), 0, 0, 0, SEA_SIZE / 2, 0)
vSea(SW_CORNER) = MOD_DX_3D.MakeVertex(Vec(SW_CORNER), 0, 0, 0, 0, SEA_SIZE / 2)
vSea(SE_CORNER) = MOD_DX_3D.MakeVertex(Vec(SE_CORNER), 0, 0, 0, SEA_SIZE / 2, SEA_SIZE / 2)

' boat
Const BOAT_WIDTH = 1
Const BOAT_HEIGHT = 1.8
Const BOAT_DISTANCE = 12.5
DX.CreateD3DVertex -BOAT_WIDTH, SEA_LEVEL, -BOAT_DISTANCE, 0, 0, 0, 0, 1, vBoat(0)
DX.CreateD3DVertex BOAT_WIDTH, SEA_LEVEL, -BOAT_DISTANCE, 0, 0, 0, 1, 1, vBoat(1)
DX.CreateD3DVertex -BOAT_WIDTH, SEA_LEVEL + BOAT_HEIGHT, -BOAT_DISTANCE, 0, 0, 0, 0, 0.5, vBoat(2)
DX.CreateD3DVertex BOAT_WIDTH, SEA_LEVEL + BOAT_HEIGHT, -BOAT_DISTANCE, 0, 0, 0, 1, 0.5, vBoat(3)

' plane
Const PLANE_WIDTH = 1
Const PLANE_HEIGHT = 2
Const PLANE_ALT = 5
Const PLANE_DISTANCE = 12.5
DX.CreateD3DVertex -PLANE_WIDTH, PLANE_ALT, -PLANE_DISTANCE, 0, 0, 0, 1, 0.5, vPlane(0)
DX.CreateD3DVertex PLANE_WIDTH, PLANE_ALT, -PLANE_DISTANCE, 0, 0, 0, 0, 0.5, vPlane(1)
DX.CreateD3DVertex -PLANE_WIDTH, PLANE_ALT + PLANE_HEIGHT, -PLANE_DISTANCE, 0, 0, 0, 1, 0, vPlane(2)
DX.CreateD3DVertex PLANE_WIDTH, PLANE_ALT + PLANE_HEIGHT, -PLANE_DISTANCE, 0, 0, 0, 0, 0, vPlane(3)

' trees
Const TREE_WIDTH = 1
Const TREE_HEIGHT = 1
'Dim tVec(NUM_TREES) As D3DVECTOR
VecTree(0) = MOD_DX_3D.MakeVector(5, 0, 6)
VecTree(1) = MOD_DX_3D.MakeVector(4, 0, -3)
VecTree(2) = MOD_DX_3D.MakeVector(0, 0, -5)
VecTree(3) = MOD_DX_3D.MakeVector(-4, 0, 6)
vTree(0).tu = 0: vTree(0).tv = 0
vTree(1).tu = 1: vTree(1).tv = 0
vTree(2).tu = 0: vTree(2).tv = 1
vTree(3).tu = 1: vTree(3).tv = 1
'For i = 0 To NUM_TREES - 1
'    DX.CreateD3DVertex tVec(i).x, tVec(i).y, tVec(i).z, 0, 0, 0, 0, 1, vTree(i * 4)
'    DX.CreateD3DVertex tVec(i).x + TREE_WIDTH, tVec(i).y, tVec(i).z, 0, 0, 0, 1, 1, vTree(i * 4 + 1)
'    DX.CreateD3DVertex tVec(i).x, tVec(i).y + TREE_HEIGHT, tVec(i).z, 0, 0, 0, 0, 0, vTree(i * 4 + 2)
'    DX.CreateD3DVertex tVec(i).x + TREE_WIDTH, tVec(i).y + TREE_HEIGHT, tVec(i).z, 0, 0, 0, 1, 0, vTree(i * 4 + 3)
'Next

' collision detection
Dim x As Integer, y As Integer
For x = -GROUND_SIZE - 1 To GROUND_SIZE + 1
    IsSolid(x, -GROUND_SIZE - 1) = True
    IsSolid(x, GROUND_SIZE + 1) = True
Next
For y = -GROUND_SIZE - 1 To GROUND_SIZE + 1
    IsSolid(-GROUND_SIZE - 1, y) = True
    IsSolid(GROUND_SIZE + 1, y) = True
Next
For x = -WALL_LENGTH To WALL_LENGTH
For y = -WALL_WIDTH To WALL_WIDTH
    IsSolid(x, y) = True
Next
Next
End Sub

Sub MoveCamera()
Dim BackedUp As Boolean
BackUp:
Select Case Key
   Case vbKeyLeft
       Rotation = Rotation - TURN_SPEED
       ScrollPos = ScrollPos - SCROLL_SPEED
       If ScrollPos < 0 Then ScrollPos = ScrollPos + 320
   Case vbKeyRight
       Rotation = Rotation + TURN_SPEED
       ScrollPos = ScrollPos + SCROLL_SPEED
       If ScrollPos > 320 Then ScrollPos = ScrollPos - 320
   Case vbKeyUp
       Camera.z = Camera.z - Cos(Rotation) * WALK_SPEED
       Camera.x = Camera.x - Sin(Rotation) * WALK_SPEED
       If IsSolid(Camera.x, Camera.z) Then
           Key = vbKeyDown
           BackedUp = True
           GoTo BackUp
       End If
   Case vbKeyDown
       Camera.z = Camera.z + Cos(Rotation) * WALK_SPEED
       Camera.x = Camera.x + Sin(Rotation) * WALK_SPEED
       If IsSolid(Camera.x, Camera.z) Then
           Key = vbKeyUp
           BackedUp = True
           GoTo BackUp
       End If
   Case vbKeyPageUp
       Pitch = Pitch - PITCH_SPEED
   Case vbKeyPageDown
       Pitch = Pitch + PITCH_SPEED
End Select

If BackedUp Then
    BackedUp = False
    GoTo BackUp
End If

' camera movement
Dim matView As D3DMATRIX
Dim matRotation As D3DMATRIX
Dim matPitch As D3DMATRIX
Dim matLook As D3DMATRIX
Dim matPos As D3DMATRIX

DX.IdentityMatrix matView
DX.IdentityMatrix matPos
DX.IdentityMatrix matRotation
DX.RotateYMatrix matRotation, Rotation
DX.RotateXMatrix matPitch, Pitch
DX.MatrixMultiply matLook, matRotation, matPitch
matPos.rc41 = Camera.x
matPos.rc42 = -Camera.y
matPos.rc43 = Camera.z
DX.MatrixMultiply matView, matPos, matLook
DX3DDEV.SetTransform D3DTRANSFORMSTATE_VIEW, matView

End Sub

Sub MoveSprites()
Dim i As Byte
Dim matBoatRot As D3DMATRIX
Dim matNewBoat As D3DMATRIX
Dim matOldBoat As D3DMATRIX
Dim matPlaneRot As D3DMATRIX
Dim matNewPlane As D3DMATRIX
Dim matOldPlane As D3DMATRIX

' move boat + plane
DX.IdentityMatrix matBoatRot
DX.IdentityMatrix matOldBoat
DX.RotateYMatrix matBoatRot, BOAT_SPEED
DX.IdentityMatrix matPlaneRot
DX.IdentityMatrix matOldPlane
DX.RotateYMatrix matPlaneRot, PLANE_SPEED

For i = 0 To 3
     matOldBoat.rc41 = vBoat(i).x
     matOldBoat.rc42 = vBoat(i).y
     matOldBoat.rc43 = vBoat(i).z
     matOldPlane.rc41 = vPlane(i).x
     matOldPlane.rc42 = vPlane(i).y
     matOldPlane.rc43 = vPlane(i).z
     
     DX.MatrixMultiply matNewBoat, matOldBoat, matBoatRot
     DX.MatrixMultiply matNewPlane, matOldPlane, matPlaneRot
     
     vBoat(i).x = matNewBoat.rc41
     vBoat(i).y = matNewBoat.rc42
     vBoat(i).z = matNewBoat.rc43
     vPlane(i).x = matNewPlane.rc41
     vPlane(i).y = matNewPlane.rc42
     vPlane(i).z = matNewPlane.rc43
Next
End Sub

Sub RenderScene() ' draws everyfink
Dim i As Byte

' clear background with sky picture
DX3DDEV.Clear 1, Viewport(), D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1, 0

DX3DDEV.BeginScene

' draw walls
DX3DDEV.SetTexture 0, Tex(TEX_WALL)
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vNwall(0), 4, D3DDP_DEFAULT
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vSwall(0), 4, D3DDP_DEFAULT
' side walls
DX3DDEV.SetTexture 0, Tex(TEX_SIDEWALL)
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vEwall(0), 5, D3DDP_DEFAULT
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vWwall(0), 5, D3DDP_DEFAULT
' draw roof
DX3DDEV.SetTexture 0, Tex(TEX_ROOF)
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vRoof(0), 6, D3DDP_DEFAULT

' draw trees
DX3DDEV.SetTexture 0, Tex(TEX_TREE)
Dim Rot As Single
Rot = -Rotation
For i = 0 To NUM_TREES - 1
   vTree(0).x = VecTree(i).x - Cos(Rot)
   vTree(0).z = VecTree(i).z - Sin(Rot)
   vTree(2).x = VecTree(i).x - Cos(Rot)
   vTree(2).z = VecTree(i).z - Sin(Rot)
   vTree(1).x = VecTree(i).x + Cos(Rot)
   vTree(1).z = VecTree(i).z + Sin(Rot)
   vTree(3).x = VecTree(i).x + Cos(Rot)
   vTree(3).z = VecTree(i).z + Sin(Rot)
   vTree(0).y = 2
   vTree(1).y = 2
   vTree(2).y = 0
   vTree(3).y = 0
   DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vTree(0), 4, D3DDP_DEFAULT
Next

' draw plane
DX3DDEV.SetTexture 0, Tex(TEX_PLANE_BOAT)
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vPlane(0), 4, D3DDP_DEFAULT
' draw fence
DX3DDEV.SetTexture 0, Tex(TEX_FENCE)
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vFence(0), 10, D3DDP_DEFAULT
' draw boat
DX3DDEV.SetTexture 0, Tex(TEX_PLANE_BOAT)
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vBoat(0), 4, D3DDP_DEFAULT
' draw floor
DX3DDEV.SetTexture 0, Tex(TEX_GRASS)
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vGround(0), 4, D3DDP_DEFAULT
' draw sea
DX3DDEV.SetTexture 0, Tex(TEX_WATER)
DX3DDEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, vSea(0), 4, D3DDP_DEFAULT

DX3DDEV.EndScene
End Sub

Private Sub Form_Unload(Cancel As Integer)
Going = False
End Sub

Private Sub Timer1_Timer()
Debug.Print FPS
FPS = 0
End Sub

Sub Convert2BMPs()
If MOD_DX_DRAW.JPEG2BMP(App.Path & "\grass", LoadPB, SavePB) = False Then FilesMissing App.Path & "\grass"
If MOD_DX_DRAW.JPEG2BMP(App.Path & "\sky", LoadPB, SavePB) = False Then FilesMissing App.Path & "\sky"
If MOD_DX_DRAW.JPEG2BMP(App.Path & "\wall", LoadPB, SavePB) = False Then FilesMissing App.Path & "\wall"
If MOD_DX_DRAW.JPEG2BMP(App.Path & "\sidewall", LoadPB, SavePB) = False Then FilesMissing App.Path & "\sidewall"
If MOD_DX_DRAW.JPEG2BMP(App.Path & "\fence", LoadPB, SavePB) = False Then FilesMissing App.Path & "\fence"
If MOD_DX_DRAW.JPEG2BMP(App.Path & "\water", LoadPB, SavePB) = False Then FilesMissing App.Path & "\water"
If MOD_DX_DRAW.JPEG2BMP(App.Path & "\roof", LoadPB, SavePB) = False Then FilesMissing App.Path & "\roof"
If MOD_DX_DRAW.JPEG2BMP(App.Path & "\plane_boat", LoadPB, SavePB) = False Then FilesMissing App.Path & "\roof"
End Sub

Sub FilesMissing(FileName As String)
MsgBox "ERROR : Picture File Missing - Cannot Load " & FileName & ".jpg - Ending Program!!!", vbExclamation, "ERROR - FILE NOT FOUND"
End
End Sub
