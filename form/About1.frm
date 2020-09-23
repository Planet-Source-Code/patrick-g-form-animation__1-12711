VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About VBSFC FormAnimator"
   ClientHeight    =   3000
   ClientLeft      =   4500
   ClientTop       =   3000
   ClientWidth     =   3000
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   Picture         =   "About1.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   3000
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   795
   End
   Begin VB.Timer tmrAnimate 
      Interval        =   50
      Left            =   300
      Top             =   2460
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Insane Programmers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   2895
   End
   Begin ComctlLib.ImageList Backgrounds 
      Left            =   2340
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   200
      ImageHeight     =   200
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":3177
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":D209
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":1729B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":2132D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":2B3BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":35451
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":3F4E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":49575
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":53607
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "About1.frx":5D699
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private ResultRegion As Long

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion(0, 1, 1, 0, 0), True)
    'If the above two lines are modified or moved a second copy of
    'them may be added again if the form is later Modified.
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    'If the above line is modified or moved a second copy of it
    'may be added again if the form is later Modified by VBSFC.
End Sub
Private Function CreateFormRegion(Anim As Double, ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)

'!Shaped Form Region Definition
'!3,0,55,145,200,0,0,1
    ObjectRegion = CreateEllipticRgn((55 + Anim * -55) * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, (0 + Anim * 0) * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, (145 + Anim * 55) * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, (200 + Anim * 0) * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
'!3,55,0,200,145,0,0,1
    ObjectRegion = CreateEllipticRgn((0 + Anim * 0) * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, (55 + Anim * -55) * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, (200 + Anim * 0) * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, (145 + Anim * 55) * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function

Private Sub ExampleAnimate(NumSteps As Integer, Direction As Integer)
    'Direction= 1 for forwards, -1 for backwards
    Dim iStep As Double
                '0 to 1 for forwards, 1 to 0 for backwards
    For iStep = -(Direction - 1) / 2 To (1 + Direction) / 2 Step (1 / NumSteps) 'Step from 0% animation to 100% animation in NumSteps
        DeleteObject ResultRegion
        SetWindowRgn Me.hwnd, CreateFormRegion(iStep, 1, 1, 0, 0), True
        Stop
    Next iStep
    Stop
End Sub


Private Sub tmrAnimate_Timer()
    Const NumFrames = 9
    Static iFrame As Integer, Direction As Integer
    If Direction = 0 Then Direction = 1
    
    DeleteObject ResultRegion
    SetWindowRgn Me.hwnd, CreateFormRegion(CDbl(iFrame / NumFrames), 1, 1, 0, 0), True
    Me.Picture = Backgrounds.ListImages(iFrame + 1).Picture
    iFrame = iFrame + Direction
    If iFrame < 0 Then iFrame = 0: Direction = 1
    If iFrame > NumFrames Then iFrame = NumFrames: Direction = -1
End Sub
