VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrightContrast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Brightness or Contrast"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB 
      Height          =   270
      Left            =   1470
      TabIndex        =   11
      Top             =   4950
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtVal 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3090
      TabIndex        =   10
      Text            =   "0"
      Top             =   3840
      Width           =   585
   End
   Begin VB.PictureBox picRight 
      Height          =   2595
      Left            =   3825
      MouseIcon       =   "frmBrightContrast.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   6
      Top             =   510
      Width           =   2805
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   195
         MouseIcon       =   "frmBrightContrast.frx":030A
         MousePointer    =   99  'Custom
         ScaleHeight     =   187
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   178
         TabIndex        =   7
         Top             =   540
         Width           =   2670
      End
   End
   Begin VB.PictureBox picLeft 
      Height          =   2595
      Left            =   345
      MouseIcon       =   "frmBrightContrast.frx":0614
      MousePointer    =   99  'Custom
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   4
      Top             =   495
      Width           =   2805
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2790
         Left            =   210
         MouseIcon       =   "frmBrightContrast.frx":091E
         MousePointer    =   99  'Custom
         ScaleHeight     =   186
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   220
         TabIndex        =   5
         Top             =   480
         Width           =   3300
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Okay"
      Height          =   495
      Left            =   5115
      TabIndex        =   3
      Top             =   4365
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Contrast"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4365
      Width           =   1635
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   5
      Left            =   1905
      Max             =   100
      Min             =   -100
      TabIndex        =   1
      Top             =   3510
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Brightness"
      Height          =   495
      Left            =   2175
      TabIndex        =   0
      Top             =   4365
      Width           =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "More"
      Height          =   255
      Left            =   4605
      TabIndex        =   9
      Top             =   3840
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "Less"
      Height          =   225
      Left            =   1845
      TabIndex        =   8
      Top             =   3840
      Width           =   570
   End
End
Attribute VB_Name = "frmBrightContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim canMove As Boolean, Xs, Ys, stLeft, stTop, lastX, lastY, xDiff, yDiff
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Sub cmdOK_Click()
   Me.Picture1.Picture = Me.Picture1.Image
   MDI.ActiveForm.picImage.Picture = Me.Picture1.Picture
   MDI.ActiveForm.UpdateUndo
   MDI.ActiveForm.setDirty
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim x As Long
   Dim y As Long
   Dim r As Integer
   Dim g As Integer
   Dim b As Integer
   Dim pix As Pixel

   Picture1.Cls
   Picture1.Width = Picture2.Width
   Picture1.Height = Picture2.Height
   Screen.MousePointer = vbHourglass
   For x = 0 To Picture2.ScaleWidth
       PB.Value = (x / Picture2.ScaleWidth) * 100
       For y = 0 To Picture2.ScaleHeight
           pix = LongToPix(Picture2.Point(x, y))
           pix = Contrast(pix, -(HScroll1.Value) / 5)
           Picture1.PSet (x, y), PixToLong(pix)
       Next y
   Next x
   Screen.MousePointer = vbDefault
   Picture2.Picture = Picture2.Image
   PB.Value = 0
End Sub

Private Sub Command2_Click()
   Screen.MousePointer = 11
   Picture1.Picture = Picture2.Picture
   Bright HScroll1.Value, Picture1
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
   Picture1.Move 0, 0
   Picture2.Move 0, 0
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub HScroll1_Change()
   txtVal.Text = HScroll1.Value
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   canMove = True
   Xs = x
   Ys = y
   stLeft = Picture1.Left
   stTop = Picture1.Top
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If x = lastX And y = lastY Then Exit Sub
   If canMove Then
      Picture1.Left = stLeft - (Xs - x)
      Picture1.Top = stTop - (Ys - y)
   End If
   lastX = x
   lastY = y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   canMove = False
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   canMove = True
   Xs = x
   Ys = y
   stLeft = Picture2.Left
   stTop = Picture2.Top
   xDiff = x - Picture2.Left
   yDiff = y - Picture2.Top
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If canMove Then
      Picture2.Left = x - xDiff
      Picture2.Top = y - yDiff
   End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   canMove = False
End Sub
