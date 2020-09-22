VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmImage 
   Caption         =   "Image"
   ClientHeight    =   4695
   ClientLeft      =   2505
   ClientTop       =   2010
   ClientWidth     =   5310
   Icon            =   "Image.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   Tag             =   "picImage"
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1215
      Top             =   3765
   End
   Begin MSComDlg.CommonDialog ComDia 
      Left            =   555
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar vbarScroller 
      Height          =   2985
      Left            =   3270
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "bar"
      Top             =   165
      Width           =   240
   End
   Begin VB.HScrollBar hbarScroller 
      Height          =   240
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "bar"
      Top             =   3330
      Width           =   2520
   End
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   30
      MousePointer    =   99  'Custom
      ScaleHeight     =   222
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   202
      TabIndex        =   0
      Top             =   -60
      Width           =   3030
      Begin VB.PictureBox picOrig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   2415
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   15
         ScaleHeight     =   585
         ScaleWidth      =   735
         TabIndex        =   5
         Top             =   330
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox picDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   2430
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   4
         Top             =   2700
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picImage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   885
         MousePointer    =   99  'Custom
         ScaleHeight     =   175
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   1
         Tag             =   "picImage"
         Top             =   345
         Width           =   1515
         Begin PicClip.PictureClip picClip 
            Left            =   75
            Top             =   720
            _ExtentX        =   635
            _ExtentY        =   688
            _Version        =   393216
         End
         Begin VB.Shape shFr 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   330
            Left            =   105
            Top             =   2235
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Shape shRed 
            BorderColor     =   &H000000FF&
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H000000FF&
            FillStyle       =   5  'Downward Diagonal
            Height          =   405
            Left            =   870
            Shape           =   1  'Square
            Top             =   1455
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Shape shSquare 
            DrawMode        =   6  'Mask Pen Not
            Height          =   375
            Left            =   990
            Shape           =   1  'Square
            Top             =   990
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Shape shCircle 
            DrawMode        =   6  'Mask Pen Not
            Height          =   375
            Left            =   975
            Shape           =   3  'Circle
            Top             =   375
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Shape shRect 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   330
            Left            =   615
            Top             =   2010
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Shape Shape1 
            Height          =   315
            Left            =   45
            Shape           =   1  'Square
            Top             =   240
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Line Line1 
            Visible         =   0   'False
            X1              =   8
            X2              =   74
            Y1              =   8
            Y2              =   8
         End
      End
   End
   Begin MSComctlLib.ImageList ImageListCursors 
      Left            =   2325
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Image.frx":0CCE
            Key             =   "text"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Image.frx":0E32
            Key             =   "magnify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Image.frx":0F96
            Key             =   "dropper"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Image.frx":10FA
            Key             =   "flood"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Image.frx":125E
            Key             =   "pencil"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Image.frx":13C2
            Key             =   "eraser"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Image.frx":1526
            Key             =   "spray"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Image.frx":168A
            Key             =   "hand"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteFile 
         Caption         =   "&Delete File"
      End
      Begin VB.Menu mnuSize 
         Caption         =   "File &Info"
      End
      Begin VB.Menu sepone 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Begin VB.Menu mnuPasteNewImage 
            Caption         =   "As &New Image"
         End
         Begin VB.Menu mnuPasteNewSelection 
            Caption         =   "As New &Selection"
            Shortcut        =   ^V
         End
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResizeImage 
         Caption         =   "&Resize Image"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear Clipboard"
      End
   End
   Begin VB.Menu mnuFit 
      Caption         =   "Fit to Im&age"
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "Effec&ts"
      Begin VB.Menu mnuHorz 
         Caption         =   "Flip Horizonally (Mirror)"
      End
      Begin VB.Menu mnuVert 
         Caption         =   "Flip Vertically"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu90 
         Caption         =   "Rotate Right 90 Degrees"
      End
      Begin VB.Menu mnu180 
         Caption         =   "Rotate 180 Degrees"
      End
      Begin VB.Menu mnu270 
         Caption         =   "Rotate Left 90 Degrees"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBright 
         Caption         =   "Adjust Brightness"
      End
      Begin VB.Menu mnuConstrast 
         Caption         =   "Adjust Contrast"
      End
      Begin VB.Menu mnuSoften 
         Caption         =   "&Soften"
      End
      Begin VB.Menu mnuClone 
         Caption         =   "&Clone"
      End
      Begin VB.Menu mnuGlitter 
         Caption         =   "G&litter"
      End
      Begin VB.Menu mnuGray 
         Caption         =   "&Grayscale"
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "&Selection"
      Begin VB.Menu mnuFloat 
         Caption         =   "&Float Selection"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuNone 
         Caption         =   "Select &None"
      End
   End
   Begin VB.Menu mnuLayer 
      Caption         =   "Layer"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Windows"
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuTileHorz 
         Caption         =   "Tile Horizonally"
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu septwo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseWindows 
         Caption         =   "Close All Open Windows"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPopCancelSelect 
         Caption         =   "Cancel &Select"
      End
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents cGif As GIF
Attribute cGif.VB_VarHelpID = -1
Dim l, t, w, h, lDiff, tDiff, fTime As Boolean 'for clone
Dim FileSize, suffit As String
Dim X1, Y1, cr, mag As Double
Dim isMagged As Boolean
Dim handId, drawID, eraseID, textID, floodID, sprayID, pickcolorID, magnifyID
Dim xStart, yStart, XLo, YLo, XHi, YHi, XOff, YOff 'for Ants Rectangle
Dim ctx, cty, ctx1, cty1, ctx2, cty2
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" _
    (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib _
    "gdi32" (ByVal hDc As Long, _
    ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0
Private vItem As Variant
Public UniqueNum&
Public Dirty As Boolean
Public ColUndo As New Collection
Public ColRedo As New Collection

Private Sub Form_Activate()
   MDI.ActiveForm.setHorzButtons
End Sub

Private Sub Form_Load()
   Dim i
   For i = 1 To ImageListCursors.ListImages.count
       If ImageListCursors.ListImages.Item(i).Key = "pencil" Then drawID = i
       If ImageListCursors.ListImages.Item(i).Key = "eraser" Then eraseID = i
       If ImageListCursors.ListImages.Item(i).Key = "text" Then textID = i
       If ImageListCursors.ListImages.Item(i).Key = "flood" Then floodID = i
       If ImageListCursors.ListImages.Item(i).Key = "spray" Then sprayID = i
       If ImageListCursors.ListImages.Item(i).Key = "dropper" Then pickcolorID = i
       If ImageListCursors.ListImages.Item(i).Key = "magnify" Then magnifyID = i
       If ImageListCursors.ListImages.Item(i).Key = "hand" Then handId = i
   Next i
   isMagged = False
   FileSize = TFileSize
   suffit = Tsuffit
End Sub

Private Sub Form_Terminate()
   Dim i
   If MDI.ActiveForm Is Nothing Then
      For i = 1 To MDI.tbHorz.Buttons.count
          If MDI.tbHorz.Buttons(i).Tag = "all" Then
             MDI.tbHorz.Buttons(i).Visible = True
          Else
             MDI.tbHorz.Buttons(i).Visible = False
          End If
      Next i
   End If
End Sub

Private Sub mnu180_Click()
   Screen.MousePointer = 11
   DoEvents
   Call rotatePic(Me.picImage, Me.picTemp, 180)
   DoEvents
   Me.picTemp.Picture = Me.picTemp.Image
   Me.picImage.Width = Me.picTemp.Width
   Me.picImage.ScaleWidth = Me.picTemp.ScaleWidth
   Me.picImage.Height = Me.picTemp.Height
   Me.picImage.ScaleHeight = Me.picTemp.ScaleHeight
   Me.picImage.Picture = LoadPicture()
   Me.picImage.PaintPicture Me.picTemp.Picture, 0, 0
   Call Me.FitToImage
   Dirty = True
   Me.UpdateUndo
   Screen.MousePointer = 0
End Sub

Private Sub mnu270_Click()
   Screen.MousePointer = 11
   DoEvents
   Call rotatePic(Me.picImage, Me.picTemp, 270)
   DoEvents
   Me.picTemp.Picture = Me.picTemp.Image
   Me.picImage.Width = Me.picTemp.Width
   Me.picImage.ScaleWidth = Me.picTemp.ScaleWidth
   Me.picImage.Height = Me.picTemp.Height
   Me.picImage.ScaleHeight = Me.picTemp.ScaleHeight
   Me.picImage.Picture = LoadPicture()
   Me.picImage.PaintPicture Me.picTemp.Picture, 0, 0
   Call Me.FitToImage
   Dirty = True
   Me.UpdateUndo
   Screen.MousePointer = 0
End Sub

Private Sub mnu90_Click()
   Screen.MousePointer = 11
   DoEvents
   Call rotatePic(Me.picImage, Me.picTemp, 90)
   DoEvents
   Me.picTemp.Picture = Me.picTemp.Image
   Me.picImage.Width = Me.picTemp.Width
   Me.picImage.ScaleWidth = Me.picTemp.ScaleWidth
   Me.picImage.Height = Me.picTemp.Height
   Me.picImage.ScaleHeight = Me.picTemp.ScaleHeight
   Me.picImage.Picture = LoadPicture()
   Me.picImage.PaintPicture Me.picTemp.Picture, 0, 0
   Call Me.FitToImage
   Dirty = True
   Me.UpdateUndo
   Screen.MousePointer = 0
End Sub

Private Sub mnuAll_Click()
   FileSelectAll
End Sub

Public Sub FileSelectAll()
   shRect.Move 0, 0, picImage.Width, picImage.Height
   shRect.Visible = True
   Timer1.Enabled = True
   picClip.Picture = picImage.Picture
   picClip.ClipX = shRect.Left
   picClip.ClipY = shRect.Top
   picClip.ClipWidth = shRect.Width
   picClip.ClipHeight = shRect.Height
   MoveIt = True
   setHorzButtons
End Sub

Private Sub mnuBright_Click()
   frmBrightContrast.Picture2.Picture = Me.picImage.Picture
   frmBrightContrast.Show
End Sub

Private Sub mnuCascade_Click()
   MDI.Arrange 0
   Dim chkme As Form
   For Each chkme In Forms
       If chkme.Name = "frmImage" Then
          Dim ctrl
          For Each ctrl In chkme.Controls
              If ctrl.Name = "picImage" Then
                 ctrl.Left = 0
                 ctrl.Top = 0
              End If
          Next ctrl
       End If
   Next chkme
End Sub

Private Sub mnuClear_Click()
   If Clipboard.GetFormat(vbCFBitmap) Then
      Clipboard.Clear
   End If
   Call setHorzButtons
End Sub

Private Sub mnuClone_Click()
   setSwitchesFalse
   RetouchIt = True
   setPointer
   MDI.cboRetouch.Visible = True
   MDI.cboRetouch.Text = "Clone"
   shCircle.Visible = True
   MDI.Text1.Visible = False
   MDI.Text2.Visible = False
   MDI.txtS.Visible = False
   MDI.UpDownS.Visible = False
   MDI.txtP.Visible = False
   MDI.UpDownP.Visible = False
   MDI.tbVert.Buttons.Item(13).Value = tbrPressed
End Sub

Private Sub mnuConstrast_Click()
   mnuBright_Click
End Sub

Private Sub mnuCopy_Click()
   Call FileCopy
End Sub

Public Sub FileCopy()
   Clipboard.Clear
   Clipboard.SetData picClip.Clip
   picImage.MouseIcon = ImageListCursors.ListImages(5).Picture 'pencil
   Call setHorzButtons
End Sub

Public Sub FileCrop()
   Call FileCopy
   picImage.Picture = LoadPicture()
   picImage.Width = picClip.ClipWidth
   picImage.Height = picClip.ClipHeight
   Call FitToImage
   Call FilePasteNewSelection
   MoveIt = False
   canMove = False
   PasteIt = False
   UpdateUndo
   Dirty = True
   shRect.Visible = False
   Timer1.Enabled = False
   setHorzButtons
   picImage.MouseIcon = ImageListCursors.ListImages(5).Picture
End Sub

Private Sub mnuCut_Click()
   Call FileCut
End Sub

Public Sub FileCut()
   Clipboard.Clear
   Clipboard.SetData picClip.Clip
   mnuDelete_Click
   Call setHorzButtons
End Sub

Private Sub mnuDelete_Click()
   If shRect.Visible = True Then
      picImage.MousePointer = 11 'hourglass
      picDelete.Width = shRect.Width
      picDelete.Height = shRect.Height
      picDelete.Picture = picDelete.Image
      picImage.PaintPicture picDelete.Picture, shRect.Left, shRect.Top
      picImage.Picture = picImage.Image
      UpdateUndo
      Dirty = True
      picImage.MousePointer = 99 'custom
      shRect.Visible = False
      Timer1.Enabled = False
      setHorzButtons
      picImage.MouseIcon = ImageListCursors.ListImages(5).Picture 'pencil
      MoveIt = False
   End If
End Sub

Private Sub mnuDeleteFile_Click()
   Dim ret
   ret = MsgBox("Okay to delete from hard drive, the file: " & Me.Caption, vbYesNo)
   If ret = vbYes Then
      deleteFile = Me.Caption
      Kill Me.Caption
      ' See if Browse is open and refresh File1.
      Dim chkme As Form
      For Each chkme In Forms
          ret = Mid(chkme.Caption, 1, 6)
          If ret = "Browse" Then
             frmBrowse.File1.Refresh
             Call frmBrowse.deleteOne
             Exit For
          End If
      Next chkme
      ' ---------------------------------------
   End If
   Unload Me
End Sub

Private Sub mnuEdit_Click()
   mnuCut.Enabled = shRect.Visible
   mnuCopy.Enabled = shRect.Visible
   mnuDelete.Enabled = shRect.Visible
   mnuPaste.Enabled = Clipboard.GetFormat(vbCFBitmap)
End Sub

Private Sub mnuExit_Click()
   Form_QueryUnload 0, 0
   End
End Sub

Private Sub mnuFile_Click()
   Dim ret
   ret = InStr(1, MDI.ActiveForm.Caption, "Image")
   If ret = 1 Then
      MDI.ActiveForm.mnuDeleteFile.Enabled = False
   Else
      MDI.ActiveForm.mnuDeleteFile.Enabled = True
   End If
End Sub

Private Sub mnuFloat_Click()
   If mnuFloat.Checked = True Then
      mnuFloat.Checked = False
   Else
      mnuFloat.Checked = True
   End If
End Sub

Private Sub mnuGlitter_Click()
   setSwitchesFalse
   RetouchIt = True
   setPointer
   MDI.cboRetouch.Text = "Glitter"
   MDI.cboRetouch.Visible = True
   shCircle.Visible = True
   MDI.Text1.Visible = False
   MDI.Text2.Visible = False
   MDI.txtS.Visible = False
   MDI.UpDownS.Visible = False
   MDI.txtP.Visible = False
   MDI.UpDownP.Visible = False
   MDI.tbVert.Buttons.Item(13).Value = tbrPressed
End Sub

Private Sub mnuGray_Click()
   Dim pixels() As RGBTriplet
   Dim bits_per_pixel As Integer
   Dim x As Integer
   Dim y As Integer
   Dim shade As Integer
   Screen.MousePointer = 11
   picTemp.ScaleWidth = picImage.ScaleWidth
   picTemp.ScaleHeight = picImage.ScaleHeight
   picTemp.Width = picImage.Width
   picTemp.Height = picImage.Height
   ' Get the pixels from picimage.
   GetBitmapPixels picImage, pixels, bits_per_pixel
   ' Set the pixel colors.
   For y = 0 To picImage.ScaleHeight - 1
       For x = 0 To picImage.ScaleWidth - 1
           With pixels(x, y)
                shade = (CInt(.rgbRed) + .rgbGreen + .rgbBlue) / 3
                .rgbRed = shade
                .rgbGreen = shade
                .rgbBlue = shade
           End With
       Next x
   Next y
   ' Set pictemp's pixels.
   SetBitmapPixels picTemp, bits_per_pixel, pixels
   picTemp.Picture = picTemp.Image
   picImage.Picture = LoadPicture()
   picImage.PaintPicture picTemp.Picture, 0, 0
   picImage.Picture = picImage.Image
   Dirty = True
   UpdateUndo
   Screen.MousePointer = 0
End Sub

Private Sub mnuHorz_Click()
   Dim pX As Long, pY As Long, retVal As Long
   On Error GoTo errMsg
   Me.picTemp.Cls
   pX = Me.picImage.ScaleWidth
   pY = Me.picImage.ScaleHeight
   Me.picTemp.Width = Me.picImage.Width
   Me.picTemp.Height = Me.picImage.Height
   retVal = StretchBlt(Me.picTemp.hDc, pX - 1, 0, -pX, pY, _
   Me.picImage.hDc, 0, 0, pX, pY, SRCCOPY)
   Me.picImage.Cls
   Me.picTemp.Picture = Me.picTemp.Image
   Me.picImage.PaintPicture Me.picTemp.Picture, 0, 0, _
   Me.picTemp.Width, Me.picTemp.Height, 0, 0, _
   Me.picTemp.Width, Me.picTemp.Height, vbSrcCopy
   Me.picImage.Picture = Me.picImage.Image
   Me.UpdateUndo
   Me.Refresh
   Dirty = True
   Exit Sub
   
errMsg:
MsgBox "Error # " & Err.Number & " " & Err.Description
Err.Clear
Me.picTemp.Cls
Me.picTemp.Picture = LoadPicture()
End Sub

Private Sub mnuIcons_Click()
   MDI.Arrange 3
End Sub

Private Sub mnuNone_Click()
   FileSelectNone
End Sub

Public Sub FileSelectNone()
   shRect.Visible = False
   Timer1.Enabled = False
   MoveIt = False
   setSwitchesFalse
   Dim i
   For i = 1 To MDI.tbVert.Buttons.count
       MDI.tbVert.Buttons.Item(i).Value = tbrUnpressed
       picImage.MouseIcon = ImageListCursors.ListImages("pencil").Picture
   Next i
End Sub

Private Sub mnuPasteNewImage_Click()
   Call FilePasteNewImage
End Sub

Public Sub FilePasteNewImage()
   If Clipboard.GetFormat(vbCFBitmap) Then
      Dim frm As New frmImage
      shRect.Visible = False
      Timer1.Enabled = False
      setHorzButtons
      frm.Show
      With frm
           .picImage.AutoSize = True
           .picImage.AutoRedraw = True
      End With
      Num = Num + 1
      MDI.ActiveForm.Caption = "Image " & Num
      MDI.ActiveForm.picImage.Picture = Clipboard.GetData(vbCFBitmap)
      MDI.ActiveForm.FitToImage
      MDI.ActiveForm.UpdateUndo
      MDI.ActiveForm.setDirty
   End If
End Sub

Public Sub setDirty()
   Dirty = True
End Sub

Private Sub mnuPasteNewSelection_Click()
   Call FilePasteNewSelection
End Sub

Public Sub FilePasteNewSelection()
   If Clipboard.GetFormat(vbCFBitmap) Then
      picClip.Picture = Clipboard.GetData(vbCFBitmap)
      picImage.PaintPicture picClip.Picture, 0, 0
      shRect.Move 0, 0, picClip.Width + 2, picClip.Height + 2
      shRect.Visible = True
      Timer1.Enabled = True
      PasteIt = True
   End If
End Sub

Private Sub mnuPopCancelSelect_Click()
   MoveIt = False
   shRect.Visible = False
   Timer1.Enabled = False
   picImage.MouseIcon = ImageListCursors.ListImages(5).Picture 'pencil
   setHorzButtons
End Sub

Private Sub mnuPopCopy_Click()
   mnuCopy_Click
End Sub

Private Sub mnuPopCut_Click()
   mnuCut_Click
End Sub

Private Sub mnuPopDelete_Click()
   mnuDelete_Click
End Sub

Private Sub mnuResizeImage_Click()
   AR = picImage.Width / picImage.Height 'Aspect Ratio
   Me.picImage.AutoSize = True
   DoEvents
   frmNewOptions.Check1.Caption = "Keep Aspect Ratio"
   frmNewOptions.txtWidth.Text = picImage.Width
   frmNewOptions.txtHeight.Text = picImage.Height
   frmNewOptions.Check1.Value = 1
   OkNewHW = False
   cancelIt = False
   frmNewOptions.Show '1, Me
   While Not OkNewHW
      DoEvents
   Wend
   OkNewHW = True
   If cancelIt Then Exit Sub
   picTemp.Picture = picImage.Picture
   Me.picImage.Width = val(frmNewOptions.txtWidth.Text)
   Me.picImage.Height = val(frmNewOptions.txtHeight.Text)
   Me.picImage.Picture = LoadPicture()
   picImage.PaintPicture picTemp.Picture, 0, 0, picImage.Width, picImage.Height
   FitToImage
   Me.UpdateUndo
   Dirty = True
End Sub

Private Sub mnuSize_Click()
   MsgBox "File Size is " & FileSize & suffit & " : " & picImage.Width & " Pixels Wide and " & picImage.Height & " Pixels High"
End Sub

Private Sub mnuSoften_Click()
   setSwitchesFalse
   RetouchIt = True
   setPointer
   MDI.cboRetouch.Text = "Soften"
   MDI.tbVert.Buttons.Item(13).Value = tbrPressed
   ShowEffects
End Sub

Private Sub ShowEffects()
   MDI.Text1.Visible = True
   MDI.Text2.Visible = True
   MDI.txtS.Visible = True
   MDI.UpDownS.Visible = True
   MDI.txtP.Visible = True
   MDI.UpDownP.Visible = True
   MDI.cboRetouch.Visible = True
End Sub

Private Sub mnuTileHorz_Click()
   MDI.Arrange 1
   Dim chkme As Form
   For Each chkme In Forms
       If chkme.Name = "frmImage" Then
          Dim ctrl
          For Each ctrl In chkme.Controls
              If ctrl.Name = "picImage" Then
                 ctrl.Left = 0
                 ctrl.Top = 0
              End If
          Next ctrl
       End If
   Next chkme
End Sub

Private Sub mnuTileVert_Click()
   MDI.Arrange 2
   Dim chkme As Form
   For Each chkme In Forms
       If chkme.Name = "frmImage" Then
          Dim ctrl
          For Each ctrl In chkme.Controls
              If ctrl.Name = "picImage" Then
                 ctrl.Left = 0
                 ctrl.Top = 0
              End If
          Next ctrl
       End If
   Next chkme
End Sub

Private Sub mnuCloseWindows_Click()
   Dim chkme As Form
   For Each chkme In Forms
       If chkme.Caption <> "ADMDraw" Then
          Unload chkme
       End If
   Next chkme
End Sub

Private Sub mnuVert_Click()
   Dim pX As Long, pY As Long, retVal As Long
   On Error GoTo errMsg
   Me.picTemp.Cls
   pX = Me.picImage.ScaleWidth
   pY = Me.picImage.ScaleHeight
   Me.picTemp.Width = Me.picImage.Width
   Me.picTemp.Height = Me.picImage.Height
   retVal = StretchBlt(Me.picTemp.hDc, 0, pY - 1, pX, -pY, _
   Me.picImage.hDc, 0, 0, pX, pY, SRCCOPY)
   Me.picImage.Cls
   Me.picTemp.Picture = Me.picTemp.Image
   Me.picImage.PaintPicture Me.picTemp.Picture, 0, 0, _
   Me.picTemp.Width, Me.picTemp.Height, 0, 0, _
   Me.picTemp.Width, Me.picTemp.Height, vbSrcCopy
   Me.picImage.Picture = Me.picImage.Image
   UpdateUndo
   Dirty = True
   Me.Refresh
   Exit Sub
   
errMsg:
MsgBox "Error # " & Err.Number & " " & Err.Description
Err.Clear
Me.picTemp.Cls
Me.picTemp.Picture = LoadPicture()
End Sub

Private Sub picImage_DblClick()
   Exit Sub
End Sub

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If set2 = "" Then set2 = MDI.txtDrawWidth.Text
   If set2 <> MDI.txtDrawWidth.Text Then
      set1 = set2
      set2 = MDI.txtDrawWidth.Text
   End If
   If Button = 1 Then
      cr = MDI.mseFore.BackColor
   Else
      cr = MDI.mseBack.BackColor
   End If
   '===Show shRect PopUp menu
   If Button = vbRightButton And MoveIt And shRect.Visible And x > shRect.Left And x < (shRect.Left + shRect.Width) And y > shRect.Top And y < shRect.Top + shRect.Height Then
      PopupMenu mnuPopup
      MoveIt = False
      Exit Sub
   ElseIf shRect.Visible And x < shRect.Left And x > (shRect.Left + shRect.Width) And y < shRect.Top And y > shRect.Top + shRect.Height Then
      MoveIt = False
      Timer1.Enabled = False
      shRect.Visible = False
   End If
   '=======Move Paste - Down
   If PasteIt Then
      canPaste = True
      XOff = x - shRect.Left
      YOff = y - shRect.Top
      Exit Sub
   End If
   '=======Move Selected Region - Down======
   If MoveIt Then
      canMove = True
      XOff = x - shRect.Left
      YOff = y - shRect.Top
      If mnuFloat.Checked = False Then
         mnuDelete_Click
         MoveIt = True
         Timer1.Enabled = True
         shRect.Visible = True
         picImage.Cls
         picImage.PaintPicture picClip.Clip, x - XOff, y - YOff
         shRect.Left = x - XOff
         shRect.Top = y - YOff
      End If
      Exit Sub
   End If
   '=======Magnify picImage - Down=======
   If MagnifyIt Then
      Screen.MousePointer = 11
      If isMagged = False Then
         picOrig.Width = picImage.Width
         picOrig.Height = picImage.Height
         picOrig.PaintPicture picImage.Picture, 0, 0
         picOrig.Picture = picOrig.Image
      End If
      picImage.Tag = "True"
      isMagged = True
      If mag = 0 Then mag = 1
      If Button = 1 Then
         mag = mag + 1
         If mag > 3 Then mag = 3
         picImage.Picture = LoadPicture()
         picImage.Width = mag * picOrig.Width
         picImage.Height = mag * picOrig.Height
         picImage.PaintPicture picOrig.Picture, 0, 0, mag * picOrig.Width, mag * picOrig.Height
         picImage.Picture = picImage.Image
         picImage.Refresh
         FitToImage
      Else
         mag = (mag - 1)
         If mag = 0 Then mag = 1
         picImage.Picture = LoadPicture()
         picImage.Width = mag * picOrig.Width
         picImage.Height = mag * picOrig.Height
         picImage.PaintPicture picOrig.Picture, 0, 0, mag * picOrig.Width, mag * picOrig.Height
         picImage.Picture = picImage.Image
         picImage.Refresh
         FitToImage
      End If
      Screen.MousePointer = 0
      Exit Sub
   End If
   '=======Use Brush - Down=============
   If DrawIt Then
      canDraw = True
      Dirty = True
      X1 = x
      Y1 = y
      If shRect.Visible Then
         setupSelect
         picTemp.DrawWidth = val(MDI.txtDrawWidth.Text)
         picTemp.Line (X1 - picTemp.Left, Y1 - picTemp.Top)-(x - picTemp.Left, y - picTemp.Top), cr
      Else
         picImage.DrawWidth = val(MDI.txtDrawWidth.Text)
         picImage.Line (X1, Y1)-(x, y), cr
      End If
      Exit Sub
   End If
   '=======Eraser - Down=============
   If EraseIt Then
      canErase = True
      Dirty = True
      X1 = x
      Y1 = y
      If shRect.Visible Then
         setupSelect
         picTemp.DrawWidth = val(MDI.txtDrawWidth.Text)
         picTemp.Line (X1 - picTemp.Left, Y1 - picTemp.Top)-(x - picTemp.Left, y - picTemp.Top), MDI.mseBack.BackColor
      Else
         picImage.DrawWidth = val(MDI.txtDrawWidth.Text)
         picImage.Line (X1, Y1)-(x, y), MDI.mseBack.BackColor
      End If
      Exit Sub
   End If
   '=======Flood - Down==========
   If FloodIt Then
      Dirty = True
      picImage.FillStyle = 0 'Solid
      If shRect.Visible Then
         picTemp.FillStyle = 0 'Solid
         setupSelect
         If Button = 1 Then
            picTemp.FillColor = MDI.mseFore.BackColor
         Else
            picTemp.FillColor = MDI.mseBack.BackColor
         End If
         ExtFloodFill picTemp.hDc, x - picTemp.Left, y - picTemp.Top, picTemp.Point(x - picTemp.Left, y - picTemp.Top), FLOODFILLSURFACE
         keepChange
      Else
         If Button = 1 Then
            picImage.FillColor = MDI.mseFore.BackColor
         Else
            picImage.FillColor = MDI.mseBack.BackColor
         End If
         ExtFloodFill picImage.hDc, x, y, picImage.Point(x, y), FLOODFILLSURFACE
      End If
      MDI.ActiveForm.UpdateUndo
      Exit Sub
   End If
   '=======Spray Can - Down
   If SprayIt Then
      canSpray = True
      If shRect.Visible Then
         setupSelect
         DrawAirBrush Me.picTemp.hDc, x - picTemp.Left, y - picTemp.Top, val(MDI.txtDrawWidth.Text) / 2, val(MDI.txtS.Text)
      Else
         DrawAirBrush Me.picImage.hDc, x, y, val(MDI.txtDrawWidth.Text) / 2, val(MDI.txtS.Text)
      End If
      DoEvents
      Me.picImage.Picture = Me.picImage.Image
      DoEvents
      Exit Sub
   End If
   '=======Add Text - Down==============
   If TextIt = True Then
      canText = True
      curX = x
      curY = y
      frmText.Text1.ForeColor = MDI.mseFore.BackColor
      frmText.Show
      While canText
         DoEvents
      Wend
      picImage.CurrentX = curX - 5
      picImage.CurrentY = curY - 5
      picImage.Font.Size = frmText.Text1.Font.Size
      picImage.Font.Name = frmText.Text1.Font.Name
      picImage.Font.Bold = frmText.Text1.Font.Bold
      picImage.Font.Italic = frmText.Text1.Font.Italic
      picImage.Font.Underline = frmText.Text1.Font.Underline
      picImage.Font.Strikethrough = frmText.Text1.FontStrikethru
      picImage.ForeColor = MDI.mseFore.BackColor 'frmText.Text1.ForeColor
      PrintMultiline picImage, frmText.Text1.Text, x, y
      Dirty = True
      MDI.ActiveForm.UpdateUndo
      Exit Sub
   End If
   '======Line - Down===========
   If LineIt Then
      canLine = True
      X1 = x
      Y1 = y
      Line1.BorderColor = cr
      Line1.BorderWidth = val(MDI.txtDrawWidth.Text)
      Exit Sub
   End If
   '=====Rectangle Filled - Down======
   If RectIt Then
      canRect = True
      X1 = x
      Y1 = y
      Exit Sub
   End If
   '=====Rectangle Empty - Down======
   If RectEmptyIt Then
      canRectEmpty = True
      X1 = x
      Y1 = y
      Exit Sub
   End If
   '======Circle Filled - Down=====
   If CircleIt Then
      canCircle = True
      X1 = x
      Y1 = y
      Exit Sub
   End If
   '======Circle Empty - Down=======
   If CircleEmptyIt Then
      canCircleEmpty = True
      X1 = x
      Y1 = y
      Exit Sub
   End If
   '=====Pick Color - Down
   If PickColor Then
      If Button = 1 Then
         MDI.mseFore.BackColor = picImage.Point(x, y)
      Else
         MDI.mseBack.BackColor = picImage.Point(x, y)
      End If
      Exit Sub
   End If
   '=====Retouch - Down
   If RetouchIt Then
      canRetouch = True
      If MDI.cboRetouch.Text = "Clone" Then
         If Button = vbRightButton Then
            canClone = True
            shRed.Visible = True
            w = val(MDI.txtDrawWidth.Text)
            h = w
            l = x - (w / 2)
            t = y - (h / 2)
            picClip.Picture = picImage.Picture
            picClip.ClipWidth = w
            picClip.ClipHeight = h
            picClip.ClipX = l
            picClip.ClipY = t
            shRed.Width = w
            shRed.Height = h
            shRed.Left = l
            shRed.Top = t
         ElseIf Button = vbLeftButton Then
            If canClone Then
               fTime = True
               w = val(MDI.txtDrawWidth.Text)
               h = w
               l = x - (w / 2)
               t = y - (h / 2)
               picImage.PaintPicture picClip.Clip, l, t, picClip.ClipWidth, picClip.ClipHeight
               picImage.Picture = picImage.Image
               picClip.Picture = picImage.Picture
            End If
         End If
      End If
      If MDI.cboRetouch.Text = "Soften" Then
         Me.picImage.AutoRedraw = True
         DrawBlur Me.picImage.hDc, x, y, (val(MDI.txtDrawWidth.Text) / 2), val(MDI.txtS.Text)
         DoEvents
         Me.picImage.Picture = Me.picImage.Image
         DoEvents
         Exit Sub
      End If
      If MDI.cboRetouch.Text = "Glitter" Then
         With Me.picImage
            .DrawWidth = 1
         End With
         Randomize Timer
         ctx = x + (1 + (Rnd * 16))
         cty = y + (1 + (Rnd * 10))
         ctx1 = x + (1 + (Rnd * 6))
         cty1 = y + (1 + (Rnd * 7))
         ctx2 = x + (1 + (Rnd * 5))
         cty2 = y + (1 + (Rnd * 4))
         Me.picImage.PSet (ctx, cty), cr
         Me.picImage.PSet (ctx1, cty1), cr
         Me.picImage.PSet (ctx2, cty2), cr
         Exit Sub
      End If
      Exit Sub
   End If
   '====Lasso - Down========
   If LassoIt Then
      Exit Sub
   End If
   '====Ants - Down=====
   If AntsIt Then
      canAnts = True
      shRect.Visible = False
      Timer1.Enabled = False
      xStart = x
      yStart = y
      XLo = x
      YLo = y
      XHi = x
      YHi = y
      shRect.Width = Abs(XHi - XLo)
      shRect.Height = Abs(YHi - YLo)
      Exit Sub
   End If
End Sub

Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo chkErr
   clr = picImage.Point(x, y)
   MDI.ShowColor
   If canAnts Or canRect Or canRectEmpty Then
      MDI.StatusBar1.Panels(1).Text = "(" & X1 & "," & Y1 & ") => (" & x & "," & y & ") [" & Abs(X1 - x) + 1 & " X " & Abs(Y1 - y) + 1 & "]"
   Else
      MDI.StatusBar1.Panels(1).Text = x & "," & y
   End If
   MDI.StatusBar1.Panels(2).Text = "File: " & FileSize & " " & suffit
   MDI.StatusBar1.Panels(3).Text = "Image " & picImage.ScaleWidth & " X " & picImage.ScaleHeight
   '===Circle Pointer for Retouch
   If RetouchIt And MDI.cboRetouch.Text <> "Clone" Then
      shCircle.Visible = True
      shSquare.Visible = False
      shCircle.Width = val(MDI.txtDrawWidth.Text) + 2
      shCircle.Height = shCircle.Width
      shCircle.Left = x - (val(MDI.txtDrawWidth.Text) / 2) - 1
      shCircle.Top = y - (val(MDI.txtDrawWidth.Text) / 2) - 1
   End If
   '===Square Pointer for Retouch
   If RetouchIt And MDI.cboRetouch.Text = "Clone" Then
      shSquare.Visible = True
      shCircle.Visible = False
      shSquare.Width = val(MDI.txtDrawWidth.Text) + 2
      shSquare.Height = shSquare.Width
      shSquare.Left = x - (val(MDI.txtDrawWidth.Text) / 2) - 1
      shSquare.Top = y - (val(MDI.txtDrawWidth.Text) / 2) - 1
   End If
   '=======Move Scroll Bars======
   If x >= Abs(picImage.Left) + Me.ScaleWidth - Me.vbarScroller.Width - 1 Then
      If MagnifyIt Then GoTo aroundIt
      If Abs(hbarScroller.Value - 5) < Abs(hbarScroller.Max) Then
         hbarScroller.Value = hbarScroller.Value - 5
         picImage.Left = hbarScroller.Value
         picImage.Refresh
      Else
         hbarScroller.Value = hbarScroller.Max
      End If
    
exitX:
End If

      If y >= Abs(picImage.Top) + Me.ScaleHeight - Me.hbarScroller.Height - 1 Then
         If Abs(vbarScroller.Value - 5) < Abs(vbarScroller.Max) Then
            vbarScroller.Value = vbarScroller.Value - 5
            picImage.Top = vbarScroller.Value
            picImage.Refresh
         Else
            vbarScroller.Value = vbarScroller.Max
         End If
    
exitY:
End If

      If x <= Abs(picImage.Left) + 1 Then
         If hbarScroller.Value + 5 < 0 Then
            hbarScroller.Value = hbarScroller.Value + 5
            picImage.Left = hbarScroller.Value
            picImage.Refresh
         Else
            hbarScroller.Value = hbarScroller.Min
         End If
      End If
      If y <= Abs(picImage.Top) + 1 Then
         If vbarScroller.Value + 5 < 0 Then
            vbarScroller.Value = vbarScroller.Value + 5
            picImage.Top = vbarScroller.Value
            picImage.Refresh
         Else
            vbarScroller.Value = vbarScroller.Min
         End If
    
aroundIt:
End If

   '=======Move Paste - Move======
   If PasteIt Then
      If x > shRect.Left And x < shRect.Left + shRect.Width And y > shRect.Top And y < shRect.Top + shRect.Height Then
         picImage.MouseIcon = ImageListCursors.ListImages(handId).Picture 'hand
      Else
         picImage.MouseIcon = ImageListCursors.ListImages(drawID).Picture 'pencil
      End If
   End If
   If canPaste Then
      picImage.Cls
      picImage.PaintPicture Clipboard.GetData(vbCFBitmap), x - XOff, y - YOff
      shRect.Left = x - XOff
      shRect.Top = y - YOff
      Exit Sub
   End If
   '=======Move Selected Region - Move====
   If MoveIt And shRect.Visible = True Then
      If x > shRect.Left And x < shRect.Left + shRect.Width And y > shRect.Top And y < shRect.Top + shRect.Height Then
         picImage.MouseIcon = ImageListCursors.ListImages(handId).Picture 'hand
      Else
         picImage.MouseIcon = ImageListCursors.ListImages(drawID).Picture 'pencil
      End If
   End If
   If canMove And shRect.Visible And x > shRect.Left And x < (shRect.Left + shRect.Width) And y > shRect.Top And y < shRect.Top + shRect.Height Then
      picImage.Cls
      picImage.PaintPicture picClip.Clip, x - XOff, y - YOff
      shRect.Left = x - XOff
      shRect.Top = y - YOff
      Exit Sub
   End If
   '=======Magnify picImage - Move=======
   If MagnifyIt Then
      Exit Sub
   End If
   '=======Use Brush - Move=============
   If canDraw Then
      If shRect.Visible Then
         picTemp.Line (X1 - picTemp.Left, Y1 - picTemp.Top)-(x - picTemp.Left, y - picTemp.Top), cr
      Else
         picImage.Line (X1, Y1)-(x, y), cr
      End If
      X1 = x
      Y1 = y
      Exit Sub
   End If
   '=======Eraser - Move=============
   If canErase Then
      If shRect.Visible Then
         picTemp.Line (X1 - picTemp.Left, Y1 - picTemp.Top)-(x - picTemp.Left, y - picTemp.Top), MDI.mseBack.BackColor
      Else
         picImage.Line (X1, Y1)-(x, y), MDI.mseBack.BackColor
      End If
      X1 = x
      Y1 = y
      Exit Sub
   End If
   '=======Flood - Move==========
   If FloodIt Then Exit Sub
   '=======Spray Can - Move=======
   If canSpray Then
      If shRect.Visible Then
         DrawAirBrush Me.picTemp.hDc, x - picTemp.Left, y - picTemp.Top, val(MDI.txtDrawWidth.Text) / 2, val(MDI.txtS.Text)
      Else
         DrawAirBrush Me.picImage.hDc, x, y, val(MDI.txtDrawWidth.Text) / 2, val(MDI.txtS.Text)
      End If
      Exit Sub
   End If
   '=======Add Text - Move==============
   If TextIt Then Exit Sub
   '======Line - Move===========
   If canLine Then
      Line1.Visible = True
      Line1.X1 = X1
      Line1.Y1 = Y1
      Line1.X2 = x
      Line1.Y2 = y
      Exit Sub
   End If
   '=====Rectangle Or Circle Filled/Empty - Move======
   If canRect Or canRectEmpty Or canCircle Or canCircleEmpty Then
      If canRect Or canRectEmpty Then
         If x > X1 Then
            Shape1.Left = X1
         Else
            Shape1.Left = x
         End If
         If y > Y1 Then
            Shape1.Top = Y1
         Else
            Shape1.Top = y
         End If
         Shape1.Width = Abs((x) - (X1))
         Shape1.Height = Abs((y) - (Y1))
         Shape1.Shape = 0 'Rectangle
      End If
      If canCircle Or canCircleEmpty Then
         Shape1.Width = Abs((2 * x) - (2 * X1))
         Shape1.Height = Abs((2 * y) - (2 * Y1))
         Shape1.Left = X1 - (Shape1.Width / 2)
         Shape1.Top = Y1 - (Shape1.Height / 2)
         Shape1.Shape = 2 'Oval
      End If
      If canRect Or canCircle Then
         Shape1.BackStyle = 1 'Opaque
         Shape1.FillStyle = 0 'Solid
      Else
         Shape1.BackStyle = 0 'transparent
         Shape1.FillStyle = 1 'Transparent
      End If
      Shape1.BorderWidth = val(MDI.txtDrawWidth.Text)
      Shape1.BorderColor = cr
      Shape1.FillColor = cr
      Shape1.Visible = True
      Exit Sub
   End If
   '=====Pick Color - Move
   If PickColor Then
      Exit Sub
   End If
   '====Retouch - Move
   If canRetouch Then
      If MDI.cboRetouch.Text = "Clone" Then
         If Button = vbLeftButton And canClone Then
            'CloneLeftButton
            l = x - (picClip.ClipWidth / 2)
            t = y - (picClip.ClipHeight / 2)
            If fTime Then
               lDiff = picClip.ClipX - l
               tDiff = picClip.ClipY - t
               fTime = False
            End If
            picClip.ClipX = l + lDiff
            picClip.ClipY = t + tDiff
            shRed.Left = picClip.ClipX
            shRed.Top = picClip.ClipY
            picImage.PaintPicture picClip.Clip, l, t, picClip.ClipWidth, picClip.ClipHeight
            picImage.Picture = picImage.Image
            picClip.Picture = picImage.Picture
         End If
      End If
      If MDI.cboRetouch.Text = "Soften" Then
         DrawBlur Me.picImage.hDc, x, y, (val(MDI.txtDrawWidth.Text) / 2), val(MDI.txtS.Text)
         DoEvents
         Me.picImage.Picture = Me.picImage.Image
         DoEvents
         Exit Sub
      End If
      If MDI.cboRetouch.Text = "Glitter" Then
         With Me.picImage
            .DrawWidth = 1
         End With
         Randomize Timer
         ctx = x + (1 + (Rnd * 16))
         cty = y + (1 + (Rnd * 10))
         ctx1 = x + (1 + (Rnd * 6))
         cty1 = y + (1 + (Rnd * 7))
         ctx2 = x + (1 + (Rnd * 5))
         cty2 = y + (1 + (Rnd * 4))
         Me.picImage.PSet (ctx, cty), cr
         Me.picImage.PSet (ctx1, cty1), cr
         Me.picImage.PSet (ctx2, cty2), cr
         Exit Sub
      End If
      Exit Sub
   End If
   '====Lasso - Move========
   If canLasso Then
      Exit Sub
   End If
   '====Ants - Move=====
   If canAnts Then
      XHi = x
      YHi = y
      If XHi < 0 Then XHi = 0
      If YHi < 0 Then YHi = 0
      If XHi > picImage.ScaleWidth - 1 Then XHi = picImage.ScaleWidth - 1
      If YHi > picImage.ScaleHeight - 1 Then YHi = picImage.ScaleHeight - 1
      If XLo < 0 Then XLo = 0
      If YLo < 0 Then YLo = 0
      If XLo > picImage.ScaleWidth - 1 Then XLo = picImage.ScaleWidth - 1
      If YLo > picImage.ScaleHeight - 1 Then YLo = picImage.ScaleHeight - 1
      shRect.Width = Abs(XHi - XLo)
      shRect.Height = Abs(YHi - YLo)
      shRect.Visible = True
      Timer1.Enabled = True
      If XHi > XLo And YHi > YLo Then
         shRect.Top = YLo
         shRect.Left = XLo
      End If
      If XHi > XLo And YHi < YLo Then
         shRect.Top = YHi
         shRect.Left = XLo
      End If
      If XHi < XLo And YHi < YLo Then
         shRect.Top = YHi
         shRect.Left = XHi
      End If
      If XHi < XLo And YHi > YLo Then
         shRect.Top = YLo
         shRect.Left = XHi
      End If
      Exit Sub
   End If
   Exit Sub
   
chkErr:
If canClone Then Exit Sub
If Err.Number = 380 Then Resume Next
MsgBox "In Mouse Move: Error " & Err.Number & "  " & Err.Description
setSwitchesFalse
End Sub

Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   '==========Move Paste - Up=====
   If canPaste Then
      PasteIt = False
      canPaste = False
      UpdateUndo
      Dirty = True
      shRect.Visible = False
      Timer1.Enabled = False
      setHorzButtons
      picImage.MouseIcon = ImageListCursors.ListImages(5).Picture
      Exit Sub
   End If
   '==========Move Selected Region - Up ======
   If canMove Then
      MoveIt = False
      canMove = False
      UpdateUndo
      Dirty = True
      shRect.Visible = False
      Timer1.Enabled = False
      setHorzButtons
      picImage.MouseIcon = ImageListCursors.ListImages(5).Picture
      Exit Sub
   End If
   '=======Magnify picImage - Up=======
   If MagnifyIt Then
      Exit Sub
   End If
   '=======Use Brush - Up=============
   If canDraw Then
      If shRect.Visible Then
         keepChange
      End If
      canDraw = False
      UpdateUndo
      Exit Sub
   End If
   '=======Eraser - Up=============
   If canErase Then
      If shRect.Visible Then
         keepChange
      End If
      canErase = False
      MDI.ActiveForm.UpdateUndo
      Exit Sub
   End If
   '=======Flood - Up==========
   If FloodIt Then Exit Sub
   '=======Spray Can - Up=======
   If canSpray Then
      If shRect.Visible Then
         keepChange
      End If
      canSpray = False
      Me.picImage.Picture = Me.picImage.Image
      Dirty = True
      UpdateUndo
      Exit Sub
   End If
   '=======Add Text - Up==============
   If TextIt = True Then
      Exit Sub
   End If
   '======Line - Up===========
   If canLine Then
      canLine = False
      Dirty = True
      Line1.Visible = False
      picImage.DrawWidth = val(MDI.txtDrawWidth.Text)
      picImage.Line (X1, Y1)-(x, y), cr
      MDI.ActiveForm.UpdateUndo
      Exit Sub
   End If
   '=====Rectangle Or Circle Filled/Empty - Up======
   If canRect Or canRectEmpty Or canCircle Or canCircleEmpty Then
      Dim valRad
      Dirty = True
      picImage.DrawWidth = val(MDI.txtDrawWidth.Text)
      DoEvents
      If canRect Then picImage.Line (X1, Y1)-(x, y), cr, BF
      If canRectEmpty Then picImage.Line (X1, Y1)-(x, y), cr, B
      If canCircle Or canCircleEmpty Then
         If canCircle Then picImage.FillStyle = 0 'Solid
         picImage.FillColor = cr
         If Shape1.Width >= Shape1.Height Then
            valRad = Shape1.Width / 2
         Else
            valRad = Shape1.Height / 2
         End If
         picImage.Circle (X1, Y1), valRad, cr, , , (Shape1.Height) / (Shape1.Width)
         picImage.FillStyle = 1 'Transparent
      End If
      Shape1.Visible = False
      canRect = False
      canRectEmpty = False
      canCircle = False
      canCircleEmpty = False
      MDI.ActiveForm.UpdateUndo
      Exit Sub
   End If
   '=====Pick Color - Up
   If PickColor Then
      Exit Sub
   End If
   '====Retouch - Up
   If canRetouch Then
      canRetouch = False
      If MDI.cboRetouch.Text = "Soften" Then
         Me.picImage.AutoRedraw = False
         Dirty = True
         picImage.Picture = picImage.Image
         Me.UpdateUndo
         Exit Sub
      End If
      If MDI.cboRetouch.Text = "Glitter" Then
         Dirty = True
         picImage.Picture = picImage.Image
         UpdateUndo
         Exit Sub
      End If
      'for Clone
      If Button = vbLeftButton Then
         canClone = False
         shRed.Visible = False
      End If
      fTime = False
      Dirty = True
      UpdateUndo
      Exit Sub
   End If
   '====Lasso - Up========
   If canLasso Then
      canLasso = False
      Exit Sub
   End If
   '====Ants - Up=====
   If canAnts Then
      canAnts = False
      picClip.Picture = picImage.Picture
      picClip.ClipX = shRect.Left
      picClip.ClipY = shRect.Top
      picClip.ClipWidth = shRect.Width
      picClip.ClipHeight = shRect.Height
      picTemp.DrawWidth = val(MDI.txtDrawWidth.Text)
      MoveIt = True
      setHorzButtons
      Exit Sub
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim ret
   If Dirty = True Then
      ret = MsgBox("Do you wish to save changes to: " & Me.Caption, vbYesNoCancel)
      If ret = vbYes Then FileSave
      If ret = vbCancel Then Cancel = 1
   End If
   If ColUndo.count > 0 Then
      DeleteCollections
   End If
End Sub

Private Sub mnuFit_Click()
   FitToImage
End Sub

Public Sub FitToImage()
   Dim testw, testh, swT, shT, wT, hT, wid, hgt, deltaH, deltaW
   On Error GoTo getOut
   Me.ScaleMode = 1 'twips
   deltaW = Me.Width - Me.ScaleWidth
   deltaH = Me.Height - Me.ScaleHeight
   Me.ScaleMode = 3 'pixels
   'in pixels
   drawAreaWidth = (MDI.ScaleWidth - MDI.CoolBar1.Width - MDI.tbVert.Width)
   drawAreaHeight = (MDI.ScaleHeight - MDI.tbHorz.Height - MDI.StatusBar1.Height)
   testw = picImage.Width + vbarScroller.Width
   testh = picImage.Height + hbarScroller.Height
   wid = Me.ScaleX(testw, vbPixels, vbTwips)
   hgt = Me.ScaleY(testh, vbPixels, vbTwips)
   If wid >= drawAreaWidth Or hgt >= drawAreaHeight Then
      Me.WindowState = 2 'maximized
   Else
      Me.ScaleWidth = ((Me.picImage.Width)) + Me.vbarScroller.Width
      Me.ScaleHeight = ((Me.picImage.Height)) + Me.hbarScroller.Height
      swT = Me.ScaleX(Me.ScaleWidth, vbPixels, vbTwips)
      shT = Me.ScaleY(Me.ScaleHeight, vbPixels, vbTwips)
      Me.Height = shT + deltaH '405 '405 gets the caption bar
      Me.Width = swT + deltaW '120
   End If
   Exit Sub
   
getOut:
If Err.Number = 384 Then
   'do nothing 'can not me resized when maximized, but it is
Else
   MsgBox "In FitToImage Error " & Err.Number & "  " & Err.Description
End If
'Resume Next
End Sub

Private Sub mnuLayer_Click()
   MsgBox "Not yet coded"
End Sub

Private Sub mnuOpen_Click()
   If Not OkNewHW Then Exit Sub
   frmImage.FileOpen
   If MDI.ActiveForm Is Nothing Then Exit Sub
   MDI.ActiveForm.FitToImage
   MDI.ActiveForm.setHorzButtons
End Sub

Private Sub mnuSave_Click()
   FileSave
End Sub

Private Sub mnuSaveAs_Click()
   Call FileSaveAs
End Sub

Private Sub mnuNew_Click()
   If Not OkNewHW Then Exit Sub
   frmImage.FileNew
   MDI.ActiveForm.FitToImage
   MDI.ActiveForm.setHorzButtons
End Sub

' Position the controls.
Private Sub Form_Resize()
   On Error GoTo errChk
   If WindowState = vbMinimized Then Exit Sub
   If ScaleHeight = 0 Then Exit Sub
   picHolder.Move 0, 0, ScaleWidth - vbarScroller.Width, ScaleHeight - hbarScroller.Height
   If picImage.ScaleWidth < picHolder.ScaleWidth And picImage.ScaleHeight < picHolder.ScaleHeight Then
      picImage.Move (picHolder.ScaleWidth - picImage.Width) \ 2, (picHolder.ScaleHeight - picImage.Height) \ 2
   Else
      picImage.Move 0, 0
   End If
   hbarScroller.Move 0, ScaleHeight - hbarScroller.Height, ScaleWidth - vbarScroller.Width
   vbarScroller.Move ScaleWidth - vbarScroller.Width, 0, vbarScroller.Width, ScaleHeight - hbarScroller.Height
   ' Set the scrollbar properties.
   SetScrollBars
   Exit Sub
   
errChk:
MsgBox "In Resize: " & Err.Number & "  " & Err.Description
Resume Next
End Sub

Private Sub hbarScroller_Change()
   picImage.Left = hbarScroller.Value
End Sub

Private Sub hbarScroller_Scroll()
   picImage.Left = hbarScroller.Value
End Sub

Private Sub Timer1_Timer()
   If shRect.BorderStyle = vbBSDot Then
      shRect.BorderStyle = vbBSDashDot
   Else
      shRect.BorderStyle = vbBSDot
   End If
End Sub

Private Sub vbarScroller_Change()
   picImage.Top = vbarScroller.Value
End Sub

Private Sub vbarScroller_Scroll()
   picImage.Top = vbarScroller.Value
End Sub

' Set scroll bar properties.
Private Sub SetScrollBars()
   vbarScroller.Min = 0
   vbarScroller.Max = picHolder.ScaleHeight - picImage.Height
   vbarScroller.LargeChange = picHolder.ScaleHeight
   vbarScroller.SmallChange = picHolder.ScaleHeight / 5
   hbarScroller.Min = 0
   hbarScroller.Max = picHolder.ScaleWidth - picImage.Width
   hbarScroller.LargeChange = picHolder.ScaleWidth
   hbarScroller.SmallChange = picHolder.ScaleWidth / 5
End Sub

Public Sub changeImageSize()
   OkNewHW = False
   cancelIt = False
   frmNewOptions.Show '1, Me
   While Not OkNewHW
       DoEvents
   Wend
   OkNewHW = True
   If cancelIt Then Exit Sub
   If frmNewOptions.Check1.Value = 0 Then
      Me.WindowState = 0
      DoEvents
      Me.picImage.Width = val(frmNewOptions.txtWidth.Text)
      Me.picImage.ScaleWidth = Me.picImage.Width
      Me.picImage.Height = val(frmNewOptions.txtHeight.Text)
      Me.picImage.ScaleHeight = Me.picImage.Height
   Else
      Me.WindowState = 2
      DoEvents
      Me.picImage.Width = ((MDI.tbHorz.Width - MDI.CoolBar1.Width - MDI.tbVert.Width - vbarScroller.Width) / Screen.TwipsPerPixelX) '650
      Me.picImage.Height = ((MDI.tbVert.Height - hbarScroller.Height) / Screen.TwipsPerPixelY)  '450
   End If
   Dirty = True
   picImage.Refresh
End Sub

Public Sub FileNew()
   TFileSize = "Not Saved"
   Tsuffit = ""
   Num = Num + 1
   Dim frm As New frmImage
   frm.picImage.AutoSize = True
   OkNewHW = False
   cancelIt = False
   frmNewOptions.Check1.Value = 0
   frmNewOptions.Show '1, Me
   While Not OkNewHW
       DoEvents
   Wend
   OkNewHW = True
   If cancelIt Then Exit Sub
   If frmNewOptions.Check1.Value = 0 Then
      frm.WindowState = 0
      DoEvents
      frm.picImage.Width = val(frmNewOptions.txtWidth.Text)
      frm.picImage.Height = val(frmNewOptions.txtHeight.Text)
   Else
      frm.WindowState = 2
      DoEvents
      frm.picImage.Width = ((MDI.tbHorz.Width - MDI.CoolBar1.Width - MDI.tbVert.Width) / Screen.TwipsPerPixelX) - vbarScroller.Width '/ Screen.TwipsPerPixelX) '650
      frm.picImage.Height = (MDI.tbVert.Height / Screen.TwipsPerPixelY) - hbarScroller.Height '/ Screen.TwipsPerPixelY)  '450
   End If
   frm.Caption = "Image " & Num
   frm.Show
   frm.UpdateUndo
   Unload Me
End Sub

Public Sub setGifOptions()
   Dim setOpt, ret1, ret2
   Dim hFile         As Long
   Dim sImgHeader    As String
   Dim sFileHeader   As String
   Dim sBuff         As String
   Dim sPicsBuff     As String
   Dim nImgCount     As Long
   Dim i             As Long
   Dim j             As Long
   Dim Finished As Boolean, GifType As String
   Dim sGifMagic As String
   
   On Error GoTo ErrHandler
   'load the gif into a string buffer
   hFile = FreeFile
   Open ComDia.FileName For Binary Access Read As hFile
   'create a buffer the length of the gif being opened
   'and fill it with non-printable characters.
   'Read all the gif header and all it's frames into sBuff
   sBuff = String(LOF(hFile), Chr(0))
   Get #hFile, , sBuff
   Close #hFile
   'header and end of a gif frame (except the last frame!!)
   'null(0),exclamation point(33),u with tick(249)
   sGifMagic = Chr$(0) & Chr$(33) & Chr$(249)
   ret1 = InStr(1, sBuff, sGifMagic)
   If ret1 = 0 Then Exit Sub                           'EXIT SUB
   GifType = Mid(sBuff, 1, ret1 + 2)
   ret2 = InStr(ret1 + 1, sBuff, sGifMagic)
   If ret2 = 0 Then Exit Sub                           'EXIT SUB
   setOpt = MsgBox("This is an anmimated gif file. Do you wish to open all frames?", vbYesNo)
   If setOpt = vbNo Then Exit Sub                      'EXIT SUB
      '====Frame 1 will be loaded back in the main routine. Will now load Frame 2 to last frame.
      'set pointer ahead 3 bytes from the
      'end of the gif magic number
      i = ret2 + 3
      'create a temporary file in the current directory
      hFile = FreeFile
      Open App.Path & "\tempADMDraw.gif" For Binary As hFile
      'split out each frame of the gif, and
      'write each frame to the temporary file.
      'Then load a frmImage
      nImgCount = 1
      Do
         'increment counter
         nImgCount = nImgCount + 1
         'locate next frame end
         j = InStr(i, sBuff, sGifMagic) + 3 '3 is for 3 char in sGifMagic
         'Right here is where trouble occurs because there is no sGifMagic
         'after last image. It is either EOF or Comment starting with
         'null (0) exclamation(33) and funny p(254).
         '==========
         If j = 3 Then
            'check for comment
            j = InStr(i, sBuff, Chr$(0) & Chr$(33) & Chr$(254)) + 3
            Finished = True
            'no comment
            If j = 3 Then
               j = Len(sBuff) '- 1
               Finished = True
            End If
         End If
         '=========
         'another check
         If j > Len(sGifMagic) Then
            'pad an output string, fill with the
            'frame info, and write to disk. A header
            'needs to be added as well, to assure
            'LoadPicture recognizes it as a gif.
            sPicsBuff = Mid(sBuff, i, (j - 3) - (i))
            Put #hFile, 1, GifType & sPicsBuff 'sPicsBuff 'GifType & sPicsBuff
            'assign the data.
            If nImgCount > 1 Then
               'if this is the second or later frame
               'load the picture into a new frmImage.
               '===============
               Num = Num + 1
               Dim frm As Form
               'Dim frm As New frmImage
               Set frm = New frmImage
               With frm
                  .picImage.AutoRedraw = True
                  .picImage.AutoSize = True
                  .picImage.Picture = LoadPicture(App.Path & "\tempADMDraw.gif")
                  .Caption = "Image " & Num
                  .Show
                  .UpdateUndo
                  .FitToImage
               End With
               '================
            End If
            'update pointer
            i = j
         End If
         'when the j = Instr() command above returns 0,
         '3 is added, so if j = 3 there was no more
         'data in the header. We're done.
         If Finished Then
            Finished = False
            j = 3
         End If
      Loop Until j = 3
      'close and nuke the temp file
      Close #hFile
      Kill App.Path & "\tempADMDraw.gif"
      Exit Sub
      
ErrHandler:
MsgBox "Error No. " & Err.Number & " when reading file. Image Count " & nImgCount, vbCritical
On Error GoTo 0
End Sub

Public Sub FileOpen()
   Dim chkGif, retVal
   Static lastOpen
   If FileDrag Then GoTo fileDragged
   On Error GoTo chkErr
   cancelIt = False
   ComDia.CancelError = True
   If lastOpen <> "" Then
      ComDia.InitDir = lastOpen
   Else
      ComDia.InitDir = App.Path & "\ADMDrawPics"
   End If
   ComDia.Filter = "Images (*.bmp;*.jpg;*gif)|*.bmp;*.jpg;*.gif"
   ComDia.ShowOpen
   '======show file size
   TFileSize = FileLen(ComDia.FileName)
   If TFileSize < 1000 Then
      Tsuffit = " Bytes"
      GoTo endIt
   ElseIf TFileSize >= 1000000 Then
      TFileSize = Int(TFileSize / 1000000)
      Tsuffit = " Mb"
      GoTo endIt
   ElseIf TFileSize >= 1000 And TFileSize < 1000000 Then
      TFileSize = Int(TFileSize / 1000)
      Tsuffit = " Kb"

endIt:
End If

   '===========
   retVal = InStrRev(ComDia.FileName, "\")
   lastOpen = Mid(ComDia.FileName, 1, retVal)
   '=============
   ' See is this window is already open and bring it to top.
   Dim chkme As Form
   For Each chkme In Forms
       If chkme.Caption = ComDia.FileName Then
          chkme.WindowState = vbNormal
          chkme.SetFocus
          Exit Sub
       End If
   Next chkme
   
' ====Get Gif options========
fileDragged:
chkGif = InStrRev(LCase(ComDia.FileName), ".gif")

   If chkGif = Len(ComDia.FileName) - 3 Then setGifOptions
   '==========
   Dim frm As New frmImage
   frm.picImage.AutoRedraw = True
   frm.picImage.AutoSize = True
   frm.picImage.Picture = LoadPicture(ComDia.FileName)
   frm.Caption = ComDia.FileName
   frm.Show
   frm.UpdateUndo
   Unload Me
   FileDrag = False
   Exit Sub
      
chkErr:
If Err.Number = 32755 Then 'user pressed cancel
   cancelIt = True
   Exit Sub
Else
   MsgBox Err.Number & " - " & Err.Description
End If
End Sub

Public Sub FileSave()
   Dim retVal, retValGif, retValJpg, retValBmp
   Screen.MousePointer = 11
   retVal = InStr(Me.Caption, "Image")
   retValGif = InStrRev(LCase(Me.Caption), "gif")
   retValJpg = InStrRev(LCase(Me.Caption), "jpg")
   retValBmp = InStrRev(LCase(Me.Caption), "bmp")
   If retValGif = Len(Me.Caption) - 2 Then
      'This is gif file which can only be saved in jpg or bmp format
      mnuSaveAs_Click
      Screen.MousePointer = 0
      Exit Sub
   End If
   If retVal = 1 Then
      mnuSaveAs_Click
      Screen.MousePointer = 0
      Exit Sub
   End If
   picImage.Picture = picImage.Image
   If picImage.Tag = "True" Then 'magged
      picImage.Picture = picImage.Image
      picOrig.PaintPicture picImage.Picture, 0, 0, picOrig.Width, picOrig.Height, 0, 0, picImage.Width, picImage.Height
      picOrig.Picture = picOrig.Image
      DoEvents
      If retValBmp = Len(Me.Caption) - 2 Then SavePicture picOrig.Picture, Me.Caption
      '==============
      If retValJpg = Len(Me.Caption) - 2 Then
         Dim si As String
         Dim c As New cDIBSection
         Dim qual1
         si = Me.Caption 'fileToSave
         c.CreateFromPicture picOrig.Picture
         qual1 = 90
         If SaveJPG(c, si, qual1) Then
            'OK!
         Else
            MsgBox "Failed to save the picture to the file: '" & si & "'", vbExclamation
         End If
      End If
      '============
      If ColUndo.count > 0 Then
         DeleteCollections
      End If
      MDI.tbHorz.Buttons(undoID).Enabled = MDI.ActiveForm.ColUndo.count > 1
      MDI.tbHorz.Buttons(redoID).Enabled = MDI.ActiveForm.ColRedo.count > 0
      picImage.Picture = LoadPicture(Me.Caption)
      FitToImage
      Me.UpdateUndo
      '======show caption file size
      FileSize = FileLen(Me.Caption)
      If FileSize < 1000 Then
         suffit = " Bytes"
         GoTo endIt
      ElseIf FileSize >= 1000000 Then
         FileSize = Int(FileSize / 1000000)
         suffit = " Mb"
         GoTo endIt
      ElseIf FileSize >= 1000 And FileSize < 1000000 Then
         FileSize = Int(FileSize / 1000)
         suffit = " Kb"
        
endIt:
End If

   '===========
   Else 'not magged
      If retValBmp = Len(Me.Caption) - 2 Then SavePicture picImage.Picture, Me.Caption
      '==============
      If retValJpg = Len(Me.Caption) - 2 Then
         si = Me.Caption 'fileToSave
         c.CreateFromPicture picImage.Picture
         qual1 = 90
         If SaveJPG(c, si, qual1) Then
            'OK!
         Else
            MsgBox "Failed to save the picture to the file: '" & si & "'", vbExclamation
         End If
      End If
      '============
      picImage.Picture = LoadPicture(Me.Caption)
      FitToImage
      '======show caption file size
      FileSize = FileLen(Me.Caption)
      If FileSize < 1000 Then
         suffit = " Bytes"
         GoTo endIt2
      ElseIf FileSize >= 1000000 Then
         FileSize = Int(FileSize / 1000000)
         suffit = " Mb"
         GoTo endIt2
      ElseIf FileSize >= 1000 And FileSize < 1000000 Then
         FileSize = Int(FileSize / 1000)
         suffit = " Kb"
         
endIt2:
End If

   '===========
   End If
   Dirty = False
   isMagged = False
   picImage.Tag = ""
   mag = 0
   '==================
   ' See if Browse is open and refresh File1.
   Dim chkme As Form
   For Each chkme In Forms
       retVal = Mid(chkme.Caption, 1, 6)
       If retVal = "Browse" Then
          saveFile = Me.Caption
          Call frmBrowse.saveOne
       End If
   Next chkme
   '============
   Screen.MousePointer = 0
End Sub

Public Sub FileSaveAs()
   Dim sName, retVal, retSave
   Screen.MousePointer = 11
   '=============
   ComDia.FilterIndex = 1
   If lastSave <> "" Then
      ComDia.InitDir = lastSave
   Else
      ComDia.InitDir = App.Path & "\ADMDrawPics"
   End If
   '===========
   On Error GoTo ErrHandler
   retVal = InStrRev(Me.Caption, "\")
   If retVal <> 0 Then
      sName = Mid(Me.Caption, retVal + 1, Len(Me.Caption) - 4 - retVal)
   Else
      sName = Me.Caption
   End If
   ComDia.FileName = sName '"MyDrawing"
   ComDia.CancelError = True
   ComDia.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
   ComDia.Filter = "JPG File (*.jpg)|*.jpg|Bitmaps (*.bmp)|*.bmp|Gif (*.gif)|*.gif|Transparent Gif (*.gif)|*.gif"
   ComDia.ShowSave
   retSave = InStrRev(ComDia.FileName, "\")
   lastSave = Mid(ComDia.FileName, 1, retSave)
   DoEvents
   '============
   picImage.Picture = picImage.Image
   Select Case ComDia.FilterIndex
   Case 1 'save as jpg
        Dim si As String
        Dim c As New cDIBSection
        Dim qual1
        si = ComDia.FileName 'fileToSave
        If picImage.Tag = "True" Then
           'picImage.Picture = picImage.Image
           picOrig.PaintPicture picImage.Picture, 0, 0, picOrig.Width, picOrig.Height, 0, 0, picImage.Width, picImage.Height
           picOrig.Picture = picOrig.Image
           DoEvents
           c.CreateFromPicture picOrig.Picture
        Else
           c.CreateFromPicture picImage.Picture
        End If
        
        qual1 = 90
        '*************************************
        If si <> "" Then
           If SaveJPG(c, si, qual1) Then
              'OK!
           Else
              MsgBox "Failed to save the picture to the file: '" & si & "'", vbExclamation
           End If
        End If
   Case 2 'save as bmp
        If picImage.Tag = "True" Then
           'picImage.Picture = picImage.Image
           picOrig.PaintPicture picImage.Picture, 0, 0, picOrig.Width, picOrig.Height, 0, 0, picImage.Width, picImage.Height
           picOrig.Picture = picOrig.Image
           DoEvents
           SavePicture picOrig.Picture, ComDia.FileName
        Else
           SavePicture picImage.Picture, ComDia.FileName
        End If
   Case 3
        Set cGif = New GIF
        If picImage.Tag = "True" Then
           picOrig.PaintPicture picImage.Picture, 0, 0, picOrig.Width, picOrig.Height, 0, 0, picImage.Width, picImage.Height
           picOrig.Picture = picOrig.Image
           DoEvents
           cGif.SaveGIF picOrig.Picture, ComDia.FileName, picOrig.hDc, 0, picOrig.Point(0, 0)
        Else
           cGif.SaveGIF picImage.Picture, ComDia.FileName, picImage.hDc, 0, picImage.Point(0, 0)
        End If
        Set cGif = Nothing
   Case 4
        Set cGif = New GIF
        If picImage.Tag = "True" Then
           picOrig.PaintPicture picImage.Picture, 0, 0, picOrig.Width, picOrig.Height, 0, 0, picImage.Width, picImage.Height
           picOrig.Picture = picOrig.Image
           DoEvents
           cGif.SaveGIF picOrig.Picture, ComDia.FileName, picOrig.hDc, 1, picOrig.Point(0, 0)
        Else
           cGif.SaveGIF picImage.Picture, ComDia.FileName, picImage.hDc, 1, picImage.Point(0, 0)
        End If
        Set cGif = Nothing
   Case Else
        MsgBox "Filter Index not found."
   End Select
   isMagged = False
   picImage.Tag = ""
   Dirty = False
   mag = 0
   If ColUndo.count > 0 Then
      DeleteCollections
   End If
   MDI.tbHorz.Buttons(undoID).Enabled = MDI.ActiveForm.ColUndo.count > 1
   MDI.tbHorz.Buttons(redoID).Enabled = MDI.ActiveForm.ColRedo.count > 0
   picImage.Picture = LoadPicture(ComDia.FileName)
   FitToImage
   Me.UpdateUndo
   With Me
      .Caption = ComDia.FileName
   End With
   '======show caption file size
   FileSize = FileLen(Me.Caption)
   If FileSize < 1000 Then
      suffit = " Bytes"
      GoTo endIt
   ElseIf FileSize >= 1000000 Then
      FileSize = Int(FileSize / 1000000)
      suffit = " Mb"
      GoTo endIt
   ElseIf FileSize >= 1000 And FileSize < 1000000 Then
      FileSize = Int(FileSize / 1000)
      suffit = " Kb"
      
endIt:
End If

      '==================
      ' See if Browse is open and refresh File1.
      Dim chkme As Form
      For Each chkme In Forms
          retVal = Mid(chkme.Caption, 1, 6)
          If retVal = "Browse" Then
             Call frmBrowse.addOne
             Exit For
          End If
      Next chkme
      
      '============
      Screen.MousePointer = 0
      Exit Sub
      
ErrHandler:
If Err.Number = 32755 Then
   Screen.MousePointer = 0
   Exit Sub
Else
   If Err.Number <> 0 Then MsgBox "Error saving file: " & Err.Number & " - " & Err.Description
   Screen.MousePointer = 0
End If
End Sub

Public Sub DeleteCollections()
   If MDI.ActiveForm.Caption = "Browse" Then Exit Sub
   For Each vItem In MDI.ActiveForm.ColUndo
       MDI.ActiveForm.ColUndo.Remove 1
   Next
   For Each vItem In MDI.ActiveForm.ColRedo
       MDI.ActiveForm.ColRedo.Remove 1
   Next
   UniqueNum = 0
End Sub

Public Sub UpdateUndo()
   MDI.ActiveForm.UniqueNum = MDI.ActiveForm.UniqueNum + 1
   MDI.ActiveForm.picImage.Picture = MDI.ActiveForm.picImage.Image
   MDI.ActiveForm.ColUndo.Add Item:=MDI.ActiveForm.picImage.Picture, Key:=CStr(MDI.ActiveForm.UniqueNum)
   MDI.tbHorz.Buttons(undoID).Enabled = MDI.ActiveForm.ColUndo.count > 1
   MDI.tbHorz.Buttons(redoID).Enabled = MDI.ActiveForm.ColRedo.count > 0
   If Dirty = False Then Exit Sub
   Dirty = True
End Sub

Public Sub DoUnDo()
   MDI.ActiveForm.ColRedo.Add MDI.ActiveForm.ColUndo.Item(MDI.ActiveForm.ColUndo.count)
   MDI.ActiveForm.ColUndo.Remove MDI.ActiveForm.ColUndo.count
   MDI.ActiveForm.picImage.Picture = MDI.ActiveForm.ColUndo.Item(MDI.ActiveForm.ColUndo.count)
   MDI.ActiveForm.picImage.Refresh
   MDI.tbHorz.Buttons(undoID).Enabled = MDI.ActiveForm.ColUndo.count > 1
   MDI.tbHorz.Buttons(redoID).Enabled = MDI.ActiveForm.ColRedo.count > 0
   MDI.ActiveForm.FitToImage
End Sub

Public Sub DoReDo()
   On Error Resume Next
   MDI.tbHorz.Buttons(redoID).Visible = MDI.ActiveForm.ColRedo.count > 0
   MDI.ActiveForm.ColUndo.Add MDI.ActiveForm.ColRedo.Item(MDI.ActiveForm.ColRedo.count)
   MDI.ActiveForm.ColRedo.Remove MDI.ActiveForm.ColRedo.count
   MDI.ActiveForm.picImage.Picture = MDI.ActiveForm.ColUndo.Item(MDI.ActiveForm.ColUndo.count)
   MDI.ActiveForm.picImage.Refresh
   MDI.tbHorz.Buttons(redoID).Enabled = MDI.ActiveForm.ColRedo.count > 0
   MDI.tbHorz.Buttons(undoID).Enabled = MDI.ActiveForm.ColUndo.count > 1
   MDI.ActiveForm.FitToImage
End Sub

Public Sub ClearRedo()
   For Each vItem In MDI.ActiveForm.ColRedo
       MDI.ActiveForm.ColRedo.Remove 1
   Next
End Sub

Public Sub setPointer()
   If picImage.MouseIcon <> ImageListCursors.ListImages(drawID).Picture Then picImage.MouseIcon = ImageListCursors.ListImages(drawID).Picture
   If DrawIt Then picImage.MouseIcon = ImageListCursors.ListImages(drawID).Picture
   If EraseIt Then picImage.MouseIcon = ImageListCursors.ListImages(eraseID).Picture
   If TextIt Then picImage.MouseIcon = ImageListCursors.ListImages(textID).Picture
   If FloodIt Then picImage.MouseIcon = ImageListCursors.ListImages(floodID).Picture
   If SprayIt Then picImage.MouseIcon = ImageListCursors.ListImages(sprayID).Picture
   If PickColor Then picImage.MouseIcon = ImageListCursors.ListImages(pickcolorID).Picture
   If MagnifyIt Then picImage.MouseIcon = ImageListCursors.ListImages(magnifyID).Picture
   If RetouchIt Then picImage.MouseIcon = ImageListCursors.ListImages("pencil").Picture
End Sub

Public Sub setHorzButtons()
   Dim i
   If MDI.ActiveForm Is Nothing Then
      For i = 1 To MDI.tbHorz.Buttons.count
          If MDI.tbHorz.Buttons(i).Key = "pastenew" Then pasteNewID = i
          If MDI.tbHorz.Buttons(i).Key = "undo" Then undoID = i
          If MDI.tbHorz.Buttons(i).Key = "redo" Then redoID = i
          If MDI.tbHorz.Buttons(i).Tag = "all" Then
             MDI.tbHorz.Buttons(i).Visible = True
          Else
             MDI.tbHorz.Buttons(i).Visible = False
          End If
      Next i
      If Clipboard.GetFormat(vbCFBitmap) = False Then
         MDI.tbHorz.Buttons(pasteNewID).Visible = False
      Else
         MDI.tbHorz.Buttons(pasteNewID).Visible = True
      End If
      Exit Sub
   End If
   For i = 1 To MDI.tbHorz.Buttons.count
       If MDI.tbHorz.Buttons(i).Key = "pastenew" Then pasteNewID = i
       If MDI.tbHorz.Buttons(i).Tag = "all" Or MDI.tbHorz.Buttons(i).Tag = "form" Then
          MDI.tbHorz.Buttons(i).Visible = True
       Else
          MDI.tbHorz.Buttons(i).Visible = False
       End If
       If MDI.tbHorz.Buttons(i).Tag = "rect" And MDI.ActiveForm.shRect.Visible Then MDI.tbHorz.Buttons(i).Visible = True
       If MDI.tbHorz.Buttons(i).Tag = "rect" And Not MDI.ActiveForm.shRect.Visible Then MDI.tbHorz.Buttons(i).Visible = False
       If MDI.tbHorz.Buttons(i).Tag = "formcb" And Clipboard.GetFormat(vbCFBitmap) = True Then MDI.tbHorz.Buttons(i).Visible = True
       If MDI.tbHorz.Buttons(i).Tag = "allcb" And Clipboard.GetFormat(vbCFBitmap) = True Then MDI.tbHorz.Buttons(i).Visible = True
   Next i
   If Clipboard.GetFormat(vbCFBitmap) = False Then
      MDI.tbHorz.Buttons(pasteNewID).Visible = False
   Else
      MDI.tbHorz.Buttons(pasteNewID).Visible = True
   End If
End Sub

Public Sub PrintMultiline(ByVal obj As Object, ByVal txt As String, ByVal x As Single, ByVal y As Single)
   Dim one_line As String
   Dim pos As Integer
   obj.CurrentY = y
   Do While Len(txt) > 0
      ' Find the next line.
      pos = InStr(txt, vbCrLf)
      If pos = 0 Then
         ' This is the last line.
         one_line = txt
         txt = ""
      Else
         one_line = Left$(txt, pos - 1)
         txt = Right$(txt, Len(txt) - pos + 1 - Len(vbCrLf))
      End If
      ' Print the line. This moves CurrentY to
      ' the next line.
      obj.CurrentX = x
      obj.Print one_line
   Loop
End Sub

Public Function rotatePic(picSrc As Object, picDest As Object, rotdeg As Integer)
   Dim x As Long
   Dim y As Long
   MDI.PB.Visible = True
   If rotdeg = 90 Or rotdeg = 270 Then
      picDest.Height = picSrc.Width
      picDest.Width = picSrc.Height
      picDest.ScaleHeight = picSrc.ScaleWidth
      picDest.ScaleWidth = picSrc.ScaleHeight
   Else 'it's 180
      picDest.Height = picSrc.Height
      picDest.Width = picSrc.Width
      picDest.ScaleHeight = picSrc.ScaleHeight
      picDest.ScaleWidth = picSrc.ScaleWidth
   End If
   For y = 0 To picDest.ScaleHeight - 1
       For x = 0 To picDest.ScaleWidth - 1
           If rotdeg = 90 Then
              Call SetPixel(picDest.hDc, x, y, GetPixel(picSrc.hDc, y, picSrc.ScaleHeight - 1 - x))
           End If
           If rotdeg = 180 Then
              Call SetPixel(picDest.hDc, x, y, GetPixel(picSrc.hDc, picSrc.ScaleWidth - 1 - x, picSrc.ScaleHeight - 1 - y))
           End If
           If rotdeg = 270 Then
              Call SetPixel(picDest.hDc, x, y, GetPixel(picSrc.hDc, picSrc.ScaleWidth - 1 - y, x))
           End If
       Next x
       'see some action every 5 lines
       If y Mod 5 = 0 Then
          DoEvents
          MDI.PB.Value = (y / picDest.ScaleHeight) * 100
       End If
   Next y
   MDI.PB.Visible = False
End Function

Public Sub printPic()
   picImage.Picture = picImage.Image
   Screen.MousePointer = 11
   Printer.Orientation = vbPRORLandscape
   Printer.ColorMode = vbPRCMColor
   Printer.Copies = 1
   Printer.PrintQuality = vbPRPQHigh
   Printer.PaperSize = vbPRPSLetter
   Printer.PaintPicture picImage.Picture, 1900, 2300
   Printer.EndDoc
   Screen.MousePointer = 0
End Sub

Private Sub setupSelect()
   On Error GoTo errChk
   picTemp.Picture = picClip.Clip
   picTemp.ScaleWidth = shRect.Width
   picTemp.Width = picTemp.ScaleWidth
   picTemp.ScaleHeight = shRect.Height
   picTemp.Height = picTemp.ScaleHeight
   picTemp.Left = shRect.Left
   picTemp.Top = shRect.Top
   picTemp.Visible = True
   shFr.Width = shRect.Width + 2
   shFr.Height = shRect.Height + 2
   shFr.Left = shRect.Left - 1
   shFr.Top = shRect.Top - 1
   shFr.Visible = True
   DoEvents
   Exit Sub
   
errChk:
MsgBox "In setupSelect: " & Err.Number & " " & Err.Description
Resume Next
End Sub

Private Sub keepChange()
   picTemp.Picture = picTemp.Image
   picImage.PaintPicture picTemp.Picture, shRect.Left, shRect.Top
   picImage.Picture = picImage.Image
   picTemp.Visible = False

   picClip.Picture = picImage.Picture
   picClip.ClipX = shRect.Left + 1
   picClip.ClipY = shRect.Top + 1
   picClip.ClipWidth = shRect.Width
   picClip.ClipHeight = shRect.Height
   shFr.Visible = False
End Sub

