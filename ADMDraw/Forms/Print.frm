VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Print Preview"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5085
   Icon            =   "Print.frx":0000
   ScaleHeight     =   5010
   ScaleWidth      =   5085
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar vbarScroller 
      Height          =   2985
      Left            =   3240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   240
   End
   Begin VB.HScrollBar hbarScroller 
      Height          =   240
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3285
      Width           =   2520
   End
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   75
      MousePointer    =   99  'Custom
      ScaleHeight     =   3330
      ScaleWidth      =   3030
      TabIndex        =   0
      Top             =   -165
      Width           =   3030
      Begin VB.PictureBox picPrint 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   15840
         Left            =   0
         ScaleHeight     =   15840
         ScaleWidth      =   12240
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   12240
      End
      Begin VB.PictureBox picZoom 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   195
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   4
         Top             =   285
         Width           =   400
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "&Zoom"
      Begin VB.Menu mnu100 
         Caption         =   "100%"
      End
      Begin VB.Menu mnu50 
         Caption         =   "50%"
      End
      Begin VB.Menu mnu25 
         Caption         =   "25%"
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   mnu25_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   picPrint.Picture = LoadPicture()
   picZoom.Picture = LoadPicture()
End Sub

' Position the controls.
Private Sub Form_Resize()
   If WindowState = vbMinimized Then Exit Sub
   If ScaleHeight = 0 Then Exit Sub
   picHolder.Move 0, 0, ScaleWidth - vbarScroller.Width, ScaleHeight - hbarScroller.Height
   If picZoom.ScaleWidth < picHolder.ScaleWidth And picZoom.ScaleHeight < picHolder.ScaleHeight Then
      picZoom.Move (picHolder.ScaleWidth - picZoom.Width) \ 2, (picHolder.ScaleHeight - picZoom.Height) \ 2
   Else
      picZoom.Move 0, 0
   End If
   hbarScroller.Move 0, ScaleHeight - hbarScroller.Height, ScaleWidth - vbarScroller.Width
   vbarScroller.Move ScaleWidth - vbarScroller.Width, 0, vbarScroller.Width, ScaleHeight - hbarScroller.Height

   ' Set the scrollbar properties.
   SetScrollBars
End Sub

' Set scroll bar properties.
Private Sub SetScrollBars()
   vbarScroller.Min = 0
   vbarScroller.Max = picHolder.ScaleHeight - picZoom.Height
   vbarScroller.LargeChange = picHolder.ScaleHeight
   vbarScroller.SmallChange = picHolder.ScaleHeight / 5
   hbarScroller.Min = 0
   hbarScroller.Max = picHolder.ScaleWidth - picZoom.Width
   hbarScroller.LargeChange = picHolder.ScaleWidth
   hbarScroller.SmallChange = picHolder.ScaleWidth / 5
End Sub

Private Sub hbarScroller_Change()
   picZoom.Left = hbarScroller.Value
End Sub

Private Sub hbarScroller_Scroll()
   picZoom.Left = hbarScroller.Value
End Sub


Private Sub mnu100_Click()
   picZoom.Visible = False
   picZoom.Width = picPrint.Width
   picZoom.Height = picPrint.Height
   picZoom.ScaleWidth = picPrint.ScaleWidth
   picZoom.ScaleHeight = picPrint.ScaleHeight
   picZoom.Move 0, 0
   picPrint.Picture = picPrint.Image
   picZoom.PaintPicture picPrint, 0, 0
   picZoom.Visible = True
End Sub

Private Sub mnu25_Click()
   picZoom.Visible = False
   picZoom.Width = picPrint.Width / 4
   picZoom.Height = picPrint.Height / 4
   picZoom.ScaleWidth = picPrint.ScaleWidth / 4
   picZoom.ScaleHeight = picPrint.ScaleHeight / 4
   picZoom.Move (picHolder.ScaleWidth / 2) - (picZoom.Width / 2), (picHolder.ScaleHeight / 2) - (picZoom.Height / 2)
   picPrint.Picture = picPrint.Image
   picZoom.PaintPicture picPrint, 0, 0, picZoom.Width, picZoom.Height
   picZoom.Visible = True
End Sub

Private Sub mnu50_Click()
   picZoom.Visible = False
   picZoom.Width = picPrint.Width / 2
   picZoom.Height = picPrint.Height / 2
   picZoom.ScaleWidth = picPrint.ScaleWidth / 2
   picZoom.ScaleHeight = picPrint.ScaleHeight / 2
   picZoom.Move (picHolder.ScaleWidth / 2) - (picZoom.Width / 2), (picHolder.ScaleHeight / 2) - (picZoom.Height / 2)
   picPrint.Picture = picPrint.Image
   picZoom.PaintPicture picPrint, 0, 0, picZoom.Width, picZoom.Height
   picZoom.Visible = True
End Sub

Private Sub mnuClose_Click()
   picPrint.Picture = LoadPicture()
   picZoom.Picture = LoadPicture()
   Unload Me
End Sub

Private Sub mnuPrint_Click()
   printPic
End Sub

Private Sub vbarScroller_Change()
   picZoom.Top = vbarScroller.Value
End Sub

Private Sub vbarScroller_Scroll()
   picZoom.Top = vbarScroller.Value
End Sub

Public Sub printPic()
   picPrint.Picture = picPrint.Image
   Screen.MousePointer = 11
   If picPrint.Height > picPrint.Width Then
      Printer.Orientation = vbPRORPortrait
   Else
      Printer.Orientation = vbPRORLandscape
   End If
   Printer.ColorMode = vbPRCMColor
   Printer.Copies = 1
   Printer.PrintQuality = vbPRPQHigh
   Printer.PaperSize = vbPRPSLetter
   Printer.PaintPicture picPrint.Picture, 0, 0
   Printer.EndDoc
   Screen.MousePointer = 0
End Sub
