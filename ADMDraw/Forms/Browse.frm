VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowse 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Browse"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5055
   Icon            =   "Browse.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MousePointer    =   9  'Size W E
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   Tag             =   "browse"
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   240
      Left            =   1770
      TabIndex        =   6
      Top             =   3990
      Visible         =   0   'False
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   15
      Width           =   1875
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   225
      Pattern         =   "*gif;*.jpg;*.bmp"
      TabIndex        =   1
      Top             =   3525
      Width           =   1365
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   150
      TabIndex        =   0
      Top             =   375
      Width           =   1215
   End
   Begin VB.PictureBox picOuter 
      BackColor       =   &H00FFFFFF&
      Height          =   3990
      Left            =   1980
      ScaleHeight     =   262
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   4
      Top             =   -45
      Width           =   2820
      Begin VB.PictureBox picHolder 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3645
         Left            =   45
         ScaleHeight     =   243
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   123
         TabIndex        =   5
         Top             =   -60
         Width           =   1845
         Begin VB.Shape tShape 
            BorderColor     =   &H00000080&
            BorderWidth     =   4
            Height          =   1620
            Index           =   0
            Left            =   855
            Shape           =   1  'Square
            Top             =   930
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Image tImage 
            Height          =   1500
            Index           =   0
            Left            =   465
            Stretch         =   -1  'True
            Top             =   30
            Width           =   1500
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2820
      Left            =   4695
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   255
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      Begin VB.Menu mnuWindowsOpt 
         Caption         =   "Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuWindowsOpt 
         Caption         =   "Tile Vertically"
         Index           =   1
      End
      Begin VB.Menu mnuWindowsOpt 
         Caption         =   "Tile Horizonally"
         Index           =   2
      End
      Begin VB.Menu mnuWindowsOpt 
         Caption         =   "Arrange Icons"
         Index           =   3
      End
      Begin VB.Menu mnuWindowsOpt 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuWindowsOpt 
         Caption         =   "Close All Open Windows"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Selected File"
      End
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim canSplit As Boolean, xSplit
Dim tRowsSet, TperRow, rowsRequired, tSpaceRequired
Dim tPathArray(), oldListCount

Public Sub setBrowseButtons()
   Dim i
   For i = 1 To MDI.tbHorz.Buttons.count
       If MDI.tbHorz.Buttons(i).Key = "pastenew" Then pasteNewID = i
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
End Sub

Private Sub Dir1_Change()
   Dim i
   File1.Path = Dir1.Path
   'For some reason, this code does not work on initial opening in IDE
   'but works okay in compiled exe.
   On Error GoTo chkErr

   lastBrowseDir = Dir1.Path
   Screen.MousePointer = 11
   For i = 0 To oldListCount - 1
       tImage(i).Picture = LoadPicture()
       tShape(i).Visible = False
       tImage(i).Visible = False
   Next i
   oldListCount = File1.ListCount
   'check and recheck before removing this line'loadControlArray 'seems to work okay without this
   Call loadThumbs
   Screen.MousePointer = 0
   lastBrowseDir = Dir1.Path
   lastBrowseDrive = Drive1.Drive
   If tImage.UBound <> tShape.UBound Or tImage.UBound <> UBound(tPathArray) Then
      'OK!
   End If
   Exit Sub
   
chkErr:
If Err.Number = 340 Then 'element not loaded, but, will be in next few lines of code.
   Resume Next
Else
   MsgBox "In Dir1_Change: " & Err.Number & " - " & Err.Description
   Resume Next
End If
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MousePointer = 0
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
   lastBrowseDrive = Drive1.Drive
End Sub

Private Sub File1_Click()
   'Show Rectangle
   Dim i
   For i = 0 To File1.ListCount - 1
       tShape(i).Visible = False
   Next i
   tShape(File1.ListIndex).Visible = True
End Sub

Private Sub File1_DblClick()
   Dim imgOpen, chkGif
   imgOpen = Dir1.Path & "\" & File1.List(File1.ListIndex)
   '==================
   ' See is this window is already open and bring it to top.
   Dim chkme As Form
   For Each chkme In Forms
       If LCase(chkme.Caption) = LCase(imgOpen) Then
          chkme.WindowState = vbNormal
          chkme.SetFocus
          Exit Sub
       End If
   Next chkme
   '====Get Gif options========
   chkGif = InStrRev(File1.List(File1.ListIndex), ".gif")
   If chkGif = Len(File1.List(File1.ListIndex)) - 3 Then
      frmImage.ComDia.FileName = Dir1.Path & "\" & File1.List(File1.ListIndex)
      Call frmImage.setGifOptions
   End If
   '==========
   Dim frm As New frmImage
   With frm
      .picImage.AutoRedraw = True
      .picImage.AutoSize = True
      .picImage.Picture = LoadPicture(imgOpen)
      .Caption = imgOpen
   End With
   frm.Show
   MDI.ActiveForm.FitToImage
   MDI.ActiveForm.UpdateUndo
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      PopupMenu mnuPop
   End If
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
   setBrowseButtons
End Sub

Private Sub Form_Load()
   Dim retVal
   MDI.MousePointer = 11
   File1.Pattern = "*gif;*.jpg;*.bmp"
   retVal = Mid(App.Path, 1, 3)
   If lastBrowseDrive <> "" Then
      Drive1.Drive = lastBrowseDrive
   Else
      Drive1.Drive = retVal
   End If
   If lastBrowseDir <> "" Then
      Dir1.Path = lastBrowseDir
   Else
      Dir1.Path = App.Path & "\ADMDrawPics"
   End If
   File1.Path = Dir1.Path
   'For some reason, this code does not work on initial opening in IDE
   'but works okay in compiled exe.
   MDI.MousePointer = 0
End Sub

Private Sub loadControlArray()
   Dim i
   On Error GoTo errChk
   For i = 1 To File1.ListCount '- 1
       Load tImage(i)
       Load tShape(i)
   Next i
   Exit Sub
   
errChk:
If Err.Number = 360 Then
   Resume Next
Else
   MsgBox "In loadControlArray: " & Err.Number & " - " & Err.Description
   Resume Next
End If
End Sub

Private Sub loadThumbs()
   Dim ratio As Double
   Dim i As Integer
   On Error GoTo chkErr
   frmBrowse.Show
   Me.Caption = "Browse: " & i & " Of " & File1.ListCount & " Thumbnails Loaded"
   ReDim Preserve tPathArray(0)
   tPathArray(0) = Dir1.Path & "\" & File1.List(0)
   loadControlArray
   For i = 1 To File1.ListCount '- 1
       ReDim Preserve tPathArray(i)
       tPathArray(i) = Dir1.Path & "\" & File1.List(i)
   Next i
   positionThumbShapes
   pBar.Visible = True
   For i = 0 To File1.ListCount - 1
       tImage(i).Picture = LoadPicture(tPathArray(i))
       ratio = tImage(i).Picture.Width / tImage(i).Picture.Height
       If tImage(i).Picture.Height >= tImage(i).Picture.Width Then
          tImage(i).Height = tShape(i).Height - (2 * tShape(i).BorderWidth)
          tImage(i).Width = (tShape(i).Height * ratio) - (2 * tShape(i).BorderWidth)
       Else
          tImage(i).Width = tShape(i).Width - (2 * tShape(i).BorderWidth)
          tImage(i).Height = (tShape(i).Width / ratio) - (2 * tShape(i).BorderWidth)
       End If
       tImage(i).Left = tShape(i).Left + (tShape(i).Width - tImage(i).Width) / 2
       tImage(i).Top = tShape(i).Top + (tShape(i).Height - tImage(i).Height) / 2
       tImage(i).Visible = True
       Me.Caption = "Browse: " & (i + 1) & " Of " & File1.ListCount & " Thumbnails Loaded"
       pBar.Value = ((i + 1) / File1.ListCount) * 100
   Next i
   pBar.Value = 0
   pBar.Visible = False
   Exit Sub
   
chkErr:
Resume Next
End Sub

Private Sub unloadThumbs()
   Dim i As Integer
   For i = 1 To tImage.UBound
       Unload tImage(i)
       Unload tShape(i)
   Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   canSplit = True
   xSplit = x
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(3).Text = ""
   Me.MousePointer = 9
   If canSplit Then
      If x <= 0 Or x >= Me.ScaleWidth - (tShape(0).Width + 10) - VScroll1.Width - 3 Then Exit Sub
      Drive1.Width = x
      Dir1.Width = x
      File1.Width = x
      picOuter.Width = Me.ScaleWidth - Drive1.Width - VScroll1.Width - 3
      picOuter.Left = x + 3
      picHolder.Width = picOuter.Width
   End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   canSplit = False
   Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   unloadThumbs
   Unload Me
End Sub

Private Sub setScrollSize()
   VScroll1.Min = 0
   VScroll1.Max = picOuter.ScaleHeight - picHolder.Height
   VScroll1.LargeChange = picHolder.ScaleHeight
   VScroll1.SmallChange = picHolder.ScaleHeight / 5
End Sub

Private Sub Form_Resize()
   On Error GoTo chkErr
   VScroll1.Move Me.ScaleWidth - VScroll1.Width, 0, VScroll1.Width, Me.ScaleHeight '- pBar.Height
   Drive1.Move 0, 0
   Dir1.Move 0, Drive1.Height, Drive1.Width
   File1.Move 0, Drive1.Height + Dir1.Height, Drive1.Width, Me.ScaleHeight - Dir1.Height - Drive1.Height + 11
   picOuter.Move Drive1.Width + 3, 0, Me.ScaleWidth - Drive1.Width - VScroll1.Width - 3, Me.ScaleHeight '- pBar.Height
   pBar.Move Drive1.Width + 3, picOuter.Height - pBar.Height, Me.ScaleWidth - Drive1.Width - VScroll1.Width - 3, pBar.Height
   picHolder.Move 0, 0, picOuter.Width, picOuter.Height
   setScrollSize
   ThumbRowsRequired
   If rowsRequired <> tRowsSet Then
      picHolder.Visible = False
      loadControlArray
      Call ReThumb
      picHolder.Visible = True
   End If
   Exit Sub
   
chkErr:
Resume Next
End Sub

Private Sub mnuDelete_Click()
   Dim i, fSel As Boolean, ret, idx
   For i = 0 To File1.ListCount - 1
       If File1.Selected(i) = True Then
          fSel = True
          idx = i
       End If
   Next i
   If fSel Then
      ret = MsgBox("Okay to delete from hard drive, the file: " & Dir1.Path & "\" & File1.List(idx), vbYesNo)
      If ret = vbYes Then
         deleteFile = Dir1.Path & "\" & File1.List(idx)
         Kill Dir1.Path & "\" & File1.List(idx)
         File1.Refresh
         Call deleteOne
      End If
   Else
      MsgBox "Select a File first"
   End If
End Sub

Private Sub mnuWindowsOpt_Click(Index As Integer)
   Select Case Index
   Case 0
        MDI.Arrange 0 'Cascade
   Case 1
        MDI.Arrange 2 'Tile Vertically
   Case 2
        MDI.Arrange 1 'Tile Horizonally
   Case 3
        MDI.Arrange 3 'Arrange Icons
   Case 4
        'Seperator
   Case 5  'Close all windows
        Dim chkme As Form
        For Each chkme In Forms
            If chkme.Caption <> "ADMDraw" Then
                Unload chkme
            End If
        Next chkme
   End Select
End Sub

Private Sub picHolder_DblClick()
   picHolder.Refresh
End Sub

Private Sub picHolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MousePointer = 0
   MDI.StatusBar1.Panels(3).Text = ""
End Sub

Private Sub picOuter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MousePointer = 0
End Sub

Private Sub tImage_Click(Index As Integer)
   Dim i
   For i = 0 To File1.ListCount - 1
       tShape(i).Visible = False
   Next i
   tShape(Index).Visible = True
   File1.Selected(Index) = True
End Sub

Private Sub tImage_DblClick(Index As Integer)
   Dim chkGif
   Dim frm As New frmImage
   '==================
   ' See is this window is already open and bring it to top.
   Dim chkme As Form
   For Each chkme In Forms
       If LCase(chkme.Caption) = LCase(tPathArray(Index)) Then
          chkme.WindowState = vbNormal
          chkme.SetFocus
          Exit Sub
       End If
   Next chkme
   '==========' ====Get Gif options========
   chkGif = InStrRev(LCase(tPathArray(Index)), ".gif")
   If chkGif = Len(tPathArray(Index)) - 3 Then
      frmImage.ComDia.FileName = tPathArray(Index)
      Call frmImage.setGifOptions
   End If
   '============
   With frm
      .picImage.AutoRedraw = True
      .picImage.AutoSize = True
      .picImage.Picture = LoadPicture(tPathArray(Index))
      .Caption = tPathArray(Index)
   End With
   frm.Show
   MDI.ActiveForm.FitToImage
   MDI.ActiveForm.UpdateUndo
   Exit Sub
   
errChk:
If Err.Number = 53 Then
   MsgBox "Image has been moved or deleted. Close frmBrowse and reopen it to refresh file list."
   Exit Sub
Else
   MsgBox Err.Number & " - " & Err.Description
End If
End Sub

Private Sub VScroll1_Change()
   picHolder.Top = VScroll1.Value
End Sub

Public Sub ThumbRowsRequired()
   tSpaceRequired = (10 + tShape(0).Width) '10 is margin :tSpace is square so X & Y are same
   'How many thumbnail will fit on a row?
   TperRow = picHolder.Width \ tSpaceRequired
   'How many rows required?
   rowsRequired = (File1.ListCount \ TperRow) + 1
End Sub

Private Sub ReThumb()
   Dim i
   Me.MousePointer = 11
   Call positionThumbShapes
   For i = 0 To File1.ListCount - 1
       tImage(i).Left = tShape(i).Left + (tShape(i).Width - tImage(i).Width) / 2
       tImage(i).Top = tShape(i).Top + (tShape(i).Height - tImage(i).Height) / 2
       tImage(i).Visible = True
   Next i
   Me.MousePointer = 0
End Sub

Public Sub positionThumbShapes()
   Dim i, r, c, startPosX, startPosY
   Call ThumbRowsRequired
   tRowsSet = rowsRequired
   On Error GoTo errChk
   For r = 1 To rowsRequired
       For c = 1 To TperRow
           If i = File1.ListCount Then GoTo setHeight
           tShape(i).Move startPosX, startPosY
           '======
           tImage(i).Left = tShape(i).Left + (tShape(i).Width - tImage(i).Width) / 2
           tImage(i).Top = tShape(i).Top + (tShape(i).Height - tImage(i).Height) / 2
           '=======
           startPosX = startPosX + tSpaceRequired
           i = i + 1
       Next c
       startPosX = 0
       startPosY = startPosY + tSpaceRequired
   Next r
   
setHeight:
picHolder.Height = rowsRequired * tSpaceRequired
picHolder.ScaleHeight = picHolder.Height
picHolder.Move 0, 0, picOuter.Width, picHolder.Height
setScrollSize
Exit Sub

errChk:
If Err.Number = 340 Then
   MsgBox "positionThumbShapes bad i= " & i
   Resume Next
End If
End Sub

Private Sub VScroll1_Scroll()
   picHolder.Top = VScroll1.Value
End Sub

Public Sub deleteOne()
   Dim i, j
   For i = 0 To UBound(tPathArray)
       If tPathArray(i) = deleteFile Then
          For j = i To UBound(tPathArray) - 1
              tImage(j).Width = tImage(j + 1).Width
              tImage(j).Height = tImage(j + 1).Height
              tImage(j) = tImage(j + 1)
              tShape(j) = tShape(j + 1)
              tImage(j).Left = tShape(j).Left + (tShape(j).Width - tImage(j).Width) / 2
              tImage(j).Top = tShape(j).Top + (tShape(j).Height - tImage(j).Height) / 2
              tPathArray(j) = tPathArray(j + 1)
          Next j
          Unload tImage(UBound(tPathArray))
          Unload tShape(UBound(tPathArray))
          ReDim Preserve tPathArray(UBound(tPathArray) - 1)
          Call positionThumbShapes
          Me.Caption = "Browse: " & File1.ListCount & " Of " & File1.ListCount & " Thumbnails Loaded"
          Me.Refresh
          Exit For
       End If
   Next i
   picHolder.Refresh
End Sub

Public Sub addOne()
   Dim ratio As Double
   Dim i, j
   File1.Refresh
   While tImage.UBound + 1 < File1.ListCount
      Load tImage(UBound(tPathArray) + 1)
      tImage(UBound(tPathArray) + 1).Visible = True
      Load tShape(UBound(tPathArray) + 1)
   Wend
   While UBound(tPathArray) < File1.ListCount - 1
      ReDim Preserve tPathArray(UBound(tPathArray) + 1)
   Wend
   For i = 0 To UBound(tPathArray)
       If Dir1.Path & "\" & File1.List(i) <> tPathArray(i) Then
          For j = UBound(tPathArray) - 1 To i Step -1
              tPathArray(j + 1) = tPathArray(j)
              tImage(j + 1).Picture = tImage(j).Picture
              tShape(j + 1) = tShape(j)
              tImage(j + 1).Width = tImage(j).Width
              tImage(j + 1).Height = tImage(j).Height
              tImage(j + 1).Left = tShape(j).Left + (tShape(j).Width - tImage(j).Width) / 2
              tImage(j + 1).Top = tShape(j).Top + (tShape(j).Height - tImage(j).Height) / 2
          '==========
          Next j
          tPathArray(i) = Dir1.Path & "\" & File1.List(i)
          '========
          tImage(i) = LoadPicture(Dir1.Path & "\" & File1.List(i))
          ratio = tImage(i).Picture.Width / tImage(i).Picture.Height
          If tImage(i).Picture.Height >= tImage(i).Picture.Width Then
             tImage(i).Height = tShape(i).Height - (2 * tShape(i).BorderWidth)
             tImage(i).Width = (tShape(i).Height * ratio) - (2 * tShape(i).BorderWidth)
          Else
             tImage(i).Width = tShape(i).Width - (2 * tShape(i).BorderWidth)
             tImage(i).Height = (tShape(i).Width / ratio) - (2 * tShape(i).BorderWidth)
          End If
          tImage(i).Visible = True
          Me.Caption = "Browse: " & File1.ListCount & " Of " & File1.ListCount & " Thumbnails Loaded"
          '==========
          Call ReThumb
          DoEvents
          Me.Refresh
          Exit For
       End If
   Next i
   picHolder.Refresh
End Sub

Public Sub saveOne()
   Dim i
   For i = 0 To UBound(tPathArray)
       If tPathArray(i) = saveFile Then
          tImage(i).Picture = LoadPicture(saveFile)
          Me.Refresh
       End If
   Next i
   picHolder.Refresh
End Sub
