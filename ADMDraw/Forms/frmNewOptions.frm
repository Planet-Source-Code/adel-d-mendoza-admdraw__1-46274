VERSION 5.00
Begin VB.Form frmNewOptions 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "New Form Options"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1350
      TabIndex        =   6
      Top             =   1590
      Width           =   795
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fill Drawing Area"
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   1170
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Okay"
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Top             =   1590
      Width           =   705
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   330
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "400"
      Top             =   645
      Width           =   495
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   330
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "400"
      Top             =   225
      Width           =   510
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2055
      Left            =   60
      Top             =   75
      Width           =   2475
   End
   Begin VB.Label lblHeight 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Height in Pixels"
      Height          =   375
      Left            =   915
      TabIndex        =   4
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Width in Pixels"
      Height          =   300
      Left            =   930
      TabIndex        =   2
      Top             =   270
      Width           =   1185
   End
End
Attribute VB_Name = "frmNewOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Sub Check1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0 'stops beep
      Check1.Value = vbChecked
   End If
End Sub

Private Sub cmdCancel_Click()
   OkNewHW = True
   cancelIt = True
   Check1.Caption = "Fill Drawing Area"
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   If val(txtWidth.Text) < 1 Or val(txtHeight.Text) < 1 Then
      MsgBox "Width and Height must be 1 or greater."
      Exit Sub
   End If
   OkNewHW = True
   Check1.Caption = "Fill Drawing Area"
   Me.Hide
End Sub

Private Sub Form_Activate()
   txtWidth.SetFocus
   txtWidth.SelStart = 0
   txtWidth.SelLength = 3
End Sub

Private Sub Form_Load()
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub txtHeight_Change()
   'Check1.Value = 0 'unchecked
End Sub

Private Sub txtHeight_GotFocus()
   txtHeight.SelStart = 0
   txtHeight.SelLength = 3
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0 'stops beep
      cmdOk.SetFocus
   End If
End Sub

Private Sub txtHeight_LostFocus()
   If Check1.Caption = "Keep Aspect Ratio" Then
      If Check1.Value = 1 Then
         txtWidth.Text = Int(val(txtHeight.Text) * AR)
      End If
   End If
End Sub

Private Sub txtWidth_Change()
   'Check1.Value = 0 'unchecked
End Sub

Private Sub txtWidth_GotFocus()
   txtWidth.SelStart = 0
   txtWidth.SelLength = 3
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0 'stops beep
      txtHeight.SetFocus
   End If
End Sub

Private Sub txtWidth_LostFocus()
   If Check1.Caption = "Keep Aspect Ratio" Then
      If Check1.Value = 1 Then
         txtHeight.Text = Int(val(txtWidth.Text) / AR)
      End If
   End If
End Sub
