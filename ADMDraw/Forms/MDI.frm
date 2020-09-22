VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "ADMDraw"
   ClientHeight    =   8160
   ClientLeft      =   285
   ClientTop       =   165
   ClientWidth     =   11625
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   3  'Align Left
      Height          =   7305
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   12885
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   1065
      _CBHeight       =   7305
      _Version        =   "6.0.8450"
      MinHeight1      =   1005
      Width1          =   7245
      NewRow1         =   0   'False
      Begin VB.CommandButton Command2 
         Caption         =   "<->"
         Height          =   195
         Left            =   645
         TabIndex        =   25
         Top             =   120
         Width           =   375
      End
      Begin VB.ComboBox cboRetouch 
         Height          =   315
         Left            =   60
         TabIndex        =   24
         Text            =   "Soften"
         Top             =   5625
         Visible         =   0   'False
         Width           =   960
      End
      Begin MSComCtl2.UpDown UpDownS 
         Height          =   285
         Left            =   600
         TabIndex        =   23
         Top             =   6255
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtS"
         BuddyDispid     =   196611
         OrigLeft        =   60
         OrigTop         =   135
         OrigRight       =   300
         OrigBottom      =   165
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDownP 
         Height          =   285
         Left            =   615
         TabIndex        =   22
         Top             =   6810
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtP"
         BuddyDispid     =   196612
         OrigLeft        =   60
         OrigTop         =   135
         OrigRight       =   300
         OrigBottom      =   165
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtS 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   165
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "10"
         Top             =   6255
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   165
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "100"
         Top             =   6795
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   105
         TabIndex        =   21
         Text            =   "Pressure"
         Top             =   6600
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   105
         TabIndex        =   20
         Text            =   "Steps"
         Top             =   6030
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton cmdJump 
         Caption         =   "Jump"
         Height          =   240
         Left            =   255
         TabIndex        =   16
         ToolTipText     =   "Last Used Setting"
         Top             =   4920
         Width           =   555
      End
      Begin MSComDlg.CommonDialog ComDia 
         Left            =   345
         Top             =   1815
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame FrColor 
         Caption         =   "Color"
         Height          =   1080
         Left            =   105
         TabIndex        =   11
         Top             =   2910
         Width           =   885
         Begin VB.Label lblB 
            Height          =   210
            Left            =   360
            TabIndex        =   15
            Top             =   660
            Width           =   435
         End
         Begin VB.Label lblG 
            Height          =   210
            Left            =   375
            TabIndex        =   14
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblR 
            Height          =   210
            Left            =   375
            TabIndex        =   13
            Top             =   285
            Width           =   435
         End
         Begin VB.Label lblRGB 
            Caption         =   "R-  G-  B-"
            Height          =   645
            Left            =   75
            TabIndex        =   12
            Top             =   270
            Width           =   210
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox PicDrawWidth 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   810
         Left            =   165
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   10
         Top             =   4080
         Width           =   810
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   5205
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDrawWidth"
         BuddyDispid     =   196622
         OrigLeft        =   60
         OrigTop         =   135
         OrigRight       =   300
         OrigBottom      =   165
         Max             =   50
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtDrawWidth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   285
         TabIndex        =   8
         Text            =   "4"
         Top             =   5160
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Caption         =   "More"
         Height          =   270
         Left            =   195
         TabIndex        =   7
         ToolTipText     =   "Left or Right Click for Back Color"
         Top             =   2535
         Width           =   690
      End
      Begin VB.PictureBox mseFore 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   225
         ScaleHeight     =   330
         ScaleWidth      =   315
         TabIndex        =   5
         ToolTipText     =   "Double Click to Change"
         Top             =   195
         Width           =   375
      End
      Begin VB.PictureBox mseColor 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   105
         MousePointer    =   99  'Custom
         Picture         =   "MDI.frx":1CFA
         ScaleHeight     =   1755
         ScaleWidth      =   870
         TabIndex        =   1
         ToolTipText     =   "Left or Right Click to Pick Color"
         Top             =   750
         Width           =   870
      End
      Begin VB.PictureBox mseBack 
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   480
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   6
         ToolTipText     =   "Double Click to Change"
         Top             =   345
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList ImageListVert 
      Left            =   1890
      Top             =   3435
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6DAC
            Key             =   "magnify"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7A8C
            Key             =   "brush"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":876C
            Key             =   "erase"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":8BC0
            Key             =   "line"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":8D24
            Key             =   "flood"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":9A04
            Key             =   "spray"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A6E4
            Key             =   "text"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B3C4
            Key             =   "rect"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B528
            Key             =   "rectempty"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":C208
            Key             =   "circle"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":C36C
            Key             =   "circleempty"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":C4D0
            Key             =   "dropper"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":D1B0
            Key             =   "retouch"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":DE04
            Key             =   "ants"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":EAE4
            Key             =   "lasso"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListHorz 
      Left            =   1965
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":EC48
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":F928
            Key             =   "new"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":FA8C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1076C
            Key             =   "saveas"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1144C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1212C
            Key             =   "browse"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":12580
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":13260
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":13F40
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":14C20
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":15900
            Key             =   "print"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":165E0
            Key             =   "canvas"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":172C0
            Key             =   "pastenew"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":17FA0
            Key             =   "crop"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":18C80
            Key             =   "selectall"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":18DE4
            Key             =   "selectnone"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbVert 
      Align           =   3  'Align Left
      Height          =   660
      Left            =   1065
      TabIndex        =   3
      Top             =   600
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageListVert"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "magnify"
            Object.ToolTipText     =   "Zoom"
            ImageKey        =   "magnify"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "brush"
            Object.ToolTipText     =   "Brush"
            ImageKey        =   "brush"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "erase"
            Object.ToolTipText     =   "Eraser"
            ImageKey        =   "erase"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "flood"
            Object.ToolTipText     =   "Flood"
            ImageKey        =   "flood"
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "spray"
            Object.ToolTipText     =   "Spray Can"
            ImageKey        =   "spray"
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "text"
            Object.ToolTipText     =   "Text"
            ImageKey        =   "text"
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "line"
            Object.ToolTipText     =   "Line"
            ImageKey        =   "line"
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "rect"
            Object.ToolTipText     =   "Rectangle"
            ImageKey        =   "rect"
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "rectempty"
            Object.ToolTipText     =   "Rectangle Outline"
            ImageKey        =   "rectempty"
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "circle"
            Object.ToolTipText     =   "Circle"
            ImageKey        =   "circle"
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "circleempty"
            Object.ToolTipText     =   "Circle Outline"
            ImageKey        =   "circleempty"
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dropper"
            Object.ToolTipText     =   "Pick Color"
            ImageKey        =   "dropper"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "retouch"
            Object.ToolTipText     =   "Retouch"
            ImageKey        =   "retouch"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "lasso"
            Object.ToolTipText     =   "Random Select"
            ImageKey        =   "lasso"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ants"
            Object.ToolTipText     =   "Select Area"
            ImageKey        =   "ants"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbHorz 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageListHorz"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Description     =   "New File"
            Object.ToolTipText     =   "New"
            Object.Tag             =   "all"
            ImageKey        =   "new"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Description     =   "Open File"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   "all"
            ImageKey        =   "open"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Description     =   "Save File"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   "form"
            ImageKey        =   "save"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "saveas"
            Object.ToolTipText     =   "Save As"
            Object.Tag             =   "form"
            ImageKey        =   "saveas"
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "all"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "browse"
            Description     =   "Browse Directory"
            Object.ToolTipText     =   "Browse"
            Object.Tag             =   "all"
            ImageKey        =   "browse"
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print Preview"
            Object.Tag             =   "form"
            ImageKey        =   "print"
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "all"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   "all"
            ImageKey        =   "undo"
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "redo"
            Object.ToolTipText     =   "Redo"
            Object.Tag             =   "all"
            ImageKey        =   "redo"
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "all"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   "rect"
            ImageKey        =   "cut"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   "rect"
            ImageKey        =   "copy"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste in Active Image"
            Object.Tag             =   "formcb"
            ImageKey        =   "paste"
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pastenew"
            Object.ToolTipText     =   "Paste in New Image"
            Object.Tag             =   "all"
            ImageKey        =   "pastenew"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "all"
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "canvas"
            Object.ToolTipText     =   "Resize Canvas"
            Object.Tag             =   "form"
            ImageKey        =   "canvas"
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crop"
            Object.ToolTipText     =   "Crop Image to Selection"
            Object.Tag             =   "rect"
            ImageKey        =   "crop"
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "selectall"
            Object.ToolTipText     =   "Select All"
            Object.Tag             =   "form"
            ImageKey        =   "selectall"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "selectnone"
            Object.ToolTipText     =   "Select None"
            Object.Tag             =   "form"
            ImageKey        =   "selectnone"
         EndProperty
      EndProperty
      MouseIcon       =   "MDI.frx":19100
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   7905
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14870
            MinWidth        =   1879
            Object.ToolTipText     =   "Shows X,Y coordinates"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Size WxH"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PB 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuRect 
         Caption         =   "Rectangle"
      End
      Begin VB.Menu mnuOval 
         Caption         =   "Oval"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowColor()
   r = clr Mod 256
   g = (clr \ 256) Mod 256
   b = clr \ 256 \ 256
   lblR.Caption = r
   lblG.Caption = g
   lblB.Caption = b
End Sub
Private Sub HideColor()
   lblR.Caption = ""
   lblG.Caption = ""
   lblB.Caption = ""
End Sub

Private Sub cboRetouch_Click()
   MDI.ActiveForm.picImage.AutoRedraw = True
   MDI.ActiveForm.picImage.AutoSize = False
   If cboRetouch.Text = "Brightness" Or cboRetouch.Text = "Contrast" Then
      frmImage.shCircle.Visible = False
      frmImage.shSquare.Visible = False
      frmBrightContrast.Picture2 = MDI.ActiveForm.picImage.Picture
      frmBrightContrast.Show
      Text1.Visible = False
      Text2.Visible = False
      txtS.Visible = False
      UpDownS.Visible = False
      txtP.Visible = False
      UpDownP.Visible = False
   ElseIf cboRetouch.Text = "Glitter" Then
      frmImage.shCircle.Visible = False
      frmImage.shSquare.Visible = False
      Text1.Visible = False
      Text2.Visible = False
      txtS.Visible = False
      UpDownS.Visible = False
      txtP.Visible = False
      UpDownP.Visible = False
   Else
      If cboRetouch.Text <> "Clone" Then
         frmImage.shCircle.Visible = True
         Text1.Visible = True
         Text2.Visible = True
         txtS.Visible = True
         UpDownS.Visible = True
         txtP.Visible = True
         UpDownP.Visible = True
      Else
         frmImage.shSquare.Visible = True
         Text1.Visible = False
         Text2.Visible = False
         txtS.Visible = False
         UpDownS.Visible = False
         txtP.Visible = False
         UpDownP.Visible = False
      End If
   End If
End Sub

Private Sub cmdJump_Click()
   If set1 <> "" And set2 <> "" Then
      MDI.txtDrawWidth.Text = set1
      set1 = set2
      set2 = MDI.txtDrawWidth.Text
   End If
   PicDrawWidth.SetFocus
End Sub

Private Sub cmdJump_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Works like 'Jump' on TV remote control. Jumps between last two settings."
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ComDia.CancelError = True
   On Error GoTo ErrHandler
   ComDia.Flags = cdlCCFullOpen
   ComDia.ShowColor
   If ComDia.Color = RGB(197, 197, 197) Then ComDia.Color = RGB(196, 196, 196)
   If Button = 1 Then
      mseFore.BackColor = ComDia.Color
   Else
      mseBack.BackColor = ComDia.Color
   End If
   Exit Sub
   
ErrHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   HideColor
   MDI.StatusBar1.Panels(1).Text = "Shows Color wheel to select color."
End Sub

Private Sub Command2_Click()
   Dim temp
   temp = mseFore.BackColor
   mseFore.BackColor = mseBack.BackColor
   mseBack.BackColor = temp
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Switch Left Button and Right Button Colors."
End Sub

Private Sub CoolBar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   HideColor
   MDI.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub FrColor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   HideColor
   StatusBar1.Panels(1).Text = "Color under mouse pointer."
End Sub

Private Sub MDIForm_Initialize()
   Const ATTR_DIRECTORY = 16
   OkNewHW = True
   If Dir$(App.Path & "\ADMDrawPics", ATTR_DIRECTORY) <> "" Then
      '("The ADMDrawPics directory exist")
   Else
      MkDir App.Path & "\ADMDrawPics"
      '("The ADMDrawPics directory does not exist")
   End If
   If Command$ <> "" Then
      FileDrag = True
      frmImage.ComDia.FileName = Command$
      frmImage.FileOpen
      If MDI.ActiveForm Is Nothing Then Exit Sub
      MDI.ActiveForm.FitToImage
      MDI.ActiveForm.setHorzButtons
   End If
End Sub

Private Sub MDIForm_Load()
   'VB has a mind of it's own
   'The unloads are a reminder to VB that I want to be in control
   Unload frmNewOptions
   frmImage.setHorzButtons
   Unload frmImage
   PicDrawWidth.DrawWidth = val(txtDrawWidth.Text)
   PicDrawWidth.PSet (25, 25), vbBlack

   OkNewHW = True
   set1 = "1" 'set draw width
   set2 = "4"
   cboRetouch.AddItem "Soften"
   cboRetouch.AddItem "Clone"
   cboRetouch.AddItem "Glitter"
   cboRetouch.AddItem "Brightness"
   cboRetouch.AddItem "Contrast"
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next 'Error occurs if active form is frmBrowse
   HideColor
   StatusBar1.Panels(1).Text = ""
   If MDI.ActiveForm Is Nothing Then Exit Sub
   MDI.ActiveForm.shCircle.Visible = False
   MDI.ActiveForm.shSquare.Visible = False
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   OkNewHW = True
   Unload frmBrowse
   Unload frmImage
   Unload frmNewOptions
   Unload frmText
End Sub

Private Sub mnuExit_Click()
   MDIForm_QueryUnload 0, 0
   End
End Sub

Private Sub mnuFile_Click()
   If Not OkNewHW Then Exit Sub
End Sub

Private Sub mnuNew_Click()
   If Not OkNewHW Then Exit Sub
   frmImage.FileNew
End Sub

Private Sub mnuOpen_Click()
   If Not OkNewHW Then Exit Sub
   frmImage.FileOpen
   If MDI.ActiveForm Is Nothing Then Exit Sub
   MDI.ActiveForm.FitToImage
   MDI.ActiveForm.setHorzButtons
End Sub

Private Sub mnuOval_Click()
   MDI.ActiveForm.shRect.Shape = 2
End Sub

Private Sub mnuRect_Click()
   MDI.ActiveForm.shRect.Shape = 0
End Sub

Private Sub mseBack_DblClick()
   Call Command1_MouseDown(2, 0, 0, 0)
End Sub

Private Sub mseBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   clr = mseColor.Point(x, y)
   ShowColor
   MDI.StatusBar1.Panels(1).Text = "Color produced by Right Button."
End Sub

Private Sub mseColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 Then
      mseFore.BackColor = mseColor.Point(x, y)
   Else
      mseBack.BackColor = mseColor.Point(x, y)
   End If
End Sub

Private Sub mseColor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   mseColor.MouseIcon = frmImage.ImageListCursors.ListImages(3).Picture
   clr = mseColor.Point(x, y)
   ShowColor
   MDI.StatusBar1.Panels(1).Text = "Set Mouse Color."
End Sub

Private Sub mseFore_DblClick()
   Call Command1_MouseDown(1, 0, 0, 0)
End Sub

Private Sub mseFore_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   clr = mseColor.Point(x, y)
   ShowColor
   MDI.StatusBar1.Panels(1).Text = "Color produced by Left Button."
End Sub

Private Sub PicDrawWidth_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Graphical representation of draw width."
End Sub

Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   StatusBar1.Panels(1).Text = ""
End Sub

Private Sub tbHorz_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Not OkNewHW Then Exit Sub
   Select Case Button.Key
   Case "new"
      frmImage.FileNew
      If cancelIt Then GoTo exSub
      MDI.ActiveForm.FitToImage
      MDI.ActiveForm.setHorzButtons
   Case "open"
      frmImage.FileOpen
      If cancelIt Then GoTo exSub
      If MDI.ActiveForm Is Nothing Then GoTo exSub
      MDI.ActiveForm.FitToImage
      MDI.ActiveForm.setHorzButtons
   Case "save"
      If MDI.ActiveForm Is Nothing Then GoTo exSub
      MDI.ActiveForm.FileSave
   Case "saveas"
      If MDI.ActiveForm Is Nothing Then GoTo exSub
      MDI.ActiveForm.FileSaveAs
   Case "browse"
      ' See if Browse is already open and bring it to top.
      Dim ret
      Dim chkme As Form
      For Each chkme In Forms
          ret = Mid(chkme.Caption, 1, 6)
            
          If ret = "Browse" Then
             chkme.WindowState = vbNormal
             chkme.SetFocus
             Button.Value = tbrUnpressed
             Exit Sub
          End If
      Next chkme
      Load frmBrowse
      frmBrowse.Show
      frmBrowse.setBrowseButtons
   Case "print"
      'Load frmPrint
      frmPrint.picPrint.Picture = LoadPicture()
      frmPrint.picPrint.Width = 12240
      frmPrint.picPrint.Height = 15840
      frmPrint.picZoom.Width = 12240
      frmPrint.picZoom.Height = 15840
      'Test if picImage will fit on page without resizing
      Dim wid, hgt, aspect
      wid = MDI.ActiveForm.picImage.ScaleX(MDI.ActiveForm.picImage.Width, vbPixels, vbTwips)
      hgt = MDI.ActiveForm.picImage.ScaleY(MDI.ActiveForm.picImage.Height, vbPixels, vbTwips)
      aspect = wid / hgt
      If wid < 10800 And hgt < 14400 Then GoTo PrintIt
      If wid > 10800 And wid < 14400 And hgt < 10800 Then
         frmPrint.picPrint.Width = 15840
         frmPrint.picPrint.Height = 12240
         frmPrint.picZoom.Width = 15840
         frmPrint.picZoom.Height = 12240
         GoTo PrintIt
      End If
      If wid >= 14400 Or hgt >= 10800 Then
         If hgt >= wid Then
            hgt = 14400
            wid = hgt * aspect
         Else
            wid = 10800
            hgt = wid / aspect
         End If
         If wid > hgt Then
            frmPrint.picPrint.Width = 15840
            frmPrint.picPrint.Height = 12240
            frmPrint.picZoom.Width = 15840
            frmPrint.picZoom.Height = 12240
         End If
      End If
    
PrintIt:
frmPrint.picPrint.PaintPicture MDI.ActiveForm.picImage.Picture, ((frmPrint.picPrint.Width / 2) - ((wid) / 2)), ((frmPrint.picPrint.Height / 2) - (hgt / 2)), (wid), hgt
frmPrint.picZoom.PaintPicture MDI.ActiveForm.picImage.Picture, ((frmPrint.picZoom.Width / 2) - ((wid) / 2)), ((frmPrint.picZoom.Height / 2) - (hgt / 2)), (wid), hgt
frmPrint.Show
'MDI.ActiveForm.printPic

   Case "undo"
      MDI.ActiveForm.DoUnDo
   Case "redo"
      MDI.ActiveForm.DoReDo
   Case "cut"
      MDI.ActiveForm.FileCut
   Case "copy"
      MDI.ActiveForm.FileCopy
   Case "paste"
      MDI.ActiveForm.FilePasteNewSelection
   Case "pastenew"
      If Clipboard.GetFormat(vbCFBitmap) Then
         Dim frm As New frmImage
         frm.Show
         MDI.ActiveForm.setHorzButtons
         With frm
              .picImage.AutoSize = True
              .picImage.AutoRedraw = True
         End With
         MDI.ActiveForm.picImage.Picture = Clipboard.GetData(vbCFBitmap)
         MDI.ActiveForm.FitToImage
         MDI.ActiveForm.UpdateUndo
         MDI.ActiveForm.setDirty
         Num = Num + 1
         MDI.ActiveForm.Caption = "Image " & Num
      Else
         MsgBox "No Image on Clipboard."
         MDI.tbHorz.Buttons(pasteNewID).Visible = False
      End If
   Case "canvas"
      If MDI.ActiveForm Is Nothing Then GoTo exSub
      MDI.ActiveForm.changeImageSize
      MDI.ActiveForm.FitToImage
      MDI.ActiveForm.UpdateUndo
   Case "crop"
      If MDI.ActiveForm Is Nothing Then GoTo exSub
      MDI.ActiveForm.FileCrop
      MDI.ActiveForm.UpdateUndo
      MDI.ActiveForm.setDirty
   Case "selectall"
      MDI.ActiveForm.FileSelectAll
   Case "selectnone"
      MDI.ActiveForm.FileSelectNone
   Case Else
      MsgBox "Not yet coded."
   End Select
   Button.Value = tbrUnpressed
   Exit Sub
   
exSub:
Button.Value = tbrUnpressed
End Sub

Private Sub tbHorz_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next 'Error occurs if active form is frmBrowse
   HideColor
   StatusBar1.Panels(1).Text = ""
   If MDI.ActiveForm Is Nothing Then Exit Sub
   MDI.ActiveForm.shCircle.Visible = False
   MDI.ActiveForm.shSquare.Visible = False
End Sub

Private Sub tbVert_ButtonClick(ByVal Button As MSComctlLib.Button)
   'MsgBox Button.Key
   If MDI.ActiveForm Is Nothing Then
      MsgBox "Open New or Existing File First"
      GoTo exSub
   End If
   If Not OkNewHW Then GoTo exSub
   On Error GoTo exSub
   setSwitchesFalse

   Select Case Button.Key
   Case "magnify"
      MagnifyIt = True
      MDI.ActiveForm.setPointer
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "brush"
      DrawIt = True
      MDI.ActiveForm.setPointer
   Case "erase"
      EraseIt = True
      MDI.ActiveForm.setPointer
   Case "flood"
      FloodIt = True
      MDI.ActiveForm.setPointer
   Case "spray"
      SprayIt = True
      txtP.Text = "50"
      Text1.Visible = True
      Text2.Visible = True
      txtS.Visible = True
      UpDownS.Visible = True
      txtP.Visible = True
      UpDownP.Visible = True
      MDI.ActiveForm.setPointer
   Case "text"
      TextIt = True
      MDI.ActiveForm.setPointer
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "line"
      LineIt = True
      MDI.ActiveForm.setPointer
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "rect"
      RectIt = True
      MDI.ActiveForm.setPointer
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "rectempty"
      RectEmptyIt = True
      MDI.ActiveForm.setPointer
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "circle"
      CircleIt = True
      MDI.ActiveForm.setPointer
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "circleempty"
      CircleEmptyIt = True
      MDI.ActiveForm.setPointer
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "dropper"
      PickColor = True
      MDI.ActiveForm.setPointer
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "retouch"
      RetouchIt = True
      Text1.Visible = True
      Text2.Visible = False
      txtS.Visible = True
      UpDownS.Visible = True
      txtP.Visible = False
      UpDownP.Visible = False
      MDI.cboRetouch.Visible = True
      MDI.ActiveForm.setPointer
   Case "lasso"
      LassoIt = True
      MDI.ActiveForm.setPointer
      MsgBox "Not yet coded."
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   Case "ants"
      AntsIt = True
      MDI.ActiveForm.setPointer
   Case Else
      MsgBox "Not yet coded."
      MDI.ActiveForm.Timer1.Enabled = False
      MDI.ActiveForm.shRect.Visible = False
   End Select
   Exit Sub
   
exSub:
Button.Value = tbrUnpressed
End Sub

Private Sub tbVert_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next 'Error occurs if active form is frmBrowse
   StatusBar1.Panels(1).Text = ""
   If MDI.ActiveForm Is Nothing Then Exit Sub
   MDI.ActiveForm.shCircle.Visible = False
   MDI.ActiveForm.shSquare.Visible = False
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Steps for use in applicable tools."
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Pressure for use in applicable tools."
End Sub

Private Sub txtDrawWidth_Change()
   On Error GoTo errMsg
   If val(txtDrawWidth.Text) > 50 Or val(txtDrawWidth.Text) < 1 Then GoTo errMsg
      PicDrawWidth.Cls
      PicDrawWidth.DrawWidth = val(txtDrawWidth.Text)
      PicDrawWidth.PSet (25, 25), vbBlack
      Exit Sub
      
errMsg:
MsgBox "Value must be between 1 and 50"
txtDrawWidth.Text = 25
End Sub

Private Sub txtDrawWidth_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Draw width size in pixels."
End Sub

Private Sub txtP_Change()
   On Error GoTo errMsg
   If val(txtP.Text) > 100 Or val(txtP.Text) < 1 Then GoTo errMsg
   Exit Sub
   
errMsg:
MsgBox "Value must be between 1 and 100"
txtP.Text = 100
End Sub

Private Sub txtP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Pressure for use in applicable tools."
End Sub

Private Sub txtS_Change()
   On Error GoTo errMsg
   If val(txtS.Text) > 100 Or val(txtS.Text) < 1 Then GoTo errMsg
   Exit Sub
   
errMsg:
MsgBox "Value must be between 1 and 100"
txtS.Text = 100
End Sub

Private Sub txtS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Steps for use in applicable tools."
End Sub

Private Sub UpDown1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MDI.StatusBar1.Panels(1).Text = "Up/Down for Draw width size in pixels."
End Sub

