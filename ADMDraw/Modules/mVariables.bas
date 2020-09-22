Attribute VB_Name = "mVariables"
Option Explicit
Public TFileSize, Tsuffit As String
Public AR 'Aspect Ratio
Public FileDrag As Boolean
Public deleteFile As String, addFile As String, saveFile As String
Public lastSave, lastBrowseDir, lastBrowseDrive
Public pasteNewID, redoID, undoID 'Index of button in tbHorz
Public set1 As String, set2 As String
Public r As Integer, g As Integer, b As Integer, clr As Long
Public drawAreaWidth, drawAreaHeight
Public Num As Integer, curX, curY
Public OkNewHW As Boolean
Public cancelIt As Boolean
Public PasteIt As Boolean, canPaste As Boolean
Public MoveIt As Boolean, canMove As Boolean
Public MagnifyIt As Boolean
Public DrawIt As Boolean, canDraw As Boolean
Public EraseIt As Boolean, canErase As Boolean
Public FloodIt As Boolean
Public SprayIt As Boolean, canSpray As Boolean
Public TextIt As Boolean, canText As Boolean
Public LineIt As Boolean, canLine As Boolean
Public RectIt As Boolean, canRect As Boolean
Public RectEmptyIt As Boolean, canRectEmpty As Boolean
Public CircleIt As Boolean, canCircle As Boolean
Public CircleEmptyIt As Boolean, canCircleEmpty As Boolean
Public RetouchIt As Boolean, canRetouch As Boolean, canClone As Boolean
Public LassoIt As Boolean, canLasso As Boolean
Public AntsIt As Boolean, canAnts As Boolean
Public PickColor As Boolean
Public Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hDc As Long, ByVal nXStart As Long, ByVal nYStart As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Public Sub setSwitchesFalse()
   MDI.ActiveForm.picTemp.Visible = False  'just extra precaution
   PasteIt = False
   MoveIt = False
   MagnifyIt = False
   DrawIt = False
   EraseIt = False
   FloodIt = False
   SprayIt = False
   TextIt = False
   LineIt = False
   RectIt = False
   RectEmptyIt = False
   CircleIt = False
   CircleEmptyIt = False
   RetouchIt = False
   canClone = False
   PickColor = False
   LassoIt = False
   AntsIt = False
   MDI.tbHorz.Buttons(12).Visible = False 'cut button
   MDI.tbHorz.Buttons(13).Visible = False 'copy button
   MDI.tbHorz.Buttons(14).Visible = False 'paste button
   MDI.Text1.Visible = False
   MDI.Text2.Visible = False
   MDI.txtS.Visible = False
   MDI.UpDownS.Visible = False
   MDI.txtP.Visible = False
   MDI.UpDownP.Visible = False
   MDI.cboRetouch.Visible = False
   MDI.ActiveForm.picImage.AutoRedraw = True
   MDI.ActiveForm.picImage.AutoSize = False
   MDI.ActiveForm.shRed.Visible = False
End Sub


