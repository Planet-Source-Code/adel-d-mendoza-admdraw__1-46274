Attribute VB_Name = "mBrightContrast"
Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Declare Function ShellExecuteEX Lib "shell32.dll" Alias _
"ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

' VB Pixel library  version 1.2
' Written by Mike D Sutton of EDais

'About:
'A library of useful functions to deal with colour in VB
'
'You use this code at your own risk, I don't accept any
' responsibility for anything nasty it may do to your machine!
'
'Please don't rip my work off...  I'm distributing this library
' free of charge because I think it can help other developers,
' this doesn't give you the right to take credit for it.  By
' all means use it, yes, but please don't claim it's your own
' work or charge for it.

Public Type Pixel  '24-Bit colour
   blue As Integer
   green As Integer
   red As Integer
   Trans As Integer '<- Uncomment this for 32-Bit colour (All functions will still work)
End Type

' * * * * * * * * * * * *
' Construction functions
' * * * * * * * * * * * *
Public Function NewPixel(inRed As Integer, inGreen As Integer, inBlue As Integer) As Pixel
   With NewPixel 'Pixel constructor
      .red = inRed
      .green = inGreen
      .blue = inBlue
   End With
End Function

Public Function PixToLong(inPix As Pixel) As Long
   With inPix 'Converts a pixel to long
      PixToLong = RGB(.red, .green, .blue)
   End With
End Function

Public Function LongToPix(inCol As Long) As Pixel
   With LongToPix 'Converts a long to pixel
      .red = inCol And &HFF
      .green = (inCol \ &H100) And &HFF
      .blue = (inCol \ &H10000) And &HFF
   End With
End Function

Public Function GreyscalePix(inVal As Integer) As Pixel
   'Returns a 24/32-Bit greyscale 'colour'
   GreyscalePix = NewPixel(inVal, inVal, inVal)
End Function

' * * * * * * * * *
' Colour adjustment
' * * * * * * * * *
Public Function InvertPix(inPix As Pixel) As Pixel
   With InvertPix 'Inverts a pixel value
      .red = Not inPix.red
      .green = Not inPix.green
      .blue = Not inPix.blue
   End With
End Function

Public Function SameCol(inPixA As Pixel, inPixB As Pixel) As Boolean
   'Checks to see if two colours are the same
   SameCol = (inPixA.red = inPixB.red) And (inPixA.green = inPixB.green) And (inPixA.blue = inPixB.blue)
End Function

Public Function GetGreyPix(inPix As Pixel) As Integer
   'Returns the 8-Bit greyscale value of the pixel
   GetGreyPix = ((inPix.red * 0.222) + (inPix.green * 0.707) + (inPix.blue * 0.071))
End Function

Public Function LightenPix(inCol As Pixel, inAmt As Integer) As Pixel
   'Lightens the colour by a specified amount
   LightenPix = NewPixel(CheckHighByte(inCol.red + inAmt), _
     CheckHighByte(inCol.green + inAmt), CheckHighByte(inCol.blue + inAmt))
End Function

Public Function DarkenPix(inCol As Pixel, inAmt As Integer) As Pixel
   'Darkens the colour by a specified amount
   DarkenPix = NewPixel(CheckLowByte(inCol.red - inAmt), _
     CheckLowByte(inCol.green - inAmt), CheckLowByte(inCol.blue - inAmt))
End Function

' * * * * * * * * * *
' Blending functions
' * * * * * * * * * *
Public Function TransPix(PixA As Pixel, PixB As Pixel, inAmt As Single) As Pixel
   With TransPix 'Linearly interpolate one pixel to another
      .red = LinearB(PixA.red, PixB.red, inAmt)
      .green = LinearB(PixA.green, PixB.green, inAmt)
      .blue = LinearB(PixA.blue, PixB.blue, inAmt)
   End With
End Function

Public Function TransAddPix(inPixA As Pixel, inPixB As Pixel, inAmt As Single) As Pixel
   TransAddPix = TransPix(BlendAdd(inPixA, inPixB), inPixB, inAmt) 'Additive transparency
End Function

Public Function TransSubPix(inPixA As Pixel, inPixB As Pixel, inAmt As Single) As Pixel
   TransSubPix = TransPix(BlendSub(inPixA, inPixB), inPixB, inAmt) 'Subtractive transparency
End Function

Public Function TransLightPix(inPixA As Pixel, inPixB As Pixel, inAmt As Single) As Pixel
   TransLightPix = TransPix(BlendLight(inPixA, inPixB), inPixB, inAmt) 'Lighten transparency
End Function

Public Function TransDarkPix(inPixA As Pixel, inPixB As Pixel, inAmt As Single) As Pixel
   TransDarkPix = TransPix(BlendDark(inPixA, inPixB), inPixB, inAmt) 'Darken transparency
End Function

Public Function TransDiffPix(inPixA As Pixel, inPixB As Pixel, inAmt As Single) As Pixel
   TransDiffPix = TransPix(BlendDiff(inPixA, inPixB), inPixB, inAmt) 'Difference transparency
End Function

Public Function TransScrnPix(inPixA As Pixel, inPixB As Pixel, inAmt As Single) As Pixel
   TransScrnPix = TransPix(BlendScrn(inPixA, inPixB), inPixB, inAmt) 'Screen transparency
End Function

Public Function TransExclPix(inPixA As Pixel, inPixB As Pixel, inAmt As Single) As Pixel
   TransExclPix = TransPix(BlendExcl(inPixA, inPixB), inPixB, inAmt) 'Exclusion transparency
End Function

Public Function BlendAdd(inPixA As Pixel, inPixB As Pixel) As Pixel
   With BlendAdd 'Blend with additive blend mode
      .red = CheckHighByte(inPixA.red + inPixB.red)
      .green = CheckHighByte(inPixA.green + inPixB.green)
      .blue = CheckHighByte(inPixA.blue + inPixB.blue)
   End With
End Function

Public Function BlendSub(inPixA As Pixel, inPixB As Pixel) As Pixel
   With BlendSub 'Blend with subtractive blend mode
      .red = CheckLowByte(inPixA.red - inPixB.red)
      .green = CheckLowByte(inPixA.green - inPixB.green)
      .blue = CheckLowByte(inPixA.blue - inPixB.blue)
   End With
End Function

Public Function BlendLight(inPixA As Pixel, inPixB As Pixel) As Pixel
   With BlendLight 'Blend with lighten blend mode
      .red = MaxB(inPixA.red, inPixB.red)
      .green = MaxB(inPixA.green, inPixB.green)
      .blue = MaxB(inPixA.blue, inPixB.blue)
   End With
End Function

Public Function BlendDark(inPixA As Pixel, inPixB As Pixel) As Pixel
   With BlendDark 'Blend with darken blend mode
      .red = MinB(inPixA.red, inPixB.red)
      .green = MinB(inPixA.green, inPixB.green)
      .blue = MinB(inPixA.blue, inPixB.blue)
   End With
End Function

Public Function BlendDiff(inPixA As Pixel, inPixB As Pixel) As Pixel
   With BlendDiff 'Blend with difference blend mode
      .red = Abs(inPixA.red - inPixB.red)
      .green = Abs(inPixA.green - inPixB.green)
      .blue = Abs(inPixA.blue - inPixB.blue)
   End With
End Function

Public Function BlendScrn(inPixA As Pixel, inPixB As Pixel) As Pixel
   With BlendScrn 'Blend with screen blend mode
      .red = CheckHighByte(inPixA.red * (1 + (inPixB.red / 255)))
      .green = CheckHighByte(inPixA.green * (1 + (inPixB.green / 255)))
      .blue = CheckHighByte(inPixA.blue * (1 + (inPixB.blue / 255)))
   End With
End Function

Public Function BlendExcl(inPixA As Pixel, inPixB As Pixel) As Pixel
   With BlendExcl 'Blend with exculusion blend mode
      .red = LinearB(inPixA.red, Not inPixA.red, inPixB.red / 255)
      .green = LinearB(inPixA.green, Not inPixA.green, inPixB.green / 255)
      .blue = LinearB(inPixA.blue, Not inPixA.blue, inPixB.blue / 255)
   End With
End Function


' * * * * * * * *
' Misc functions
' * * * * * * * *
Function LinearB(inValA As Integer, inValB As Integer, inPos As Single) As Integer
   'Linear interpolation routine (Byte version)
   LinearB = (inValA * (1 - inPos)) + (inValB * inPos)
End Function

Function CheckHighByte(inVal As Integer) As Integer
   'Makes sure a value will fit in a byte (Only checks for >255)
   CheckHighByte = IIf(inVal > 255, 255, inVal)
End Function

Function CheckLowByte(inVal As Integer) As Integer
   'Makes sure a value will fit in a byte (Only checks for <0)
   CheckLowByte = IIf(inVal < 0, 0, inVal)
End Function

Function CheckByte(inVal As Integer) As Integer
   'Makes sure a value will fit in a byte (Checks for both <0 and >255)
   If inVal < 0 Then CheckByte = 0 Else If inVal > 255 Then CheckByte = 255 Else CheckByte = inVal
End Function

Public Function MaxB(inValA As Integer, inValB As Integer) As Integer
   'Returns the maximum of two values
   MaxB = IIf(inValA > inValB, inValA, inValB)
End Function

Public Function MinB(inValA As Integer, inValB As Integer) As Integer
   'Returns the minumum of two values
   MinB = IIf(inValA < inValB, inValA, inValB)
End Function

Public Sub ShowProps(FileName As String, OwnerhWnd As Long)
   Dim SEI As SHELLEXECUTEINFO
   Dim r As Long
   With SEI
      .cbSize = Len(SEI)
      .fMask = SEE_MASK_NOCLOSEPROCESS Or _
      SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
      .hwnd = OwnerhWnd
      .lpVerb = "Properties"
      .lpFile = FileName
      .lpParameters = vbNullChar
      .lpDirectory = vbNullChar
      .nShow = 0
      .hInstApp = 0
      .lpIDList = 0
   End With
   r = ShellExecuteEX(SEI)
End Sub

Public Function Contrast(pix As Pixel, iAmnt As Integer) As Pixel
   If pix.red < 128 Then
      pix.red = CheckLowByte(pix.red + iAmnt)
   Else
      pix.red = CheckHighByte(pix.red - iAmnt)
   End If
   If pix.blue < 128 Then
      pix.blue = CheckLowByte(pix.blue + iAmnt)
   Else
      pix.blue = CheckHighByte(pix.blue - iAmnt)
   End If
   If pix.green < 128 Then
      pix.green = CheckLowByte(pix.green + iAmnt)
   Else
      pix.green = CheckHighByte(pix.green - iAmnt)
   End If
   Contrast = pix
End Function

Public Sub Bright(val As Integer, oPic As PictureBox)
   On Error Resume Next
   'variables for brightness, color calculation, positioning
   Dim Brightness As Single
   Dim NewColor As Long
   Dim x, y As Integer
   Dim r, g, b As Integer
   'change the brightness to a percent
   Brightness = val '/ 100
   'run a loop through the picture to change every pixel
   For x = 0 To oPic.ScaleWidth
       frmBrightContrast.PB.Value = (x / oPic.ScaleWidth) * 100
       For y = 0 To oPic.ScaleHeight
           'get the current color value
           NewColor = GetPixel(oPic.hDc, x, y)
           'extract the R,G,B values from the long returned by GetPixel
           r = (NewColor Mod 256)
           b = (Int(NewColor / 65536))
           g = ((NewColor - (b * 65536) - r) / 256)
           'change the RGB settings to their appropriate brightness
           r = r + Brightness
           b = b + Brightness
           g = g + Brightness
           'make sure the new variables aren't too high or too low
           If r > 255 Then r = 255
           If r < 0 Then r = 0
           If b > 255 Then b = 255
           If b < 0 Then b = 0
           If g > 255 Then g = 255
           If g < 0 Then g = 0
           'set the new pixel
           SetPixelV oPic.hDc, x, y, RGB(r, g, b)
           'continue through the loop
       Next y
       'refresh the picture box every 10 lines (a nice progress bar effect)
       If x Mod 10 = 0 Then oPic.Refresh
   Next x
   'final picture refresh
   oPic.Refresh
   frmBrightContrast.PB.Value = 0
End Sub

