Attribute VB_Name = "SimplePaints"
Option Explicit
Public OriginalSelX As Integer
Public OriginalSelY As Integer
Public PrevX As Integer
Public PrevY As Integer
Public Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Sub DrawAirBrush(Target As Long, x As Single, y As Single, Radius As Long, NumberOfSteps As Long)
   Dim cx As Long 'X counter
   Dim cy As Long 'Y counter
   Dim TempColor As Long
   Dim TempRadius As Integer
   Dim i As Integer
   Dim red As Long
   Dim green As Long
   Dim blue As Long
   Dim Done() As Boolean
   ReDim Done(-Radius To Radius, -Radius To Radius)
   For i = 1 To NumberOfSteps
       TempRadius = Radius / NumberOfSteps * i
       For cx = -TempRadius To TempRadius
           For cy = -TempRadius To TempRadius
               If Not Done(cx, cy) Then
                  If (cx * cx) + (cy * cy) <= TempRadius * TempRadius Then
                     TempColor = GetPixel(Target, cx + x, cy + y)
                     SetPixelV Target, cx + x, cy + y, GetAirColor(TempColor, NumberOfSteps * (11 - val(MDI.txtP.Text) / 10), NumberOfSteps * (10 - val(MDI.txtP.Text) / 10) + CLng(i))
                     Done(cx, cy) = True
                  End If
               End If
           Next cy
       Next cx
   Next i
End Sub

Public Function GetAirColor(Color As Long, Max As Long, Position As Long) As Long
   Dim C1(2) As Byte
   Dim C2(2) As Byte
   Dim i As Integer
   Dim RS As Double, GS As Double, BS As Double
   Dim r As Double, g As Double, b As Double
   Dim Red1 As Double, Blue1 As Double, Green1 As Double
   Dim Red2 As Double, Blue2 As Double, Green2 As Double
   If Max <= 0 Then
      GetAirColor = Color
      Exit Function
   End If
   b = MDI.mseFore.BackColor \ 65536
   g = (MDI.mseFore.BackColor - b * 65536) \ 256
   r = MDI.mseFore.BackColor - b * 65536 - g * 256
   Red1 = r
   Green1 = g
   Blue1 = b
   Blue2 = Color \ 65536
   Green2 = (Color - Blue2 * 65536) \ 256
   Red2 = Color - Blue2 * 65536 - Green2 * 256
   If Red1 <> Red2 Then
      RS = ((Red1 - Red2) / Max)
   Else
      RS = 0
   End If
   If Green1 <> Green2 Then
      GS = ((Green1 - Green2) / Max)
   Else
      GS = 0
   End If
   If Blue1 <> Blue2 Then
      BS = ((Blue1 - Blue2) / Max)
   Else
      BS = 0
   End If
   r = r - RS * (Position)
   g = g - GS * (Position)
   b = b - BS * (Position)
   If r < 0 Then r = 0
   If r > 255 Then r = 255
   If g < 0 Then g = 0
   If g > 255 Then g = 255
   If b < 0 Then b = 0
   If b > 255 Then b = 255
   GetAirColor = RGB(CInt(r), CInt(g), CInt(b))
End Function

Public Sub DrawBlur(Target As Long, x As Single, y As Single, Radius As Long, NumberOfSteps As Long)
   Dim cx As Long 'X counter
   Dim cy As Long 'Y counter
   Dim TempColor(8) As Long
   Dim TempRadius As Integer
   Dim i As Integer
   Dim u As Integer
   Dim red(3) As Long
   Dim green(3) As Long
   Dim blue(3) As Long
   Dim Color As Long
   Dim Done() As Boolean
   ReDim Done(-Radius To Radius, -Radius To Radius)
   For i = 1 To NumberOfSteps
       TempRadius = Radius / NumberOfSteps * i
       For cx = -TempRadius To TempRadius 'Step 2
           For cy = -TempRadius To TempRadius 'Step 2
               If Not Done(cx, cy) Then
                  If (cx * cx) + (cy * cy) <= TempRadius * TempRadius Then
                     TempColor(0) = GetPixel(Target, cx + x, cy + y)
                     TempColor(1) = GetPixel(Target, cx + x, cy + y - 1)
                     TempColor(2) = GetPixel(Target, cx + x - 1, cy + y - 1)
                     TempColor(3) = GetPixel(Target, cx + x - 1, cy + y)
                     For u = 0 To 3
                         blue(u) = TempColor(u) \ 65536
                         green(u) = (TempColor(u) - blue(u) * 65536) \ 256
                         red(u) = TempColor(u) - blue(u) * 65536 - green(u) * 256
                         If red(u) < 0 Then red(u) = 0
                         If red(u) > 255 Then red(u) = 255
                         If green(u) < 0 Then green(u) = 0
                         If green(u) > 255 Then green(u) = 255
                         If blue(u) < 0 Then blue(u) = 0
                         If blue(u) > 255 Then blue(u) = 255
                     Next u
                     Color = RGB((red(0) + red(1) + red(2) + red(3)) / 4, (green(0) + green(1) + green(2) + green(3)) / 4, (blue(0) + blue(1) + blue(2) + blue(3)) / 4)
                     SetPixelV Target, cx + x, cy + y, Color
                     SetPixelV Target, cx + x, cy + y - 1, Color
                     SetPixelV Target, cx + x - 1, cy + y - 1, Color
                     SetPixelV Target, cx + x - 1, cy + y, Color
                     Done(cx, cy) = True
                  End If
               End If
           Next cy
       Next cx
   Next i
   PrevX = x
   PrevY = y
End Sub

