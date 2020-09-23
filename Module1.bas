Attribute VB_Name = "Module1"
Public Mem As Picture
Public R&(), G&(), B&(), Color&
Public xx%, yy%
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Sub Colorwait(CW As Boolean)
Form1.Label1.Caption = "Reading colors - Please wait"
Form1.Label1.Visible = CW
DoEvents
End Sub

Public Sub Imagewait(IW As Boolean)
Form1.Label1.Caption = "Processing Image - Please wait"
Form1.Label1.Visible = IW
DoEvents
End Sub

'read colors from pic1 and store them
'as R, G and B in an array
'This sub is called every time before an effect is processed, but
'normally this schould be read only once or when the source-picture changes

Public Sub ReadColor(Rx1%, Ry1%, Rx2%, Ry2%)
Form1.Pic4.Picture = Form1.Pic1.Picture
Colorwait True
On Error Resume Next
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
Color = GetPixel(Form1.Pic1.hdc, xx, yy)
R(xx, yy) = Color Mod 256&
G(xx, yy) = ((Color And &HFF00) / 256&) Mod 256&
B(xx, yy) = (Color And &HFF0000) / 65536
Next yy
Next xx
Colorwait False
End Sub

'Gradient with alpha blending
'First set the gradient in the destination-picture
'Read the colors of the destination-picture and mix then with the
'colors of the source-picture (with alpha-blend).
'Put the new colors in the destination-picture
'Note: this sub can be made shorter and faster

Public Sub GradientCol(Ob As Object, AB As Single, R1 As Single, G1 As Single, B1 As Single, R2%, G2%, B2%)

On Error Resume Next 'just in case

Dim H%, rt As Single, Gt As Single, Bt As Single
Imagewait True
AB = AB / 10 'alpha blending
H = Ob.Height - 1
rt = (R2 - R1) / H
Gt = (G2 - G1) / H
Bt = (B2 - B1) / H
'Set the gradient
For xx = 0 To H
Ob.Line (0, xx)-(Ob.Width - 1, xx), RGB(R1, G1, B1)
R1 = R1 + rt
G1 = G1 + Gt
B1 = B1 + Bt
Next xx
'Read the gradient-colors, mix the with alpha-blend
'and put the new colors back.
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
    Color = GetPixel(Ob.hdc, xx, yy)
    R1 = Color Mod 256&
    G1 = ((Color And &HFF00) / 256&) Mod 256&
    B1 = (Color And &HFF0000) / 65536
'This is the actual alpha-blending
        R(xx, yy) = (R(xx, yy) * (1 - AB)) + (R1 * AB)
        G(xx, yy) = (G(xx, yy) * (1 - AB)) + (G1 * AB)
        B(xx, yy) = (B(xx, yy) * (1 - AB)) + (B1 * AB)
'put the new colors back
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

'Add or substract color; just color 1 is working
'True = add color
'False = substract color
Public Sub AddColor(Ob As Object, R1%, G1%, B1%, P As Boolean)
On Error Resume Next 'just in case
Imagewait True
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
    If P = True Then 'add colors
        R(xx, yy) = R(xx, yy) + R1
        G(xx, yy) = G(xx, yy) + G1
        B(xx, yy) = B(xx, yy) + B1
    Else 'P = False: substract colors
        R(xx, yy) = R(xx, yy) - R1
        G(xx, yy) = G(xx, yy) - G1
        B(xx, yy) = B(xx, yy) - B1
    End If
If R(xx, yy) > 255 Then R(xx, yy) = 255
If R(xx, yy) < 0 Then R(xx, yy) = 0
If G(xx, yy) > 255 Then G(xx, yy) = 255
If G(xx, yy) < 0 Then G(xx, yy) = 0
If B(xx, yy) > 255 Then B(xx, yy) = 255
If B(xx, yy) < 0 Then B(xx, yy) = 0
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

'Brighten or darken picture
'Brighten: read the color and add a percentage of its value to it
'then put the color back
'Darken: substract a percentage
Public Sub BrightenPicture(Ob As Object, Strength%)
Dim Am%
On Error Resume Next
Imagewait True
If Strength > 0 Then
Am = 11 - Strength 'brighten
Else
Am = -11 - Strength 'darken
End If
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
    Color = Abs((R(xx, yy) + G(xx, yy) + B(xx, yy))) / Am
        R(xx, yy) = R(xx, yy) + Color
        G(xx, yy) = G(xx, yy) + Color
        B(xx, yy) = B(xx, yy) + Color
            If R(xx, yy) > 255 Then R(xx, yy) = 255
            If G(xx, yy) > 255 Then G(xx, yy) = 255
            If B(xx, yy) > 255 Then B(xx, yy) = 255
            If R(xx, yy) < 0 Then R(xx, yy) = 0
            If G(xx, yy) < 0 Then G(xx, yy) = 0
            If B(xx, yy) < 0 Then B(xx, yy) = 0
    
    SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
    Next yy
    Next xx
Imagewait False
Ob.Refresh
End Sub

'Kill red, green or blue
Public Sub KillColor(Ob As Object, Cc%)
On Error Resume Next
Imagewait True
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
If Cc = 0 Then SetPixel Ob.hdc, xx, yy, RGB(0, G(xx, yy), B(xx, yy)) 'kill red
If Cc = 1 Then SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), 0, B(xx, yy)) 'kill green
If Cc = 2 Then SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), 0) 'kill blue
Next yy, xx
Ob.Refresh
Imagewait False
End Sub

'Swap red, green or blue
'if cc=0 then swap red & green
'if cc=1 then swap red & blue
'if cc=2 then swap green & blue
Public Sub SwapColor(Ob As Object, Cc%)
On Error Resume Next
Imagewait True
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
If Cc = 0 Then SetPixel Ob.hdc, xx, yy, RGB(G(xx, yy), R(xx, yy), B(xx, yy))
If Cc = 1 Then SetPixel Ob.hdc, xx, yy, RGB(B(xx, yy), G(xx, yy), R(xx, yy))
If Cc = 2 Then SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), B(xx, yy), G(xx, yy))
Next yy
Next xx
Imagewait False
Ob.Refresh
End Sub

'Negative color (photo-negative)
'This is easy done by an exclusive Or ---> original color XOR 255
'Cc=0 = only negative red
'Cc=1 = only negative green
'Cc=2 = only negative blue
'Cc=3 = negative red and green and blue (photo-negative)
Public Sub NegativeColor(Ob As Object, Cc%)
On Error Resume Next
Imagewait True
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
If Cc = 0 Then SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy) Xor 255, G(xx, yy), B(xx, yy))
If Cc = 1 Then SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy) Xor 255, B(xx, yy))
If Cc = 2 Then SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy) Xor 255)
If Cc = 3 Then SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy) Xor 255, G(xx, yy) Xor 255, B(xx, yy) Xor 255)
Next yy, xx
Ob.Refresh
Imagewait False
End Sub

'Greyscale --> turns picture into greyed colors
'It's not wise to do it like this:
'   newcolor = (R + G + B) / 3
'The human eye reacts differently to the red, green and blue components of a color.
'The formula used in this sub is more "real".
Public Sub GreyColor(Ob As Object) 'grey
On Error Resume Next
Imagewait True
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
'first make the red component
R(xx, yy) = R(xx, yy) * 0.3 + G(xx, yy) * 0.59 + B(xx, yy) * 0.11
If R(xx, yy) > 255 Then R(xx, yy) = 255
If R(xx, yy) < 0 Then R(xx, yy) = 0
'then equal green and blue to red
G(xx, yy) = R(xx, yy)
B(xx, yy) = R(xx, yy)
'put the new colors back
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
Next xx
Imagewait False
Ob.Refresh
End Sub

'Mix source picture (Ob) with another picture (Ob2)
'Alpha blending is in progress
Public Sub MixPic(Ob As Object, Ob2 As Object, AB As Single)
On Error Resume Next
Dim R1&, G1&, B1&
AB = AB / 10
Imagewait True
For xx = 0 To Ob2.Width - 1
For yy = 0 To Ob2.Height - 1
    Color = GetPixel(Ob2.hdc, xx, yy)
    R1 = Color Mod 256&
    G1 = ((Color And &HFF00) / 256&) Mod 256&
    B1 = (Color And &HFF0000) / 65536
'This is the actual alpha-blending
        R(xx, yy) = (R(xx, yy) * (1 - AB)) + (R1 * AB)
        G(xx, yy) = (G(xx, yy) * (1 - AB)) + (G1 * AB)
        B(xx, yy) = (B(xx, yy) * (1 - AB)) + (B1 * AB)
'put the new colors back
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

'Mix with pattern - alpha blending
'First, tile the pattern on the destination-picture
'Read the colors of the destination-picture and mix then with the
'colors of the source-picture (with alpha-blend).
'Put the new colors in the destination-picture
Public Sub MixPat(Ob As Object, Ob1 As Object, AB As Single)
Dim Rm, Gm, Bm
On Error Resume Next
AB = AB / 10
Imagewait True
For xx = 0 To Ob.Width / Ob1.Width
For yy = 0 To Ob.Height / Ob1.Height
Ob.PaintPicture Ob1, xx * Ob1.Width, yy * Ob1.Height, Ob1.Width, Ob1.Height
Next yy, xx
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
    Color = GetPixel(Ob.hdc, xx, yy)
    Rm = Color Mod 256&
    Gm = ((Color And &HFF00) / 256&) Mod 256&
    Bm = (Color And &HFF0000) / 65536
    R(xx, yy) = (R(xx, yy) * (1 - AB)) + (Rm * AB)
    G(xx, yy) = (G(xx, yy) * (1 - AB)) + (Gm * AB)
    B(xx, yy) = (B(xx, yy) * (1 - AB)) + (Bm * AB)
    SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
    Next yy, xx
Imagewait False
Ob.Refresh
End Sub

'Add picture with scaling & alpha-blend
'First put the picture on the destination-picturebox, according to the scale
'(from 10% - 100%)

'NOTE:
'The surrounding color of the picture to add is black (&H0, RGB(0,0,0), vbBlack)
'This works as the mask-color.
'To do this properly, the picture to add is .Gif or .Bmp, .Jpg isn't clean
'With .Gif, you can only have 255 colors, but it works fine...

'Read the colors of the destination-picture
'when this color <> 0 (black) then mix them with the
'colors of the source-picture (with alpha-blend).
'Put the new colors in the destination-picture
Public Sub AddPic(Ob As Object, Ob1 As Object, AB As Single, Sc As Single) 'add picture
On Error Resume Next
Dim Rm, Gm, Bm
Imagewait True
AB = AB / 10 'alpha blend
Sc = Sc / 10 'Scale
Ob.PaintPicture Ob1, (Ob.Width - (Ob1.Width * Sc)) / 2, (Ob.Height - (Ob1.Height * Sc)) / 2, Ob1.Width * Sc, Ob1.Height * Sc
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
    Color = GetPixel(Ob.hdc, xx, yy)
    If Color <> 0 Then
        Rm = Color Mod 256&
        Gm = ((Color And &HFF00) / 256&) Mod 256&
        Bm = (Color And &HFF0000) / 65536
    R(xx, yy) = (R(xx, yy) * (1 - AB)) + (Rm * AB)
    G(xx, yy) = (G(xx, yy) * (1 - AB)) + (Gm * AB)
    B(xx, yy) = (B(xx, yy) * (1 - AB)) + (Bm * AB)
    End If
    SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
    Next yy, xx
Imagewait False
Ob.Refresh
End Sub

'Add silhouette with scaling & alpha-blend
'First put the picture on the destination-picturebox, according to the scale
'(from 10% - 100%)

'NOTE:
'The surrounding color of the picture to add is black (&H0, RGB(0,0,0), vbBlack)
'This works as the mask-color.
'To do this properly, the picture to add is .Gif or .Bmp, .Jpg isn't clean
'With .Gif, you can only have 255 colors, but it works fine...

'Read the colors of the destination-picture
'when this color <> 0 (black) then make them black and mix with the
'colors of the source-picture (with alpha-blend).
'Put the new colors in the destination-picture
Public Sub AddSil(Ob As Object, Ob1 As Object, AB As Single, Sc As Single)
On Error Resume Next
Dim Rm, Gm, Bm
Imagewait True
AB = AB / 10
Sc = Sc / 10
Ob.PaintPicture Ob1, (Ob.Width - (Ob1.Width * Sc)) / 2, (Ob.Height - (Ob1.Height * Sc)) / 2, Ob1.Width * Sc, Ob1.Height * Sc
    For xx = (Ob.Width - (Ob1.Width * Sc)) / 2 To ((Ob.Width - (Ob1.Width * Sc)) / 2) + Ob1.Width * Sc
    For yy = (Ob.Height - (Ob1.Height * Sc)) / 2 To ((Ob.Height - (Ob1.Height * Sc)) / 2) + Ob1.Height * Sc
    Color = GetPixel(Ob.hdc, xx, yy)
    If Color <> 0 Then
        'make colors black
        Rm = 0
        Gm = 0
        Bm = 0
    R(xx, yy) = (R(xx, yy) * (1 - AB)) + (Rm * AB)
    G(xx, yy) = (G(xx, yy) * (1 - AB)) + (Gm * AB)
    B(xx, yy) = (B(xx, yy) * (1 - AB)) + (Bm * AB)
    End If
    SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
    Next yy, xx
Imagewait False
Ob.Refresh
End Sub

'Slide picture
'First put the second picture on the destination-picturebox
'set alpha blending to 0
'Mix with the source colors and put the new colors back.
'Increase alpha-blending
'At the end (most right) of the destination-picture, alpha blending is 1.
Public Sub SlidePic(Ob As Object, Ob1 As Object)
Dim Rm, Gm, Bm, S As Single, Alpha As Single
On Error Resume Next
S = 1 / Ob.Width
Alpha = 0
Imagewait True
Ob.PaintPicture Ob1, 0, 0, Ob.Width, Ob.Height
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
    Color = GetPixel(Ob.hdc, xx, yy)
    Rm = Color Mod 256&
    Gm = ((Color And &HFF00) / 256&) Mod 256&
    Bm = (Color And &HFF0000) / 65536
    R(xx, yy) = (R(xx, yy) * (1 - Sqr(Alpha))) + (Rm * Sqr(Alpha))
    G(xx, yy) = (G(xx, yy) * (1 - Sqr(Alpha))) + (Gm * Sqr(Alpha))
    B(xx, yy) = (B(xx, yy) * (1 - Sqr(Alpha))) + (Bm * Sqr(Alpha))
    SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
    Next yy
     Alpha = Alpha + S
        Next xx
    Ob.Refresh
Imagewait False
End Sub

Public Sub EmbossPicture(Ob As Object) 'emboss & emboss
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
        R(xx, yy) = (Abs(R(xx, yy) - R(xx + 1, yy + 1) + 128))
        G(xx, yy) = (Abs(G(xx, yy) - G(xx + 1, yy + 1) + 128))
        B(xx, yy) = (Abs(B(xx, yy) - B(xx + 1, yy + 1) + 128))
        SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub HoldRed(Ob As Object, Pct%) 'Hold red
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
If R(xx, yy) < Pct Then 'this can be in the range 1 - 128
R(xx, yy) = (Abs(R(xx, yy) - R(xx + 1, yy + 1) + 128))
G(xx, yy) = (Abs(G(xx, yy) - G(xx + 1, yy + 1) + 128))
B(xx, yy) = (Abs(B(xx, yy) - B(xx + 1, yy + 1) + 128))
End If
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub HoldGreen(Ob As Object, Pct%) 'Hold green & emboss
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
If G(xx, yy) < Pct Then 'this can be in the range 1 - 128
R(xx, yy) = (Abs(R(xx, yy) - R(xx + 1, yy + 1) + 128))
G(xx, yy) = (Abs(G(xx, yy) - G(xx + 1, yy + 1) + 128))
B(xx, yy) = (Abs(B(xx, yy) - B(xx + 1, yy + 1) + 128))
End If
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub HoldBlue(Ob As Object, Pct%) 'Hold blue & emboss
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
If B(xx, yy) < Pct Then 'this can be in the range 1 - 128
R(xx, yy) = (Abs(R(xx, yy) - R(xx + 1, yy + 1) + 128))
G(xx, yy) = (Abs(G(xx, yy) - G(xx + 1, yy + 1) + 128))
B(xx, yy) = (Abs(B(xx, yy) - B(xx + 1, yy + 1) + 128))
End If
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub BlurPicture(Ob As Object) 'blur
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = (Abs(R(xx - 1, yy - 1) + R(xx - 1, yy) + R(xx - 1, yy + 1) + R(xx, yy - 1) + R(xx, yy) + R(xx, yy + 1) + R(xx + 1, yy - 1) + R(xx + 1, yy) + R(xx + 1, yy + 1))) / 9
G(xx, yy) = (Abs(G(xx - 1, yy - 1) + G(xx - 1, yy) + G(xx - 1, yy + 1) + G(xx, yy - 1) + G(xx, yy) + G(xx, yy + 1) + G(xx + 1, yy - 1) + G(xx + 1, yy) + G(xx + 1, yy + 1))) / 9
B(xx, yy) = (Abs(B(xx - 1, yy - 1) + B(xx - 1, yy) + B(xx - 1, yy + 1) + B(xx, yy - 1) + B(xx, yy) + B(xx, yy + 1) + B(xx + 1, yy - 1) + B(xx + 1, yy) + B(xx + 1, yy + 1))) / 9
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub BlurPictureMore(Ob As Object) 'blur more
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = (Abs(R(xx - 2, yy - 2) + R(xx - 2, yy - 1) + R(xx - 2, yy) + R(xx - 2, yy + 1) + R(xx - 2, yy + 2) + R(xx - 1, yy - 2) + R(xx - 1, yy - 1) + R(xx - 1, yy) + R(xx - 1, yy + 1) + R(xx - 1, yy + 2) + R(xx, yy - 2) + R(xx, yy - 1) + R(xx, yy) + R(xx, yy + 1) + R(xx, yy + 2) + R(xx + 1, yy - 2) + R(xx + 1, yy - 1) + R(xx + 1, yy) + R(xx + 1, yy + 1) + R(xx + 1, yy + 2) + R(xx + 2, yy - 2) + R(xx + 2, yy - 1) + R(xx + 2, yy) + R(xx + 2, yy + 1) + R(xx + 2, yy + 2))) / 25
G(xx, yy) = (Abs(G(xx - 2, yy - 2) + G(xx - 2, yy - 1) + G(xx - 2, yy) + G(xx - 2, yy + 1) + G(xx - 2, yy + 2) + G(xx - 1, yy - 2) + G(xx - 1, yy - 1) + G(xx - 1, yy) + G(xx - 1, yy + 1) + G(xx - 1, yy + 2) + G(xx, yy - 2) + G(xx, yy - 1) + G(xx, yy) + G(xx, yy + 1) + G(xx, yy + 2) + G(xx + 1, yy - 2) + G(xx + 1, yy - 1) + G(xx + 1, yy) + G(xx + 1, yy + 1) + G(xx + 1, yy + 2) + G(xx + 2, yy - 2) + G(xx + 2, yy - 1) + G(xx + 2, yy) + G(xx + 2, yy + 1) + G(xx + 2, yy + 2))) / 25
B(xx, yy) = (Abs(B(xx - 2, yy - 2) + B(xx - 2, yy - 1) + B(xx - 2, yy) + B(xx - 2, yy + 1) + B(xx - 2, yy + 2) + B(xx - 1, yy - 2) + B(xx - 1, yy - 1) + B(xx - 1, yy) + B(xx - 1, yy + 1) + B(xx - 1, yy + 2) + B(xx, yy - 2) + B(xx, yy - 1) + B(xx, yy) + B(xx, yy + 1) + B(xx, yy + 2) + B(xx + 1, yy - 2) + B(xx + 1, yy - 1) + B(xx + 1, yy) + B(xx + 1, yy + 1) + B(xx + 1, yy + 2) + B(xx + 2, yy - 2) + B(xx + 2, yy - 1) + B(xx + 2, yy) + B(xx + 2, yy + 1) + B(xx + 2, yy + 2))) / 25
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

'diffusion rate from 50 to 500
Public Sub DiffusePicture(Ob As Object, Diffuse%) 'diffuse
Dim tt1%, tt%
On Error Resume Next
tt = Diffuse * 50
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
tt1 = (Rnd * tt) - 2
R(xx, yy) = Abs(R(xx, yy) + tt1)
G(xx, yy) = Abs(G(xx, yy) + tt1)
B(xx, yy) = Abs(B(xx, yy) + tt1)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub SharpenPicture(Ob As Object)  'sharpen
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = R(xx, yy) + 0.5 * (R(xx, yy) - R(xx - 1, yy - 1))
G(xx, yy) = G(xx, yy) + 0.5 * (G(xx, yy) - G(xx - 1, yy - 1))
B(xx, yy) = B(xx, yy) + 0.5 * (B(xx, yy) - B(xx - 1, yy - 1))
            If R(xx, yy) > 255 Then R(xx, yy) = 255
            If R(xx, yy) < 0 Then R(xx, yy) = 0
            If G(xx, yy) > 255 Then G(xx, yy) = 255
            If G(xx, yy) < 0 Then G(xx, yy) = 0
            If B(xx, yy) > 255 Then B(xx, yy) = 255
            If B(xx, yy) < 0 Then B(xx, yy) = 0
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub ErodePicture(Ob As Object, Pct%) 'erode
On Error Resume Next
Pct = Pct * 8
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = Abs(R(xx, yy) Xor Pct)
G(xx, yy) = Abs(G(xx, yy) Xor Pct)
B(xx, yy) = Abs(B(xx, yy) Xor Pct)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub BlowPicture(Ob As Object, Pct As Single) 'blow
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = Abs(R(xx, yy) Xor (R(xx, yy) / Pct))
G(xx, yy) = Abs(G(xx, yy) Xor (G(xx, yy) / Pct))
B(xx, yy) = Abs(B(xx, yy) Xor (B(xx, yy) / Pct))
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub ContrastPicture(Ob As Object, Pct%) 'contrast
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
If R(xx, yy) < 128 Then R(xx, yy) = R(xx, yy) - Int(R(xx, yy) / Pct)
If R(xx, yy) < 0 Then R(xx, yy) = 0
If G(xx, yy) < 128 Then G(xx, yy) = G(xx, yy) - Int(G(xx, yy) / Pct)
If G(xx, yy) < 0 Then G(xx, yy) = 0
If B(xx, yy) < 128 Then B(xx, yy) = B(xx, yy) - Int(B(xx, yy) / Pct)
If B(xx, yy) < 0 Then B(xx, yy) = 0
If R(xx, yy) >= 128 Then R(xx, yy) = R(xx, yy) + Int(R(xx, yy) / Pct)
If R(xx, yy) > 255 Then R(xx, yy) = 255
If G(xx, yy) >= 128 Then G(xx, yy) = G(xx, yy) + Int(G(xx, yy) / Pct)
If G(xx, yy) > 255 Then G(xx, yy) = 255
If B(xx, yy) >= 128 Then B(xx, yy) = B(xx, yy) + Int(B(xx, yy) / Pct)
If B(xx, yy) > 255 Then B(xx, yy) = 255
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub FogPicture(Ob As Object, Pct%) 'fog
Dim tt1%
On Error Resume Next
Imagewait True
Pct = Pct * 10
    For yy = 0 To Ob.Height - 1
tt1 = (Rnd * Pct) - 2
    For xx = 0 To Ob.Width - 1
R(xx, yy) = Abs(R(xx, yy) + tt1)
G(xx, yy) = Abs(G(xx, yy) + tt1)
B(xx, yy) = Abs(B(xx, yy) + tt1)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next xx, yy
Imagewait False
Ob.Refresh
End Sub

Public Sub AddNoise(Ob As Object) 'addnoise
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = ((Rnd * R(xx, yy)) + R(xx, yy)) / 2
G(xx, yy) = ((Rnd * G(xx, yy)) + G(xx, yy)) / 2
B(xx, yy) = ((Rnd * B(xx, yy)) + B(xx, yy)) / 2
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Freeze(Ob As Object, Pct As Single) 'freeze
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = Abs((R(xx, yy) - G(xx, yy) - B(xx, yy)) * Pct)
G(xx, yy) = Abs((G(xx, yy) - B(xx, yy) - R(xx, yy)) * Pct)
B(xx, yy) = Abs((B(xx, yy) - R(xx, yy) - G(xx, yy)) * Pct)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub BnW(Ob As Object, Pct%) 'B & W
Dim BWColor&
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
    If R(xx, yy) < Pct And G(xx, yy) < Pct And B(xx, yy) < Pct Then
    BWColor = 0
    Else
    BWColor = &HFFFFFF
    End If
SetPixel Ob.hdc, xx, yy, BWColor
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Brown(Ob As Object) 'brown
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = Abs(G(xx, yy) * B(xx, yy)) / 256
G(xx, yy) = Abs(B(xx, yy) * R(xx, yy)) / 256
B(xx, yy) = Abs(R(xx, yy) * G(xx, yy)) / 256
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Liquid(Ob As Object) 'liquid
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = ((G(xx, yy) - B(xx, yy)) ^ 2) / 125
G(xx, yy) = ((R(xx, yy) - B(xx, yy)) ^ 2) / 125
B(xx, yy) = ((R(xx, yy) - G(xx, yy)) ^ 2) / 125
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Yellow(Ob As Object) 'yellow
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
B(xx, yy) = ((G(xx, yy) - R(xx, yy)) ^ 2) / 125
R(xx, yy) = ((G(xx, yy) - B(xx, yy)) ^ 2) / 125
G(xx, yy) = ((B(xx, yy) + R(xx, yy)) ^ 2) / 125
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Charcoal(Ob As Object) 'charcoal
Dim tCol&
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
            R(xx, yy) = Abs(R(xx, yy) * (G(xx, yy) - B(xx, yy) + G(xx, yy) + R(xx, yy))) / 256
            G(xx, yy) = Abs(R(xx, yy) * (B(xx, yy) - G(xx, yy) + B(xx, yy) + R(xx, yy))) / 256
            B(xx, yy) = Abs(G(xx, yy) * (B(xx, yy) - G(xx, yy) + B(xx, yy) + R(xx, yy))) / 256
            tCol = RGB(R(xx, yy), G(xx, yy), B(xx, yy))
            R(xx, yy) = Abs(tCol Mod 256)
            G(xx, yy) = Abs((tCol \ 256) Mod 256)
            B(xx, yy) = Abs(tCol \ 256 \ 256)
            R(xx, yy) = (R(xx, yy) + G(xx, yy) + B(xx, yy)) / 3
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), R(xx, yy), R(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub DarkMoon(Ob As Object) 'dark moon
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = Abs(R(xx, yy) - 64)
G(xx, yy) = Abs(R(xx, yy) - 64)
B(xx, yy) = Abs(R(xx, yy) - 64)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub TotalEclipse(Ob As Object) 'eclipse
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = Abs(G(xx, yy) - 64)
G(xx, yy) = Abs(G(xx, yy) - 64)
B(xx, yy) = Abs(G(xx, yy) - 64)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub PurpleRain(Ob As Object) 'purple
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = Abs(G(xx, yy) + R(xx, yy) / 2)
G(xx, yy) = Abs(B(xx, yy) + G(xx, yy) / 2)
B(xx, yy) = Abs(R(xx, yy) + B(xx, yy) / 2)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Spooky(Ob As Object) 'Spooky
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
G(xx, yy) = Abs(R(xx, yy) + G(xx, yy) / 2)
B(xx, yy) = Abs(R(xx, yy) + B(xx, yy) / 2)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub UnReal(Ob As Object) 'unreal
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
If (G(xx, yy) = 0) Or (B(xx, yy) = 0) Then
    G(xx, yy) = 1
    B(xx, yy) = 1
End If
        R(xx, yy) = Abs(Sin(Atn(G(xx, yy) / B(xx, yy))) * 125 + 20)
        G(xx, yy) = Abs(Sin(Atn(R(xx, yy) / B(xx, yy))) * 125 + 20)
        B(xx, yy) = Abs(Sin(Atn(R(xx, yy) / G(xx, yy))) * 125 + 20)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Flame(Ob As Object) 'flame
Dim C As Long
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
    C = (R(xx, yy) + G(xx, yy) + B(xx, yy)) / 3
        If R(xx, yy) > B(xx, yy) Then
            R(xx, yy) = Abs(R(xx, yy) + C)
            B(xx, yy) = Abs(B(xx, yy) - C)
        End If
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Aquarel(Ob As Object) 'aquarel
On Error Resume Next
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
If R(xx, yy) < 128 And G(xx, yy) < 128 And B(xx, yy) < 128 Then
R(xx, yy) = 2 * R(xx, yy): G(xx, yy) = 2 * G(xx, yy): B(xx, yy) = 2 * B(xx, yy)
Else
R(xx, yy) = R(xx, yy) / 2: G(xx, yy) = G(xx, yy) / 2: B(xx, yy) = B(xx, yy) / 2
End If
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub Blinds(Ob As Object, Dist%, Reverse As Boolean) 'hor blinds
Dim rt%
On Error Resume Next
Imagewait True
If Reverse = False Then
rt = 0
Else
rt = Dist
End If
    For yy = 0 To Ob.Height - 1
    For xx = 0 To Ob.Width - 1
R(xx, yy) = R(xx, yy) - (rt * R(xx, yy) / Dist)
G(xx, yy) = G(xx, yy) - (rt * G(xx, yy) / Dist)
B(xx, yy) = B(xx, yy) - (rt * B(xx, yy) / Dist)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next xx
If Reverse = False Then
    rt = rt + 1
    If rt = Dist Then rt = 0
Else
    rt = rt - 1
    If rt = 0 Then rt = Dist
End If
Next yy
Imagewait False
Ob.Refresh
End Sub

Public Sub Blinds2(Ob As Object, Dist%, Reverse As Boolean) 'vert blinds
Dim rt%
On Error Resume Next
Imagewait True
If Reverse = False Then
rt = 0
Else
rt = Dist
End If
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
R(xx, yy) = R(xx, yy) - (rt * R(xx, yy) / Dist)
G(xx, yy) = G(xx, yy) - (rt * G(xx, yy) / Dist)
B(xx, yy) = B(xx, yy) - (rt * B(xx, yy) / Dist)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
If Reverse = False Then
    rt = rt + 1
    If rt = Dist Then rt = 0
Else
    rt = rt - 1
    If rt = 0 Then rt = Dist
End If
Next xx
Imagewait False
Ob.Refresh
End Sub

Public Sub HLines(Ob As Object, Dist%, AB As Single, R2%, G2%, B2%)
Dim R1%, G1%, B1%
On Error Resume Next
Imagewait True
AB = AB / 10
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1 Step Dist
    Color = GetPixel(Ob.hdc, xx, yy)
    R1 = Color Mod 256&
    G1 = ((Color And &HFF00) / 256&) Mod 256&
    B1 = (Color And &HFF0000) / 65536
        R1 = (R1 * (1 - AB)) + (R2 * AB)
        G1 = (G1 * (1 - AB)) + (G2 * AB)
        B1 = (B1 * (1 - AB)) + (B2 * AB)
'put the new colors back
SetPixel Ob.hdc, xx, yy, RGB(R1, G1, B1)
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub VLines(Ob As Object, Dist%, AB As Single, R2%, G2%, B2%)
Dim R1%, G1%, B1%
On Error Resume Next
Imagewait True
AB = AB / 10
    For yy = 0 To Ob.Height - 1
    For xx = 0 To Ob.Width - 1 Step Dist
    Color = GetPixel(Ob.hdc, xx, yy)
    R1 = Color Mod 256&
    G1 = ((Color And &HFF00) / 256&) Mod 256&
    B1 = (Color And &HFF0000) / 65536
        R1 = (R1 * (1 - AB)) + (R2 * AB)
        G1 = (G1 * (1 - AB)) + (G2 * AB)
        B1 = (B1 * (1 - AB)) + (B2 * AB)
'put the new colors back
SetPixel Ob.hdc, xx, yy, RGB(R1, G1, B1)
Next xx, yy
Imagewait False
Ob.Refresh
End Sub

Public Sub AddSquares(Ob As Object, Dist%, AB As Single, R1%, G1%, B1%) 'add squares
On Error Resume Next
AB = AB / 10
Imagewait True
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1 Step Dist
Color = GetPixel(Ob.hdc, xx, yy)
R(xx, yy) = Color Mod 256&
G(xx, yy) = ((Color And &HFF00) / 256&) Mod 256&
B(xx, yy) = (Color And &HFF0000) / 65536
    R(xx, yy) = (R(xx, yy) * (1 - AB)) + (R1 * AB)
    G(xx, yy) = (G(xx, yy) * (1 - AB)) + (G1 * AB)
    B(xx, yy) = (B(xx, yy) * (1 - AB)) + (B1 * AB)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
    For yy = 0 To Ob.Height - 1
    For xx = 0 To Ob.Width - 1 Step Dist
Color = GetPixel(Ob.hdc, xx, yy)
R(xx, yy) = Color Mod 256&
G(xx, yy) = ((Color And &HFF00) / 256&) Mod 256&
B(xx, yy) = (Color And &HFF0000) / 65536
    R(xx, yy) = (R(xx, yy) * (1 - AB)) + (R1 * AB)
    G(xx, yy) = (G(xx, yy) * (1 - AB)) + (G1 * AB)
    B(xx, yy) = (B(xx, yy) * (1 - AB)) + (B1 * AB)
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next xx, yy
Imagewait False
Ob.Refresh
End Sub

Public Sub AddBoxes(Ob As Object, Dist%, AB As Single, R1%, G1%, B1%) 'add boxes
On Error Resume Next
Dim ttt%
Dim Rm, Gm, Bm
ttt = Dist
AB = AB / 10
Imagewait True
For xx = 0 To (Ob.Width / 2) - 1 Step Dist
If xx > (Ob.Width / 2) - Dist Or xx > (Ob.Height / 2) - Dist Then GoTo addboxes2
Ob.Line (ttt, ttt)-(Ob.Width - ttt, Ob.Height - ttt), RGB(R1, G1, B1), B
ttt = ttt + Dist
Next xx
addboxes2:
    For xx = 0 To Ob.Width - 1
    For yy = 0 To Ob.Height - 1
Color = GetPixel(Ob.hdc, xx, yy)
If Color <> 0 Then
    R(xx, yy) = (R(xx, yy) * (1 - AB)) + (R1 * AB)
    G(xx, yy) = (G(xx, yy) * (1 - AB)) + (G1 * AB)
    B(xx, yy) = (B(xx, yy) * (1 - AB)) + (B1 * AB)
End If
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub

Public Sub AddCircles(R1%, G1%, B1%, Dist%) 'add circles
On Error Resume Next
Dim ttt%
Dim Rm, Gm, Bm
CDMain.Mem2.Move 0, 0, PicDimX, PicDimY
ttt = Dist
Imagewait CDMain, True
CDMain.Mem2.Cls
CDMain.Mem2.BackColor = 0
For xx = 0 To PicDimX / Dist
CDMain.Mem2.Circle (PicDimX / 2, PicDimY / 2), ttt, RGB(R1, G1, B1)
If ttt > Int(Sqr(2 * ((PicDimX / 2) ^ 2))) Then Exit For
ttt = ttt + Dist
Next xx
End Sub

Public Sub AddDiaLines1(R1%, G1%, B1%, Dist%)  'add dia R lines
On Error Resume Next
Dim ttt%
Dim Rm, Gm, Bm
CDMain.Mem2.Move 0, 0, PicDimX, PicDimY
ttt = Dist
Imagewait CDMain, True
CDMain.Mem2.Cls
CDMain.Mem2.BackColor = 0
For xx = 0 To (PicDimX / Dist) * 2
CDMain.Mem2.Line (0, ttt)-(ttt, 0), RGB(R1, G1, B1)
ttt = ttt + Dist
Next xx
End Sub

Public Sub AddDiaLines2(R1%, G1%, B1%, Dist%) 'add dia L lines
On Error Resume Next
Dim ttt%
Dim Rm, Gm, Bm
CDMain.Mem2.Move 0, 0, PicDimX, PicDimY
ttt = Dist
Imagewait CDMain, True
CDMain.Mem2.Cls
CDMain.Mem2.BackColor = 0
For xx = 0 To (PicDimX / Dist) * 2
CDMain.Mem2.Line (0, PicDimY - ttt)-(PicDimX, (2 * PicDimY) - ttt), RGB(R1, G1, B1)
ttt = ttt + Dist
Next xx
End Sub

Public Sub AddCrossedLines(R1%, G1%, B1%, Dist%)  'add dia crossed lines
On Error Resume Next
Dim ttt%
ttt = Dist
CDMain.Mem2.Move 0, 0, PicDimX, PicDimY
CDMain.Mem2.Cls
CDMain.Mem2.BackColor = 0
For xx = 0 To (PicDimX / Dist) * 2
CDMain.Mem2.Line (0, ttt)-(ttt, 0), RGB(R1, G1, B1)
CDMain.Mem2.Line (0, PicDimY - ttt)-(PicDimX, (2 * PicDimY) - ttt), RGB(R1, G1, B1)
ttt = ttt + Dist
Next xx
End Sub


