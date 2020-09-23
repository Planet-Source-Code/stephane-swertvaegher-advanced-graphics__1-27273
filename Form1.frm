VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Graphical Effects"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PatPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   5535
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   31
      Top             =   4005
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   90
      TabIndex        =   5
      Top             =   4230
      Width           =   4920
      Begin VB.HScrollBar HScroll2 
         Height          =   285
         Left            =   1485
         Max             =   10
         Min             =   1
         TabIndex        =   32
         Top             =   675
         Value           =   5
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   285
         Left            =   1485
         Max             =   10
         Min             =   1
         TabIndex        =   12
         Top             =   315
         Value           =   5
         Width           =   1950
      End
      Begin VB.HScrollBar HS1 
         Height          =   240
         Index           =   0
         LargeChange     =   10
         Left            =   1485
         Max             =   255
         TabIndex        =   11
         Top             =   1125
         Width           =   1950
      End
      Begin VB.HScrollBar HS1 
         Height          =   240
         Index           =   1
         LargeChange     =   10
         Left            =   1485
         Max             =   255
         TabIndex        =   10
         Top             =   1440
         Width           =   1950
      End
      Begin VB.HScrollBar HS1 
         Height          =   240
         Index           =   2
         LargeChange     =   10
         Left            =   1485
         Max             =   255
         TabIndex        =   9
         Top             =   1755
         Width           =   1950
      End
      Begin VB.HScrollBar HS1 
         Height          =   240
         Index           =   3
         LargeChange     =   10
         Left            =   1485
         Max             =   255
         TabIndex        =   8
         Top             =   2295
         Width           =   1950
      End
      Begin VB.HScrollBar HS1 
         Height          =   240
         Index           =   4
         LargeChange     =   10
         Left            =   1485
         Max             =   255
         TabIndex        =   7
         Top             =   2610
         Width           =   1950
      End
      Begin VB.HScrollBar HS1 
         Height          =   240
         Index           =   5
         LargeChange     =   10
         Left            =   1485
         Max             =   255
         TabIndex        =   6
         Top             =   2925
         Width           =   1950
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3510
         TabIndex        =   34
         Top             =   675
         Width           =   510
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Scale:"
         Height          =   285
         Left            =   135
         TabIndex        =   33
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alpha Blending:"
         Height          =   285
         Left            =   135
         TabIndex        =   30
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3510
         TabIndex        =   29
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   285
         Index           =   0
         Left            =   3510
         TabIndex        =   28
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   285
         Index           =   1
         Left            =   3510
         TabIndex        =   27
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   285
         Index           =   2
         Left            =   3510
         TabIndex        =   26
         Top             =   1710
         Width           =   465
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Red:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   855
         TabIndex        =   25
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Green:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   1
         Left            =   855
         TabIndex        =   24
         Top             =   1395
         Width           =   600
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Blue:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   2
         Left            =   855
         TabIndex        =   23
         Top             =   1710
         Width           =   600
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Index           =   0
         Left            =   4185
         TabIndex        =   22
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   285
         Index           =   3
         Left            =   3510
         TabIndex        =   21
         Top             =   2250
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   285
         Index           =   4
         Left            =   3510
         TabIndex        =   20
         Top             =   2565
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   285
         Index           =   5
         Left            =   3510
         TabIndex        =   19
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Red:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   3
         Left            =   855
         TabIndex        =   18
         Top             =   2250
         Width           =   600
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Green:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   4
         Left            =   855
         TabIndex        =   17
         Top             =   2565
         Width           =   600
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Blue:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   5
         Left            =   855
         TabIndex        =   16
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Index           =   1
         Left            =   4185
         TabIndex        =   15
         Top             =   2250
         Width           =   420
      End
      Begin VB.Label Label7 
         Caption         =   "Color 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   1125
         Width           =   690
      End
      Begin VB.Label Label7 
         Caption         =   "Color 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   2295
         Width           =   690
      End
   End
   Begin VB.PictureBox Pic3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   5445
      Picture         =   "Form1.frx":51E9
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox Pic4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   5445
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   4
      Top             =   90
      Width           =   4800
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   5445
      Picture         =   "Form1.frx":E4A9
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   90
      Picture         =   "Form1.frx":11984
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   90
      Width           =   4800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   3825
      Width           =   4785
   End
   Begin VB.Menu mnuColor 
      Caption         =   "Color"
      Begin VB.Menu mnuCol 
         Caption         =   "Set gradient"
         Index           =   0
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Add color"
         Index           =   2
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Substract color"
         Index           =   3
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Brighten picture"
         Index           =   5
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Darken picture"
         Index           =   6
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Kill Red component"
         Index           =   8
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Kill Green component"
         Index           =   9
      End
      Begin VB.Menu mnuCol 
         Caption         =   "KIll Blue component"
         Index           =   10
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Swap Red and Green"
         Index           =   12
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Swap Red and Blue"
         Index           =   13
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Swap Green and Blue"
         Index           =   14
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Negative Red component"
         Index           =   16
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Negative Green component"
         Index           =   17
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Negative Blue component"
         Index           =   18
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Photo-negative"
         Index           =   19
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Greyscale"
         Index           =   21
      End
   End
   Begin VB.Menu mnuMixing 
      Caption         =   "Mixing"
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with picture"
         Index           =   0
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with pattern"
         Index           =   1
      End
      Begin VB.Menu mnuMix 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Add picture"
         Index           =   3
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Add silhouette"
         Index           =   4
      End
      Begin VB.Menu mnuMix 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Slide picture"
         Index           =   6
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "Filters"
      Begin VB.Menu mnuFil 
         Caption         =   "Emboss"
         Index           =   0
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Emboss - hold red"
         Index           =   1
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Emboss - hold green"
         Index           =   2
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Emboss - hold blue"
         Index           =   3
      End
      Begin VB.Menu mnuFil 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Blur"
         Index           =   5
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Blur more"
         Index           =   6
      End
      Begin VB.Menu mnuFil 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Diffuse picture"
         Index           =   8
      End
      Begin VB.Menu mnuFil 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Sharpen"
         Index           =   10
      End
      Begin VB.Menu mnuFil 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Erode"
         Index           =   12
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Blow"
         Index           =   13
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Contrast"
         Index           =   14
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Fog"
         Index           =   15
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Noise"
         Index           =   16
      End
      Begin VB.Menu mnuFil 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Freeze"
         Index           =   18
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Freeze more"
         Index           =   19
      End
      Begin VB.Menu mnuFil 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Black and white 1"
         Index           =   21
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Black and white 2"
         Index           =   22
      End
      Begin VB.Menu mnuFil 
         Caption         =   "Black and white 3"
         Index           =   23
      End
   End
   Begin VB.Menu mnuSpecFilters 
      Caption         =   "Special Filters"
      Begin VB.Menu mnuSFil 
         Caption         =   "Brown"
         Index           =   0
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Liquid"
         Index           =   1
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Yellow"
         Index           =   2
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Charcoal"
         Index           =   3
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Dark moon"
         Index           =   4
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Total eclipse"
         Index           =   5
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Purple rain"
         Index           =   6
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Spooky"
         Index           =   7
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Unreal"
         Index           =   8
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Flame"
         Index           =   9
      End
      Begin VB.Menu mnuSFil 
         Caption         =   "Aquarel"
         Index           =   10
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "Effects"
      Begin VB.Menu mnuEff 
         Caption         =   "Add hor. blinds"
         Index           =   0
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add reversed hor. blinds"
         Index           =   1
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add vert. blinds"
         Index           =   2
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add reversed vert. blinds"
         Index           =   3
      End
      Begin VB.Menu mnuEff 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add horiz. lines"
         Index           =   5
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add vert. lines"
         Index           =   6
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add squares"
         Index           =   7
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add boxes"
         Index           =   8
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add circles"
         Index           =   9
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ReDim R(Pic1.Width - 1, Pic1.Height - 1)
ReDim G(Pic1.Width - 1, Pic1.Height - 1)
ReDim B(Pic1.Width - 1, Pic1.Height - 1)
Set Mem = Pic1.Image
Label1.Visible = False
Label3.Caption = Format(HScroll1.Value / 10, "0.0")
Label9.Caption = Format(HScroll2.Value / 10, "00%")
HS1(0).Value = 255
HS1(1).Value = 128
HS1(2).Value = 64
HS1(3).Value = 64
HS1(4).Value = 196
HS1(5).Value = 240
End Sub

Private Sub HS1_Change(Index As Integer)
Label4(Index).Caption = Format(HS1(Index).Value, "000")
Label6(0).BackColor = RGB(HS1(0).Value, HS1(1).Value, HS1(2).Value)
Label6(1).BackColor = RGB(HS1(3).Value, HS1(4).Value, HS1(5).Value)
End Sub

Private Sub HScroll1_Change()
Label3.Caption = Format(HScroll1.Value / 10, "0.0")
End Sub

Private Sub HScroll2_Change()
Label9.Caption = Format(HScroll2.Value / 10, "00%")
End Sub

Private Sub mnuCol_Click(Index As Integer)
Pic4.Picture = Nothing
ReadColor 0, 0, Pic1.Width, Pic1.Height
Select Case Index
Case 0 'set gradient
GradientCol Pic4, HScroll1.Value, HS1(0).Value, HS1(1).Value, HS1(2).Value, HS1(3).Value, HS1(4).Value, HS1(5).Value
Case 2 'add color
AddColor Pic4, HS1(0).Value, HS1(1).Value, HS1(2).Value, True
Case 3 'substract color
AddColor Pic4, HS1(0).Value, HS1(1).Value, HS1(2).Value, False
Case 5 'Brighten picture
BrightenPicture Pic4, 5 'strength can be from 1 to 10
Case 6 'Darken picture
BrightenPicture Pic4, -5 'strength can be from -1 to -10
Case 8 'Kill red component
KillColor Pic4, 0
Case 9 'Kill green component
KillColor Pic4, 1
Case 10 'Kill blue component
KillColor Pic4, 2
Case 12 'Swap R & G
SwapColor Pic4, 0
Case 13 'Swap R & B
SwapColor Pic4, 1
Case 14 'Swap G & B
SwapColor Pic4, 2
Case 16 'negative red
NegativeColor Pic4, 0
Case 17 'negative green
NegativeColor Pic4, 1
Case 18 'negative blue
NegativeColor Pic4, 2
Case 19 'photo-negative
NegativeColor Pic4, 3
Case 21 'greyscale
GreyColor Pic4
End Select
End Sub

Private Sub mnuEff_Click(Index As Integer)
ReadColor 0, 0, Pic1.Width, Pic1.Height
Pic4.Picture = Nothing
Select Case Index
Case 0 'hor blinds
Blinds Pic4, 20, False
Case 1 'hor blinds reversed
Blinds Pic4, 20, True
Case 2 'vert blinds
Blinds2 Pic4, 20, False
Case 3 'vert blinds reversed
Blinds2 Pic4, 20, True
Case 5 'hor. lines
HLines Pic4, 20, HScroll1.Value, HS1(0).Value, HS1(1).Value, HS1(2).Value
Case 6 'vert. lines
VLines Pic4, 20, HScroll1.Value, HS1(0).Value, HS1(1).Value, HS1(2).Value
Case 7 'add squares
AddSquares Pic4, 20, HScroll1.Value, HS1(0).Value, HS1(1).Value, HS1(2).Value
Case 8 'add boxes
Pic4.Picture = Nothing
Pic4.BackColor = 0
AddBoxes Pic4, 20, HScroll1.Value, HS1(0).Value, HS1(1).Value, HS1(2).Value
Case 9 'add circles
End Select
End Sub

Private Sub mnuFil_Click(Index As Integer)
Pic4.Picture = Nothing
ReadColor 0, 0, Pic1.Width, Pic1.Height
Select Case Index
Case 0 'emboss
EmbossPicture Pic4
Case 1 'hold red and emboss
HoldRed Pic4, 64
Case 2 'hold green and emboss
HoldGreen Pic4, 64
Case 3 'hold blue and emboss
HoldBlue Pic4, 64
Case 5 'blur
BlurPicture Pic4
Case 6 'blur more
BlurPictureMore Pic4
Case 8 'diffuse
DiffusePicture Pic4, 3
Case 10 'sharpen
SharpenPicture Pic4
Case 12 'erode
ErodePicture Pic4, 5 'eroding between 2 and 10
Case 13 'blow
BlowPicture Pic4, 1.5 'blowing between 1 and 2
Case 14 'contrast
ContrastPicture Pic4, 10 'contrast between 1 and 10
Case 15 'fog
FogPicture Pic4, 20 'fog between 1 and 40
Case 16 'noise
AddNoise Pic4
Case 18 'freeze
Freeze Pic4, 1.5
Case 19 'freeze more
Freeze Pic4, 2
Case 21 ' black and white 1
BnW Pic4, 200
Case 22 ' black and white 2
BnW Pic4, 150
Case 23 ' black and white 3
BnW Pic4, 100
End Select
End Sub

Private Sub mnuMix_Click(Index As Integer)
Pic4.Picture = Nothing
ReadColor 0, 0, Pic1.Width, Pic1.Height
Select Case Index
Case 0 'mix with picture
MixPic Pic4, Pic2, HScroll1.Value
Case 1 'mix with pattern
MixPat Pic4, PatPic, HScroll1.Value
Case 3 'add picture
AddPic Pic4, Pic3, HScroll1.Value, HScroll2.Value
Case 4 'add silhouette
AddSil Pic4, Pic3, HScroll1.Value, HScroll2.Value
Case 6 'slide picture
SlidePic Pic4, Pic2
End Select
End Sub

Private Sub mnuSFil_Click(Index As Integer)
Pic4.Picture = Nothing
ReadColor 0, 0, Pic1.Width, Pic1.Height
Select Case Index
Case 0 'brown
Brown Pic4
Case 1 'liquid
Liquid Pic4
Case 2 'yellow
Yellow Pic4
Case 3 'charcoal
Charcoal Pic4
Case 4 'dark moon
DarkMoon Pic4
Case 5 'total eclipse
TotalEclipse Pic4
Case 6 'purple rain
PurpleRain Pic4
Case 7 'spooky
Spooky Pic4
Case 8 'unreal
UnReal Pic4
Case 9 'flame
Flame Pic4
Case 10 'aquarel
Aquarel Pic4
End Select
End Sub
