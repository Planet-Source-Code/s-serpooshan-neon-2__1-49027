VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormNeon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Neon Ver 2.0"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5595
   Icon            =   "FormNeon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2640
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   94
      TabIndex        =   27
      Top             =   3750
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox PicView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   180
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   0
      Top             =   3150
      Width           =   1905
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options: "
      Height          =   1275
      Left            =   210
      TabIndex        =   16
      Top             =   1560
      Width           =   2505
      Begin VB.TextBox TextTimer 
         Height          =   285
         Left            =   1860
         TabIndex        =   29
         Text            =   "100"
         ToolTipText     =   "Timer Interval (msec)"
         Top             =   300
         Width           =   450
      End
      Begin VB.TextBox TextWidthObj 
         Height          =   285
         Left            =   660
         TabIndex        =   23
         Text            =   "50"
         Top             =   300
         Width           =   450
      End
      Begin VB.TextBox TextStep 
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Text            =   "1"
         Top             =   780
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "Timer:"
         Height          =   255
         Left            =   1380
         TabIndex        =   28
         ToolTipText     =   "Timer Interval (msec)"
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Width:"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Moving Step:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   840
         Width           =   945
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5160
      Top             =   2850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.BMP ; *.JPG; *.GIF|*.BMP;*.JPG;*.GIF|All Files |*.*"
      Flags           =   1
      FontName        =   "Tahoma"
   End
   Begin VB.Frame Frame3 
      Caption         =   "Main Object: "
      Height          =   1275
      Left            =   2850
      TabIndex        =   9
      Top             =   150
      Width           =   2505
      Begin VB.CommandButton CommandMask 
         Caption         =   "Mask"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         TabIndex        =   15
         Top             =   705
         Width           =   615
      End
      Begin VB.CommandButton CommandText 
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   990
         TabIndex        =   14
         Top             =   270
         Width           =   645
      End
      Begin VB.OptionButton OptionTextImage 
         Caption         =   "Image"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   750
         Width           =   765
      End
      Begin VB.OptionButton OptionTextImage 
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton CommandImage 
         Caption         =   "Image"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   990
         TabIndex        =   11
         Top             =   705
         Width           =   645
      End
      Begin VB.CommandButton CommandFont 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         TabIndex        =   10
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pixel picture: "
      Height          =   1275
      Left            =   2850
      TabIndex        =   6
      Top             =   1560
      Width           =   2505
      Begin VB.PictureBox Picture3 
         Height          =   345
         Left            =   1320
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   21
         Top             =   720
         Width           =   975
         Begin VB.PictureBox PicOFFp 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            Picture         =   "FormNeon.frx":0E42
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   58
            TabIndex        =   22
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   345
         Left            =   180
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   19
         Top             =   720
         Width           =   975
         Begin VB.PictureBox PicONp 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            Picture         =   "FormNeon.frx":0EE8
            ScaleHeight     =   255
            ScaleWidth      =   870
            TabIndex        =   20
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.CommandButton CommandPixelPic 
         Caption         =   "OFF pixels"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton CommandPixelPic 
         Caption         =   "ON pixels"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Move: "
      Height          =   1275
      Left            =   210
      TabIndex        =   3
      Top             =   150
      Width           =   2505
      Begin VB.CommandButton CommandPlay 
         Caption         =   "Play"
         Height          =   345
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton CommandStop 
         Caption         =   "Stop"
         Height          =   345
         Left            =   1440
         TabIndex        =   25
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton OptionDir 
         Caption         =   "Left"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   870
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptionDir 
         Caption         =   "Right"
         Height          =   240
         Index           =   1
         Left            =   1590
         TabIndex        =   4
         Top             =   870
         Width           =   765
      End
   End
   Begin VB.PictureBox PicLabel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4500
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   2
      Top             =   3300
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox TextNeon 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4290
      TabIndex        =   1
      Text            =   "Visual Basic Programming     * * *    ' N e o n '     Simulator ,   By: S.Serpooshan    (2003)"
      Top             =   3840
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4230
      Top             =   2880
   End
End
Attribute VB_Name = "FormNeon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------------------------------
'      «In The Name Of The Most High »
'
' Neon Ver 1.0
' This program draw your text in a neon board
'
' by: Saeed Serpooshan - Iran - 2001 (1380)
' EMail: SSerpooshan@Yahoo.com , Admin@JamAcademic.Com
' WebPage: http://www.JamAcademic.com/vb
'------------------------------------------------------------------------------------------------------------------------------------------

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)

Dim xBegin As Integer, ShouldRedraw As Integer, dxText As Long, dyText As Long, dxOfBorder As Long
Dim ViewW As Long, ViewH As Long, dxPerViewWindow As Long, dyPerViewWindow As Long
Dim pixW As Integer, pixH As Integer
Dim xbStep As Integer
Dim RightDirection As Integer


Private Sub CommandFont_Click()
On Error Resume Next
CD1.FontName = PicLabel.Font.Name
CD1.FontBold = PicLabel.Font.Bold
CD1.FontSize = PicLabel.Font.Size
CD1.FontItalic = PicLabel.Font.Italic
CD1.FontUnderline = PicLabel.Font.Underline
CD1.FontStrikethru = PicLabel.Font.Strikethrough
Err.Clear

CD1.ShowFont
If Err Then Exit Sub

PicLabel.Font.Name = CD1.FontName
PicLabel.Font.Bold = CD1.FontBold
PicLabel.Font.Size = CD1.FontSize
PicLabel.Font.Italic = CD1.FontItalic
PicLabel.Font.Underline = CD1.FontUnderline
PicLabel.Font.Strikethrough = CD1.FontStrikethru
Set TextNeon.Font = PicLabel.Font
ResetParams

End Sub

Private Sub CommandImage_Click()
MsgBox "this will add in future versions..."
End Sub

Private Sub CommandPixelPic_Click(Index As Integer)
On Error Resume Next
CD1.ShowOpen
If Err Then Exit Sub
F = CD1.FileName
If Index = 0 Then Set PicONp.Picture = LoadPicture(F) Else Set PicOFFp.Picture = LoadPicture(F)
Call ResetParams
End Sub

Private Sub CommandStop_Click()
Timer1.Enabled = False
End Sub

Private Sub CommandPlay_Click()
Timer1.Enabled = True
End Sub

Private Sub CommandText_Click()
On Error Resume Next
a = InputBox("Enter your text:", "Set Text", TextNeon.Text)
If a <> "" Then TextNeon.Text = a: Call ResetParams

End Sub

Private Sub Form_Load()
Show
DoEvents
PicONp.AutoSize = True
PicOFFp.AutoSize = True
Call ResetParams
Call Form_Resize
xb1 = xBegin: Call DrawNeon: xBegin = xb1
Call CommandPlay_Click 'Auto Play at start up
End Sub

Private Sub Form_Resize()
If PicView.Width < FormNeon.ScaleWidth Then
    PicView.Left = (FormNeon.ScaleWidth - PicView.Width) \ 2
Else
    PicView.Left = 0
End If
End Sub

Private Sub OptionDir_Click(Index As Integer)
RightDirection = OptionDir(1).Value = True
End Sub

Private Sub OptionTextImage_Click(Index As Integer)
 Dim a As Boolean
 a = Index = 0
 CommandText.Enabled = a: CommandFont.Enabled = a
 CommandImage.Enabled = Not a: CommandMask.Enabled = Not a
End Sub

Private Sub TextHeightObj_Change()
a = Val(TextHeightObj.Text)
If a > 800 Or a < 0 Then Exit Sub
PicView.Height = a + dxOfBorder
PicTemp.Height = PicView.Height
ShouldRedraw = True
xb1 = xBegin: DrawNeon: xBegin = xb1
End Sub

Private Sub TextNeon_Change()
ResetParams
End Sub

Private Sub TextStep_Change()
xbStep = Val(TextStep.Text)
End Sub

Private Sub TextTimer_Change()
Dim a As Long
a = Val(TextTimer.Text)
If a > 0 Then Timer1.Interval = a
End Sub

Private Sub TextWidthObj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0: Call TextWidthObj_LostFocus
End Sub

Private Sub TextWidthObj_LostFocus()
a = Val(TextWidthObj.Text) * pixW
If a > 1100 Or a < 0 Then Exit Sub
If PicView.Width <> a + dxOfBorder Then
    Call ResetParams
    Call Form_Resize
End If
End Sub

Private Sub Timer1_Timer()
Dim s As Long
If RightDirection Then 'reverse motion
    xBegin = xBegin - xbStep
    If xBegin < 0 Then xBegin = dxText + dxPerViewWindow
Else 'default motion (left)
    xBegin = xBegin + xbStep
    If xBegin > dxText + dxPerViewWindow Then xBegin = 0
End If
Call DrawNeon
End Sub

Sub DrawNeon()
Dim xx As Long
xx = xBegin * pixW
'PicView.PaintPicture PicTemp.Image, 0, 0, ViewW, ViewH, xx, 0, ViewW, ViewH
'the BitBlt function is faster than PaintPicture so use it:
BitBlt PicView.hdc, 0, 0, ViewW, ViewH, PicTemp.hdc, xx, 0, SRCCOPY
PicView.PSet (0, 0), GetPixel(PicView.hdc, 0, 0) 'cause to refresh PicView picture!
End Sub

Sub ResetParams()

dxOfBorder = PicView.Width - PicView.ScaleWidth 'usually is 4

RightDirection = OptionDir(1).Value = True

dxText = PicLabel.TextWidth(TextNeon.Text)
If PicLabel.FontItalic Then dxText = dxText + PicLabel.TextWidth("X")
dyText = PicLabel.TextHeight(TextNeon.Text)

PicLabel.Width = dxText + dxOfBorder
PicLabel.Height = dyText + dxOfBorder
TextNeon.Height = PicLabel.Height

PicLabel.Cls
PicLabel.Print TextNeon.Text

xbStep = Val(TextStep.Text)

pixW = PicOFFp.Width: pixH = PicOFFp.Height

ViewW = Val(TextWidthObj.Text) * pixW
ViewH = dyText * pixH
PicView.Width = ViewW + dxOfBorder '+4 is because of borders
PicView.Height = ViewH + dxOfBorder

dxPerViewWindow = PicView.ScaleWidth / pixW
dyPerViewWindow = PicView.ScaleHeight / pixH

' dxPerViewWindow*2 is because of left & right free space required during motion!
PicTemp.Width = (PicLabel.ScaleWidth + dxPerViewWindow * 2) * pixW
PicTemp.Height = PicLabel.ScaleHeight * pixH

Call BuildPicTempImage

xBegin = IIf(RightDirection, dxText + dxPerViewWindow, 0)
Call DrawNeon

End Sub

Sub BuildPicTempImage()
Dim hdcPicTemp As Long, hdcPicLabel As Long, hdcON As Long, hdcOFF As Long
Dim X As Long, Y As Long, xx As Long, yy As Long
Dim c As Long, ColON As Long

ColON = PicLabel.ForeColor
hdcPicTemp = PicTemp.hdc: hdcPicLabel = PicLabel.hdc
hdcON = PicONp.hdc: hdcOFF = PicOFFp.hdc

'fill first column by OFF bitmap:
xx = 0: yy = 0
For Y = 0 To dyText - 1
    BitBlt hdcPicTemp, xx, yy, pixW, pixH, hdcOFF, 0, 0, SRCCOPY
    yy = yy + pixH
Next

'fill left margin by OFF pixels:
For X = 0 To dxPerViewWindow - 1
    BitBlt hdcPicTemp, xx, 0, pixW, ViewH, hdcPicTemp, 0, 0, SRCCOPY
    xx = xx + pixW
Next

'copy left margin to right margin:
xx = (dxPerViewWindow + dxText) * pixW
BitBlt hdcPicTemp, xx, 0, ViewW, ViewH, hdcPicTemp, 0, 0, SRCCOPY

'fill main mid part:
xx = ViewW 'Begin After Left Margin

For X = 0 To dxText - 1
    yy = 0
    For Y = 0 To dyText - 1
        'c = PicLabel.Point(x, y)
        c = GetPixel(hdcPicLabel, X, Y)
        If c = ColON Then
            'PicTemp.PaintPicture imgON, xx, yy
            BitBlt hdcPicTemp, xx, yy, pixW, pixH, hdcON, 0, 0, SRCCOPY
        Else
            'PicTemp.PaintPicture imgOFF, xx, yy
            BitBlt hdcPicTemp, xx, yy, pixW, pixH, hdcOFF, 0, 0, SRCCOPY
        End If
        yy = yy + pixH
        Next
        'DoEvents
    xx = xx + pixW
Next

End Sub
