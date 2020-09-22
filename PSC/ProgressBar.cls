VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**********************************
'* CODE BY: PATRICK MOORE (ZELDA) *
'* Feel free to re-distribute or  *
'* Use in your own projects.      *
'* Giving credit to me would be   *
'* nice :)                        *
'*                                *
'* Please vote for me if you find *
'* this code useful :]   -Patrick *
'**********************************
'
'PS: Please look for more submissions to PSC by me
'    shortly.  I've recently been working on a lot
'    :))  All my submissions are under author name
'    "Patrick Moore (Zelda)"
Public pic1 As PictureBox
Private lPercent As Integer
Private ProgNumColors As Integer
Private color(20) As Long
Private rRed As Long, rBlue As Long, rGreen As Long

'Public Percent As Integer
Public Property Get Percent() As Integer
    Percent = lPercent
End Property
Public Property Let Percent(aPercent As Integer)
    lPercent = aPercent
    Refresh
End Property

Public Sub ResetColors()
ProgNumColors = 0
End Sub
Public Sub AddColor(colorLong As Long)
'Add a color to the amount
ProgNumColors = ProgNumColors + 1
On Error Resume Next
color(ProgNumColors) = colorLong
End Sub
Public Sub Refresh()
Dim X As Integer, CurColor As Integer, i As Integer
'If the first color is different from the startcolor
'(meaning the user changed it) re-gradient the picturebox
'Resize the label to display the proper percent
If pic1.Point(0, 0) <> color(1) Or pic1.Point(pic1.Width, pic1.Height) <> color(ProgNumColors) Then
    If ProgNumColors = 1 Then
        'If there's only one color, set the backcolor to that
        pic1.BackColor = color(1)
    Else
        'Define a section
        i = (pic1.Width / (ProgNumColors - 1))
        'Fade from one color to another
        For CurColor = 1 To ProgNumColors - 1
            Gradient pic1, color(CurColor), color(CurColor + 1) _
            , i * (CurColor - 1), i
        Next CurColor
    End If
End If

'Draw gray lines to accurately show the percent
For X = (Percent * (pic1.Width / 100)) + 1 To pic1.Width
    pic1.Line (X, 0)-(X, pic1.Height), "&H8000000F", BF
Next X
End Sub

Private Sub pic1_Click()

End Sub



Function RGBfromLONG(LongCol As Long)
' Get The Red, Blue And Green Values Of A Colour From The Long Value
Dim Blue As Double, Green As Double, Red As Double, GreenS As Double, BlueS As Double
Blue = Fix((LongCol / 256) / 256)
Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
rRed = Red: rBlue = Blue: rGreen = Green
End Function

Public Sub Gradient(picBox, StartColor As Long, EndColor As Long, StartingXPos As Integer, Width As Integer)
Dim W As Integer, color As Long
Dim GradR As Integer, GradB As Integer, GradG As Integer
Dim B1 As Integer, B2 As Integer
Dim G1 As Integer, G2 As Integer
Dim R1 As Integer, R2 As Integer

'Determine Red, Green, and Blue values
'for the first color
RGBfromLONG StartColor
B1 = rBlue
G1 = rGreen
R1 = rRed

'Determine Red, Green, and Blue values
'for the last color
RGBfromLONG EndColor
B2 = rBlue
G2 = rGreen
R2 = rRed

'picBox.ScaleMode = 3
For W = StartingXPos To StartingXPos + Width
    GradR = ((R2 - R1) / Width * (W - StartingXPos)) + R1
    GradG = ((G2 - G1) / Width * (W - StartingXPos)) + G1
    GradB = ((B2 - B1) / Width * (W - StartingXPos)) + B1
    If GradG < 0 Then GradG = 0
    If GradR < 0 Then GradR = 0
    If GradB < 0 Then GradB = 0
    color = RGB(GradR, GradG, GradB)
    picBox.Line (W, 0)-(W, picBox.ScaleHeight), color, B
Next W
End Sub

Public Sub Clear()
    pic1.Cls
End Sub
