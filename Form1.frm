VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPicRepair 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Repair Pictures"
   ClientHeight    =   11340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   756
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   596
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CboColor 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2040
      List            =   "Form1.frx":0013
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   120
      Width           =   3255
   End
   Begin VB.HScrollBar ScrSim 
      Height          =   255
      Left            =   2640
      Max             =   255
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton CmdRepCirc 
      Caption         =   "Circular Repair"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   6840
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load Picture"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CmdRepair 
      Caption         =   "Linear Repair"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox PicOrg 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   10050
      Left            =   120
      Picture         =   "Form1.frx":006F
      ScaleHeight     =   670
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   580
      TabIndex        =   0
      Top             =   1200
      Width           =   8700
   End
   Begin VB.Label LblDamaged 
      Alignment       =   2  'Zentriert
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Difference 0"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Ausgefüllt
      Height          =   495
      Left            =   2040
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Select Transparent color by Clicking on the Picture"
      Height          =   435
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmPicRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple Image Repair
' by Scythe
Option Explicit

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Type RGBQUAD
    rgbBlue As Byte
    rgbgreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private TR As Byte
Private TB As Byte
Private TG As Byte

Private PicInfo As BITMAP
Private PicAr1() As RGBQUAD
Private MaskAR() As Boolean

Private Sub CboColor_Click()

Dim ShowThem As Boolean
    If CboColor.ListIndex <> 4 Then ShowThem = True
    Label2.Visible = ShowThem
    ScrSim.Visible = ShowThem
    
End Sub

'Load Picture and resize Form
Private Sub CmdLoad_Click()

Dim X As Long
    cmdlg.Filter = "Pictures;*.bmp,*.gif,*.jpg"

    cmdlg.ShowOpen
    If cmdlg.filename = "" Then Exit Sub
    Set PicOrg = LoadPicture(cmdlg.filename)
    X = PicOrg.Width + PicOrg.Left + 10
    If X < 410 Then X = 410
    Me.Width = X * Screen.TwipsPerPixelX
    Me.Height = (PicOrg.Height + PicOrg.Top + 50) * Screen.TwipsPerPixelY

End Sub

'Linear Repair
Private Sub CmdRepair_Click()

Dim X As Long
Dim Y As Long
Dim r As Long
Dim g As Long
Dim b As Long
Dim a As Long
Dim ctr As Long
Dim ctrOld As Long
    
'Get the Picture as array (for faster operation)
    Pic2Array PicOrg, PicAr1

'Create a Maskpicture
    PicDifference
    
    Do
    ctr = 0
'Move Thru the Picture
    For X = 1 To PicOrg.Width - 2
        For Y = 1 To PicOrg.Height - 2
'scann for transparent color
            If MaskAR(X, Y) Then
'We found a transparent pixel
                ctr = ctr + 1
'Check if there is a Colored Pixel
'If yes then add the color
                a = 0
                If MaskAR(X - 1, Y) = False Then
                    a = a + 1
                    r = PicAr1(X - 1, Y).rgbRed
                    g = PicAr1(X - 1, Y).rgbgreen
                    b = PicAr1(X - 1, Y).rgbBlue
                End If
                If MaskAR(X + 1, Y) = False Then
                    a = a + 1
                    r = r + PicAr1(X + 1, Y).rgbRed
                    g = g + PicAr1(X + 1, Y).rgbgreen
                    b = b + PicAr1(X + 1, Y).rgbBlue
                End If
                If MaskAR(X, Y - 1) = False Then
                    a = a + 1
                    r = r + PicAr1(X, Y - 1).rgbRed
                    g = g + PicAr1(X, Y - 1).rgbgreen
                    b = b + PicAr1(X, Y - 1).rgbBlue
                End If
                If MaskAR(X, Y + 1) = False Then
                    a = a + 1
                    r = r + PicAr1(X, Y + 1).rgbRed
                    g = g + PicAr1(X, Y + 1).rgbgreen
                    b = b + PicAr1(X, Y + 1).rgbBlue
                End If
'If we have 2 or mor colored pixels arround then
'fill the transparent pixel with a coombination
                If a > 1 Then
                    r = r / a
                    g = g / a
                    b = b / a
                    PicAr1(X, Y).rgbRed = CByte(r)
                    PicAr1(X, Y).rgbgreen = CByte(g)
                    PicAr1(X, Y).rgbBlue = CByte(b)
                    MaskAR(X, Y) = False
'we removed a transparent pixel
                    ctr = ctr - 1
                End If
                r = 0
                g = 0
                b = 0
            End If
        Next Y
    Next X
    
'Check if we did a new scan without a new result
    If ctr <> ctrOld Then
        ctrOld = ctr
        Else
        If ctr > 0 Then MsgBox "Some parts could not be fixed"
        Exit Do
    End If
    
'if there is still a transparent pixel open RESTART
Loop Until ctr < 1
Array2Pic PicOrg, PicAr1
PicOrg.Refresh

End Sub

'repair the Picture with a Circle Blur Fill
'thru the circle we go from the border to the middle
'In most cases this brings a better result
Private Sub CmdRepCirc_Click()

Dim X As Long
Dim Y As Long
Dim r As Long
Dim g As Long
Dim b As Long
Dim a As Long
Dim ctr As Long
Dim ctrOld As Long
Dim Direction As Long
Dim ErrCnt As Long
Dim TmpX As Long
Dim TmpY As Long
Dim X1 As Long
Dim Y1 As Long

'Get the Picture
    Pic2Array PicOrg, PicAr1

'Create a maskpicture
    PicDifference
    Me.MousePointer = 11
    
    Do
    ctr = 0
'Move Thru the Picture and search for holes
    For Y = 1 To PicOrg.Height - 2
        For X = 1 To PicOrg.Width - 2
'scann for transparent color
            If MaskAR(X, Y) Then
'We found a transparent pixel
                ctr = ctr + 1
                ErrCnt = 0
                Direction = 1
'First we check to the right
'so move the startpoint to the left
                X1 = X - 1
                Y1 = Y
'Now scan for an empty Pixel
                Do
'Select the direction
                Select Case Direction
                    Case 1
                    TmpX = 1
                    TmpY = 0
                    Case 2
                    TmpX = 1
                    TmpY = 1
                    Case 3
                    TmpX = 0
                    TmpY = 1
                    Case 4
                    TmpX = -1
                    TmpY = 1
                    Case 5
                    TmpX = -1
                    TmpY = 0
                    Case 6
                    TmpX = -1
                    TmpY = -1
                    Case 7
                    TmpX = 0
                    TmpY = -1
                    Case 8
                    TmpX = 1
                    TmpY = -1
                End Select
                
                If X1 + TmpX > PicOrg.Width - 2 Or X1 + TmpX < 1 Or Y1 + TmpY > PicOrg.Height - 2 Or Y1 + TmpY < 1 Then
                    ErrCnt = ErrCnt + 1
                    Else
'Search for a new empty Pixel
                    If MaskAR(X1 + TmpX, Y1 + TmpY) Then
'Set a new Startpoint
                        X1 = X1 + TmpX
                        Y1 = Y1 + TmpY
'Check if there is a Colored Pixel
'If yes then add the color
                        a = 0
                        If MaskAR(X1 - 1, Y1) = False Then
                            a = a + 1
                            r = PicAr1(X1 - 1, Y1).rgbRed
                            g = PicAr1(X1 - 1, Y1).rgbgreen
                            b = PicAr1(X1 - 1, Y1).rgbBlue
                        End If
                        If MaskAR(X1 + 1, Y1) = False Then
                            a = a + 1
                            r = r + PicAr1(X1 + 1, Y1).rgbRed
                            g = g + PicAr1(X1 + 1, Y1).rgbgreen
                            b = b + PicAr1(X1 + 1, Y1).rgbBlue
                        End If
                        If MaskAR(X1, Y1 - 1) = False Then
                            a = a + 1
                            r = r + PicAr1(X1, Y1 - 1).rgbRed
                            g = g + PicAr1(X1, Y1 - 1).rgbgreen
                            b = b + PicAr1(X1, Y1 - 1).rgbBlue
                        End If
                        If MaskAR(X1, Y1 + 1) = False Then
                            a = a + 1
                            r = r + PicAr1(X1, Y1 + 1).rgbRed
                            g = g + PicAr1(X1, Y1 + 1).rgbgreen
                            b = b + PicAr1(X1, Y1 + 1).rgbBlue
                        End If
'If we have 2 or mor colored pixels arround then
'fill the transparent pixel with a coombination
                        If a > 1 Then
                            r = r / a
                            g = g / a
                            b = b / a
                            PicAr1(X1, Y1).rgbRed = CByte(r)
                            PicAr1(X1, Y1).rgbgreen = CByte(g)
                            PicAr1(X1, Y1).rgbBlue = CByte(b)
                            MaskAR(X1, Y1) = False
'we removed a transparent pixel
                            ctr = ctr - 1
                            Direction = Direction - 1
                            If Direction = 0 Then Direction = 8
'We found one so reset the error counter
                            ErrCnt = 0
                            Else
                            ErrCnt = ErrCnt + 1
                            Direction = Direction + 1
                            X1 = X1 - TmpX
                            Y1 = Y1 - TmpY
                        End If
                        r = 0
                        g = 0
                        b = 0
                        Else
                        ErrCnt = ErrCnt + 1
                        Direction = Direction + 1
                    End If
                End If
                If Direction = 9 Then Direction = 1
            Loop Until ErrCnt = 8
        End If
    Next X
Next Y

If ctr <> ctrOld Then
    ctrOld = ctr
    Else
    If ctr > 0 Then MsgBox "Some parts could not be fixed"
    Exit Do
End If

Loop Until ctr < 1

Me.MousePointer = 0
Array2Pic PicOrg, PicAr1
PicOrg.Refresh

End Sub

'Create The Mask
'there are some different Method´s to create it
Private Sub PicDifference()

Dim X As Long
Dim Y As Long
Dim ctr1 As Long
Dim ctr2 As Long
Dim fnd As Boolean
Dim ActColor As Byte
Dim ScrVal As Long
    ReDim MaskAR(0 To PicOrg.Width - 1, 0 To PicOrg.Height - 1) As Boolean
    ScrVal = ScrSim.Value * 3
    ActColor = SameColor(TR, TG, TB)

    For X = 0 To PicOrg.Width - 1
        For Y = 0 To PicOrg.Height - 1
            ctr1 = ctr1 + 1
            Select Case CboColor.ListIndex
                Case 0
                If SimilarColor(PicAr1(X, Y).rgbRed, PicAr1(X, Y).rgbgreen, PicAr1(X, Y).rgbBlue, TR, TG, TB, ScrSim.Value) Then fnd = True
                Case 1
                If coldiff(PicAr1(X, Y).rgbRed, PicAr1(X, Y).rgbgreen, PicAr1(X, Y).rgbBlue, TR, TG, TB) < ScrVal Then fnd = True
                Case 2
                If brghtdiff(PicAr1(X, Y).rgbRed, PicAr1(X, Y).rgbgreen, PicAr1(X, Y).rgbBlue, TR, TG, TB) < ScrSim.Value Then fnd = True
                Case 3
                If lumdiff(PicAr1(X, Y).rgbRed, PicAr1(X, Y).rgbgreen, PicAr1(X, Y).rgbBlue, TR, TG, TB) * 12.14 < ScrVal Then fnd = True
                Case 4
                If SameColor(PicAr1(X, Y).rgbRed, PicAr1(X, Y).rgbgreen, PicAr1(X, Y).rgbBlue) = ActColor Then fnd = True
            End Select
            MaskAR(X, Y) = fnd
            If fnd Then ctr2 = ctr2 + 1
            fnd = False
        Next Y
    Next X
    LblDamaged = CLng(ctr2 * 100 / ctr1) & " % repaired"

End Sub
'Check if a Color is ind a range X% from the actual point
Private Function SimilarColor(ByVal Red1 As Long, ByVal Green1 As Long, ByVal Blue1 As Long, ByVal Red2 As Long, ByVal Green2 As Long, ByVal Blue2 As Long, ByVal ADif As Long) As Boolean

'Check if the color is in our range
    If Abs(Red1 - Red2) <= ADif And Abs(Green1 - Green2) <= ADif And Abs(Blue1 - Blue2) <= ADif Then SimilarColor = True

End Function
Private Function SameColor(Red As Byte, Blue As Byte, Green As Byte) As Byte

Dim Tmp As Byte
    If Red > Green Then Tmp = 1

    If Green > Blue Then Tmp = Tmp + 10
    If Red > Blue Then Tmp = Tmp + 100
    SameColor = Tmp

End Function
'Found this functions (as php) on
'http://www.splitbrain.org/blog/2008-09/18-calculating_color_contrast_with_php
Function coldiff(R1, G1, B1, R2, G2, B2) As Long

    coldiff = Abs(R1 - R2) + Abs(G1 - G2) + Abs(B1 - B2)

End Function
'Brightness Contrast
Function brghtdiff(R1, G1, B1, R2, G2, B2) As Long

    brghtdiff = Abs(((299 + R1 + 587 * G1 + 114 * B1) / 1000) - ((299 + R2 + 587 * G2 + 114 * B2) / 1000))

End Function
'Luminosity Contrast
Function lumdiff(R1, G1, B1, R2, G2, B2) As Single

Dim L1 As Single
Dim L2 As Single
    L1 = 0.2126 * (R1 / 255) ^ 2.2 + 0.7152 * (G1 / 255) ^ 2.2 + 0.0722 * (B1 / 255) ^ 2.2

    L2 = 0.2126 * (R2 / 255) ^ 2.2 + 0.7152 * (G2 / 255) ^ 2.2 + 0.0722 * (B2 / 255) ^ 2.2
    If L1 > L2 Then
        lumdiff = (L1 + 0.05) / (L2 + 0.05)
        Else
        lumdiff = (L2 + 0.05) / (L1 + 0.05)
    End If

End Function

'Get a Picture as Array
Private Sub Pic2Array(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)

    GetObject PicBox.Image, Len(PicInfo), PicInfo
    ReDim PicArray(0 To PicInfo.bmWidth - 1, 0 To PicInfo.bmHeight - 1) As RGBQUAD
    GetBitmapBits PicBox.Image, PicInfo.bmWidth * PicInfo.bmHeight * 4, PicArray(0, 0)

End Sub
'Write a Array to a Picture
Private Sub Array2Pic(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)

    GetObject PicBox.Image, Len(PicInfo), PicInfo
    SetBitmapBits PicBox.Image, PicInfo.bmWidth * PicInfo.bmHeight * 4, PicArray(0, 0)

End Sub
Private Sub GetRGB(col As Long, Red, Green, Blue)

    Red = col Mod 256
    Green = ((col And &HFF00) \ 256) Mod 256
    Blue = (col And &HFF0000) \ 65536

End Sub

Private Sub Form_Load()

    CboColor.ListIndex = 0

End Sub
Private Sub PicOrg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GetRGB PicOrg.Point(X, Y), TR, TG, TB
    Shape1.FillColor = RGB(TR, TG, TB)

End Sub
Private Sub ScrSim_Change()

    Label2 = "Difference " & ScrSim.Value

End Sub
