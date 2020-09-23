VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmShapedFocusRect 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Use Custom Focus"
      Height          =   255
      Index           =   1
      Left            =   1965
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3975
      Width           =   1875
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use DrawFocusRect"
      Height          =   255
      Index           =   0
      Left            =   1965
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3630
      Value           =   -1  'True
      Width           =   1875
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Change Focus Color"
      Height          =   345
      Index           =   1
      Left            =   3615
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3180
      Width           =   1680
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Change Bkg Color"
      Height          =   345
      Index           =   0
      Left            =   1890
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3195
      Width           =   1680
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H00C00000&
      Height          =   1275
      Left            =   135
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   4
      Top             =   3180
      Width           =   1710
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4665
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   3480
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   3
      Top             =   1425
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   2
      Left            =   3510
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   150
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   1
      Left            =   1845
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   135
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   0
      Left            =   180
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   105
      Width           =   1500
   End
   Begin VB.Line Line1 
      X1              =   570
      X2              =   4710
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   195
      TabIndex        =   7
      Top             =   1110
      Width           =   4905
   End
   Begin VB.Label Label2 
      Height          =   330
      Left            =   3480
      TabIndex        =   6
      Top             =   2490
      Width           =   1530
   End
   Begin VB.Label Label1 
      Height          =   675
      Left            =   480
      TabIndex        =   5
      Top             =   1635
      Width           =   2565
   End
End
Attribute VB_Name = "frmShapedFocusRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' You may ask, why not just use .CLS to erase the focus rectangle when LostFocus occurs?
' Well, lets say the picturebox was really an offscreen DC.  And the DC was painted with
' other stuff, like gradients, some images, text, or other things. Just to change from
' focus & unfocused, you wouldn't want to completely redraw the DC because that would
' be far more time consuming than simply "erasing" the last drawn thing. And really the
' same applies for pictureboxes too.

' This is where the power of XOR painting comes into play.  By painting something with
' XOR you are changing pixel values by adding or removing a known byte or range of bytes,
' and if the same thing is painted again using XOR and using the same byte or range of bytes,
' you actually erase what was just painted, restoring what was painted over.  This is
' completely different than simply painting using SrcCopy where you permanently lose whatever
' was painted over. For focus rectangles, this is idea because you are basically painting
' a single color graphic over the DC.

' With this example, I have provided examples on how to use FrameRgn to draw focus
' rectangles that are shapes, how to use DrawFocusRect API and dictate the color of
' the rectangle, and also how to draw your own custom/solid focus rect in the XOR fashion.


' Used for the "custom" focus rectangle
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type


' used for CreateStretchedRegion
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

' region creation APIs
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

' misc
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function SetROP2 Lib "gdi32.dll" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long

' Region drawing APIs
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long

' DrawFocusRect only works for rectangles, not shapes
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' local variables
Private focusRgnShaped(0 To 2)  As Long
Private focusBrush As Long

Private Sub cmdColor_Click(Index As Integer)
    On Error GoTo ExitRoutine:
    With CommonDialog1
        .CancelError = True
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        If Index = 0 Then .Color = Picture3.BackColor Else .Color = Picture3.ForeColor
    End With
    CommonDialog1.ShowColor
    If Index = 0 Then Picture3.BackColor = CommonDialog1.Color Else Picture3.ForeColor = CommonDialog1.Color
ExitRoutine:
End Sub

Private Sub Form_Load()

    Picture1(0).ScaleMode = vbPixels
    Picture1(0).AutoRedraw = True
    Picture1(1).ScaleMode = vbPixels
    Picture1(1).AutoRedraw = True
    Picture1(2).ScaleMode = vbPixels
    Picture1(2).AutoRedraw = True
    

    Label1.Caption = "Tab around to show focus rect.  There are no calls to .CLS in the code."
    Label2.Caption = "DrawFocusRect API"
    Label3.Caption = "Above focus 'rects' are shaped using DrawMode manipulation"
    
    Dim hBmp As Long, hRgn As Long, hBrush As Long
    Dim a(0 To 3) As Long
    
    a(0) = 11141205 ' Create a monochrome 8x8 bitmap
    a(1) = 11141205 ' with alternating black/white pixels
    a(2) = 11141205
    a(3) = 11141205
    hBmp = CreateBitmap(8, 8, 1, 1, a(0))
    ' create a brush from the bitmap & delete bmp
    focusBrush = CreatePatternBrush(hBmp)
    DeleteObject hBmp
    
    hBrush = CreateSolidBrush(0&)
    With Picture1(0)
        hRgn = CreateRoundRectRgn(0, 0, .ScaleWidth, .ScaleHeight, 20, 20)
        FrameRgn .hdc, hRgn, hBrush, 1, 1
        focusRgnShaped(0) = CreateStretchRgn(hRgn, .ScaleWidth, .ScaleWidth - 8, .ScaleHeight, .ScaleHeight - 8)
        SetWindowRgn .hWnd, hRgn, True
    End With
    With Picture1(1)
        hRgn = CreateEllipticRgn(0, 0, .ScaleWidth, .ScaleHeight)
        FrameRgn .hdc, hRgn, hBrush, 1, 1
        focusRgnShaped(1) = CreateStretchRgn(hRgn, .ScaleWidth, .ScaleWidth - 8, .ScaleHeight, .ScaleHeight - 8)
        SetWindowRgn .hWnd, hRgn, True
    End With
    With Picture1(2)
        hRgn = CreateRectRgn(0, 0, .ScaleWidth, .ScaleHeight)
        FrameRgn .hdc, hRgn, hBrush, 1, 1
        focusRgnShaped(2) = CreateStretchRgn(hRgn, .ScaleWidth, .ScaleWidth - 8, .ScaleHeight, .ScaleHeight - 8)
        SetWindowRgn .hWnd, hRgn, True
    End With
    DeleteObject hBrush
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If focusBrush Then DeleteObject focusBrush
    If focusRgnShaped(0) Then DeleteObject focusRgnShaped(0)
    If focusRgnShaped(1) Then DeleteObject focusRgnShaped(1)
    If focusRgnShaped(2) Then DeleteObject focusRgnShaped(2)
End Sub

' ***********************************************
'   Shaped DrawFocusRect usage
' ***********************************************

Private Sub Picture1_GotFocus(Index As Integer)
    ToggleFocusRect_Shaped Picture1(Index).hdc, focusRgnShaped(Index)
    Picture1(Index).Refresh
End Sub

Private Sub Picture1_LostFocus(Index As Integer)
    ToggleFocusRect_Shaped Picture1(Index).hdc, focusRgnShaped(Index)
    Picture1(Index).Refresh
End Sub

Private Sub ToggleFocusRect_Shaped(ByVal lDC As Long, ByVal hRgn As Long)
    ' change DC DrawMode (same function as VB's DrawMode)
    Dim lRop As Long
    lRop = SetROP2(lDC, vbXorPen)
    ' frame region
    FrameRgn lDC, hRgn, focusBrush, 1, 1
    SetROP2 lDC, lRop
End Sub

Private Function CreateStretchRgn(srcRgn As Long, OldWidth As Long, NewWidth As Long, _
                                OldHeight As Long, NewHeight As Long) As Long

    Dim xFrm As XFORM, hRgn As Long
    Dim dwCount As Long, pRgnData() As Byte

    With xFrm
        .eM11 = NewWidth / OldWidth
        .eM22 = NewHeight / OldHeight
    End With

    dwCount = GetRegionData(srcRgn, 0, ByVal 0&)
    ReDim pRgnData(1 To dwCount) As Byte
    If dwCount = GetRegionData(srcRgn, dwCount, pRgnData(1)) Then     ' success
        hRgn = ExtCreateRegion(xFrm, dwCount, pRgnData(1))
        If hRgn Then
            OffsetRgn hRgn, (OldWidth - NewWidth) \ 2, (OldHeight - NewHeight) \ 2
            CreateStretchRgn = hRgn
        End If
    End If

End Function

' ***********************************************
'   Normal DrawFocusRect usage
' ***********************************************

Private Sub Picture2_GotFocus()
ToggleFocusRect_Default
End Sub

Private Sub Picture2_LostFocus()
ToggleFocusRect_Default
End Sub

Private Sub ToggleFocusRect_Default()
    Dim r As RECT
    r.Left = 3
    r.Top = 3
    r.Right = Picture2.ScaleWidth - 3
    r.Bottom = Picture2.ScaleHeight - 3
    DrawFocusRect Picture2.hdc, r
End Sub

' ***********************************************
'   Custom DrawFocusRect usage
' ***********************************************

Private Sub Picture3_GotFocus()
    ToggleFocusRect_APIcustom
End Sub

Private Sub Picture3_LostFocus()
    ToggleFocusRect_APIcustom
End Sub

Private Sub ToggleFocusRect_APIcustom()

    Dim r As RECT, pt As POINTAPI
    Dim lBrush As Long, lOld As Long
    
    r.Left = 5
    r.Top = 5
    r.Right = Picture3.ScaleWidth - 5
    r.Bottom = Picture3.ScaleHeight - 5
    
    If Option1(0) = True Then
        ' notice we are tweaking bkg color to be XOR
        ' and in the ELSE below, we tweak forecolor
        
        lBrush = Picture3.ForeColor
        
        lOld = SetBkColor(Picture3.hdc, (Picture3.BackColor Xor lBrush))
        ' don't use Me.BackColor=lColor because VB will change it immediately
        
        DrawFocusRect Picture3.hdc, r
        
        ' replace the object's original backcolor & forecolor
        SetBkColor Picture3.hdc, lOld
        Picture3.ForeColor = lBrush
        
    Else
    
        lBrush = Picture3.ForeColor
        
        Picture3.DrawWidth = 4  ' make pen 4 pixels wide
        Picture3.ForeColor = (Picture3.BackColor Xor lBrush)
        ' change the forecolor to an XOR version
        ' If needed for offscreen: CreatePen(0, penWidth, GetBkColor(hDC) Xor BrushColor)
        '   then select the pen into the hDC
        
        
        lOld = SetROP2(Picture3.hdc, vbXorPen)
        
        With Picture3
            MoveToEx .hdc, r.Left, r.Top, pt
            LineTo .hdc, r.Right, r.Top
            LineTo .hdc, r.Right, r.Bottom
            LineTo .hdc, r.Left, r.Bottom
            LineTo .hdc, r.Left, r.Top
        End With
        
        ' If needed for offscreen DC, simply unselect & destroy pen
        Picture3.ForeColor = lBrush
        Picture3.DrawWidth = 1  ' reset pen to 1 pixel
        
        SetROP2 Picture3.hdc, lOld
        
    End If
End Sub
