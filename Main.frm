VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Real-time Image Levels - tannerhelland@hotmail.com"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   549
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDispHistogram 
      Appearance      =   0  'Flat
      Caption         =   "Display Histogram"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   6015
   End
   Begin VB.Frame frmLevels 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Levels:"
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   6015
      Begin VB.CommandButton cmdReset 
         Appearance      =   0  'Flat
         Caption         =   "Reset"
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.HScrollBar hsInM 
         Height          =   220
         Left            =   1080
         Max             =   254
         Min             =   1
         TabIndex        =   18
         Top             =   840
         Value           =   127
         Width           =   4455
      End
      Begin VB.HScrollBar hsInL 
         Height          =   220
         Left            =   1080
         Max             =   253
         TabIndex        =   13
         Top             =   600
         Width           =   4455
      End
      Begin VB.HScrollBar hsInR 
         Height          =   220
         Left            =   1080
         Max             =   255
         Min             =   2
         TabIndex        =   12
         Top             =   1080
         Value           =   255
         Width           =   4455
      End
      Begin VB.HScrollBar hsOutR 
         Height          =   220
         Left            =   1080
         Max             =   255
         TabIndex        =   9
         Top             =   2040
         Value           =   255
         Width           =   4455
      End
      Begin VB.HScrollBar hsOutL 
         Height          =   220
         Left            =   1080
         Max             =   255
         TabIndex        =   5
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Label lblRightL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   210
         Left            =   990
         TabIndex        =   22
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label lblMiddleL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   210
         Left            =   960
         TabIndex        =   21
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Midtones:"
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   690
      End
      Begin VB.Label lblMiddleR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "254"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5625
         TabIndex        =   19
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left limit:     0"
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblLeftR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "253"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5625
         TabIndex        =   16
         Top             =   600
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right limit:"
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "255"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5625
         TabIndex        =   14
         Top             =   1080
         Width           =   270
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Input levels:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "255"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5625
         TabIndex        =   10
         Top             =   2040
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right limit:  0"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "255"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5625
         TabIndex        =   6
         Top             =   1800
         Width           =   270
      End
      Begin VB.Label lblOutputL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left limit:    0"
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblOutputLevels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Output levels:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   5775
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      Picture         =   "Main.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   6030
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      Picture         =   "Main.frx":6456
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open Image"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Image Levels example ©2006 by Tanner 'DemonSpectre' Helland
'http://www.tannerhelland.com
'tannerhelland@hotmail.com

'This project is an exact model of how to adjust image levels (identical to
'Photoshop's method).  The code is well-commented, but there are some fairly involved
'math sections.  Don't feel bad if you don't understand all the parabolic stuff ;)

'However, the main routine is a (fairly simple) complete sub that can be instantly
'dropped into any VB project.

'Because a large portion of this project relies on DIB sections, I would recommend
'that you first read "From PSet to DIB Sections - your comprehensive guide to VB
'Graphics Programming."  This article can be downloaded from several places, most
'notably http://www.studentsofgamedesign.com

'For additional cool code and tutorials, check out
'http://www.studentsofgamedesign.com

'Check out my original video game music at
'http://www.tannerhelland.com

Option Explicit

'**************
'  VARIABLES  '
'**************

'Constants required for creating a gamma curve from .1 to 10
Private Const MAXGAMMA As Double = 1.8460498941512
Private Const MIDGAMMA As Double = 0.68377223398334
Private Const ROOT10 As Double = 3.16227766

'Used to track the ratio of the midtones scrollbar, so that when the left and
'right values get changed, we automatically set the midtone to the same ratio
'(i.e. as Photoshop does it)
Dim midRatio As Double

'Whether or not changing the midtone scrollbar is user-generated or program-generated
'(so we only refresh if the user moved it - otherwise we get bad looping)
Dim iRefresh As Boolean


'When the program starts, automatically initialize several things...
Private Sub Form_Load()
    'Upon loading the form, automatically set two histogram variables:
     'Luminance is the default histogram source
     lastHistSource = DRAWMETHOD_LUMINANCE
     'Line graph is the default drawing option
     lastHistMethod = DRAWMETHOD_BARS

    'Also, set the default midtone scrollbar ratio to 1/2
    midRatio = 0.5
    
    '...and allow refreshing
    iRefresh = True
End Sub


'The histogram information will be displayed on a separate form
Private Sub cmdDispHistogram_Click()
    frmHistogram.Show
End Sub


'This will reset the scrollbars to default levels
Private Sub cmdReset_Click()
    'Allow refreshing
    iRefresh = True
    'Set the output levels to (0-255)
    hsOutL.Value = 0
    hsOutR.Value = 255
    'Set the input levels to (0-255)
    hsInL.Value = 0
    hsInR.Value = 255
    FixScrollBars
    'Set the midtone level to default (127)
    midRatio = 0.5
    hsInM.Value = 127
    FixScrollBars
End Sub


'*********************************************************************************
'The following 10 subroutines are for changing/scrolling any of the scrollbars
'on the main form
'*********************************************************************************
Private Sub hsInL_Change()
    FixScrollBars
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub

Private Sub hsInL_Scroll()
    FixScrollBars
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub

Private Sub hsInM_Change()
    If iRefresh = True Then
        midRatio = (CDbl(hsInM.Value) - CDbl(hsInL.Value)) / (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        FixScrollBars True
        MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
    End If
End Sub

Private Sub hsInM_Scroll()
    If iRefresh = True Then
        midRatio = (CDbl(hsInM.Value) - CDbl(hsInL.Value)) / (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        FixScrollBars True
        MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
    End If
End Sub

Private Sub hsInR_Change()
    FixScrollBars
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub

Private Sub hsInR_Scroll()
    FixScrollBars
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub

Private Sub hsOutL_Change()
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub

Private Sub hsOutL_Scroll()
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub

Private Sub hsOutR_Change()
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub

Private Sub hsOutR_Scroll()
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub


'Subroutine for loading new images
Private Sub MnuOpenImage_Click()
    'Common dialog interface
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    'String returned from the common dialog wrapper
    Dim sFile As String
    'This string contains the filters for loading different kinds of images.  Using
    'this feature correctly makes our common dialog box a LOT more pleasant to use.
    Dim cdfStr As String
    cdfStr = "All Compatible Graphics|*.bmp;*.jpg;*.jpeg;*.gif;*.wmf;*.emf;*.ico;*.dib;*.rle|"
    cdfStr = cdfStr & "BMP - Windows Bitmaps only (non-OS2)|*.bmp|DIB - Windows DIBs only (non-OS2)|*.dib|EMF - Enhanced Meta File|*.emf|GIF - Compuserve|*.gif|ICO - Windows Icon|*.ico|JPG - JPEG - JFIF Compliant|*.jpg;*.jpeg|RLE - Windows only (non-Compuserve)|*.rle|WMF - Windows Meta File|*.wmf|All files|*.*"
    'If cancel isn't selected, load a picture from the user-specified file
    If CC.VBGetOpenFileName(sFile, , , , , True, cdfStr, 1, , "Open an image", , frmMain.hWnd, 0) Then
        picBack.Picture = LoadPicture(sFile)
        'This will copy the image, automatically resized, from the background
        'picture box to the foreground one
        Dim fDraw As New FastDrawing
        Dim ImageData() As Byte
        Dim iWidth As Long, iHeight As Long
        iWidth = fDraw.GetImageWidth(frmMain.picBack)
        iHeight = fDraw.GetImageHeight(frmMain.picBack)
        fDraw.GetImageData2D frmMain.picBack, ImageData()
        fDraw.SetImageData2D frmMain.picMain, iWidth, iHeight, ImageData()
        'Reset all of the scrollbars
        hsOutL.Value = 0
        hsOutR.Value = 255
        hsInL.Value = 0
        hsInR.Value = 255
        FixScrollBars
        hsInM.Value = 127
        midRatio = 0.5
        FixScrollBars
    End If
End Sub


'Draw an image based on user-adjusted input and output levels
Private Sub MapImageLevels(srcPic As PictureBox, dstPic As PictureBox, ByVal inLLimit As Long, ByVal inMLimit As Long, ByVal inRLimit As Long, ByVal outLLimit As Long, ByVal outRLimit As Long)

    'This array will hold the image's pixel data
    Dim ImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(frmMain.picBack)
    iHeight = fDraw.GetImageHeight(frmMain.picBack)
    fDraw.GetImageData2D frmMain.picBack, ImageData()
    
    'These variables will hold temporary pixel color values
    Dim R As Long, G As Long, B As Long, L As Long
    
    'Look-up table for the midtone (gamma) leveled values
    Dim gValues(0 To 255) As Double
    
    'WARNING: This next chunk of code is a lot of messy math.  Don't worry too much
    'if you can't make sense of it ;)
    
    'Fill the gamma table with appropriate gamma values (from 10 to .1, ranged quadratically)
    'NOTE: This table is constant, and could be loaded from file instead of generated mathematically every time we run this function
    Dim gStep As Double
    gStep = (MAXGAMMA + MIDGAMMA) / 127
    For x = 0 To 127
        gValues(x) = (CDbl(x) / 127) * MIDGAMMA
    Next x
    For x = 128 To 255
        gValues(x) = MIDGAMMA + (CDbl(x - 127) * gStep)
    Next x
    For x = 0 To 255
        gValues(x) = 1 / ((gValues(x) + 1 / ROOT10) ^ 2)
    Next x
    
    'Because we've built our look-up tables on a 0-255 scale, correct the inMLimit
    'value (from the midtones scroll bar) to simply represent a ratio on that scale
    Dim tRatio As Double
    tRatio = (inMLimit - inLLimit) / (inRLimit - inLLimit)
    tRatio = tRatio * 255
    'Then convert that ratio into a byte (so we can access a look-up table with it)
    Dim bRatio As Byte
    bRatio = CByte(tRatio)
    
    'Calculate a look-up table of gamma-corrected values based on the midtones scrollbar
    Dim gLevels(0 To 255) As Byte
    Dim tmpGamma As Double
    For x = 0 To 255
        tmpGamma = CDbl(x) / 255
        tmpGamma = tmpGamma ^ (1 / gValues(bRatio))
        tmpGamma = tmpGamma * 255
        If tmpGamma > 255 Then
            tmpGamma = 255
        ElseIf tmpGamma < 0 Then
            tmpGamma = 0
        End If
        gLevels(x) = tmpGamma
    Next x
    
    'Look-up table for the input leveled values
    Dim newLevels(0 To 255) As Byte
    
    'Fill the look-up table with appropriately mapped input limits
    Dim pStep As Single
    pStep = 255 / (CSng(inRLimit) - CSng(inLLimit))
    For x = 0 To 255
        If x < inLLimit Then
            newLevels(x) = 0
        ElseIf x > inRLimit Then
            newLevels(x) = 255
        Else
            newLevels(x) = ByteMe(((CSng(x) - CSng(inLLimit)) * pStep))
        End If
    Next x
    
    'Now run all input-mapped values through our midtone-correction look-up
    For x = 0 To 255
        newLevels(x) = gLevels(newLevels(x))
    Next x
    
    'Last of all, remap all image values to match the user-specified output limits
    Dim oStep As Double
    oStep = (CSng(outRLimit) - CSng(outLLimit)) / 255
    For x = 0 To 255
        newLevels(x) = ByteMe(CSng(outLLimit) + (CSng(newLevels(x)) * oStep))
    Next x
    
    
    'Now run a quick loop through the image, adjusting pixel values with the look-up tables
    Dim QuickX As Long
    For x = 0 To iWidth - 1
        QuickX = x * 3
    For y = 0 To iHeight - 1
        'Grab red, green, and blue
        R = ImageData(QuickX + 2, y)
        G = ImageData(QuickX + 1, y)
        B = ImageData(QuickX, y)
        'Correct them all
        ImageData(QuickX + 2, y) = newLevels(R)
        ImageData(QuickX + 1, y) = newLevels(G)
        ImageData(QuickX, y) = newLevels(B)
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D picMain, iWidth, iHeight, ImageData()

End Sub


'Used to make sure the scroll bars have appropriate limits
Private Sub FixScrollBars(Optional midMoving As Boolean = False)
    'Make sure that the input scrollbar values don't overlap, and update the labels
    'to display such
    hsInM.Min = hsInL.Value + 1
    lblMiddleL.Caption = hsInL.Value + 1
    hsInR.Min = hsInL.Value + 2
    lblRightL.Caption = hsInL.Value + 2
    hsInL.Max = hsInR.Value - 2
    lblLeftR.Caption = hsInR.Value - 2
    hsInM.Max = hsInR.Value - 1
    lblMiddleR.Caption = hsInR.Value - 1
    'If the user hasn't moved the midtones scrollbar, attempt to preserve its ratio
    If midMoving = False Then
        iRefresh = False
        Dim newValue As Long
        newValue = hsInL.Value + midRatio * (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        If newValue > hsInM.Max Then
            newValue = hsInM.Max
        ElseIf newValue < hsInM.Min Then
            newValue = hsInM.Min
        End If
        hsInM.Value = newValue
        DoEvents
        iRefresh = True
    End If
End Sub


'Used to restrict values to the (0-255) range
Private Function ByteMe(ByVal val As Long) As Byte
    If val > 255 Then
        ByteMe = 255
    ElseIf val < 0 Then
        ByteMe = 0
    Else
        ByteMe = val
    End If
End Function

