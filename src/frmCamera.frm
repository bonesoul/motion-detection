VERSION 5.00
Object = "{DF6D6558-5B0C-11D3-9396-008029E9B3A6}#1.0#0"; "ezVidC60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{49F811F7-6005-4AAF-AE00-9D98766A6E26}#1.0#0"; "NTGraph.ocx"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Begin VB.Form frmSurveillance 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hareket Takip"
   ClientHeight    =   6675
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   863
   ShowInTaskbar   =   0   'False
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl vbalExplorerBarCtl1 
      Height          =   6495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11456
      BackColorEnd    =   0
      BackColorStart  =   0
   End
   Begin NTGRAPHLib.NTGraph graph 
      Height          =   2175
      Left            =   2880
      TabIndex        =   5
      Top             =   4320
      Width           =   9975
      _Version        =   65536
      _ExtentX        =   17595
      _ExtentY        =   3836
      _StockProps     =   194
      FrameStyle      =   0
      ControlFrameColor=   12632256
      ElementCount    =   1
      ElementLineColor=   0
      ElementPointColor=   0
      ElementName     =   "Element-0"
      BeginProperty TickFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty IdentFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PlotAreaPicture =   "frmCamera.frx":0000
      ControlFramePicture=   "frmCamera.frx":001C
   End
   Begin MSComctlLib.Slider ToleranceSlider 
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   3840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      _Version        =   393216
      LargeChange     =   100
      SmallChange     =   25
      Max             =   1000
      TickStyle       =   1
      TickFrequency   =   100
      TextPosition    =   1
   End
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   7920
      Top             =   3840
   End
   Begin VB.PictureBox picCamera 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   7920
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   328
      TabIndex        =   1
      Top             =   240
      Width           =   4920
   End
   Begin vbVidC60.ezVidCap VidCap 
      Height          =   3660
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   6456
      AutoSize        =   0   'False
      BackColor       =   15299894
      BorderStyle     =   0
      StreamMaster    =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Histogram"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Motion View:"
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Live:"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmSurveillance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private graph_x_cord



Private Sub init_graph()
    With graph
       .PlotAreaColor = vbBlack
       .Caption = ""
       .XLabel = ""
       .YLabel = ""
         
       .ClearGraph 'delete all elements and create a new one
       .ElementLineColor = RGB(255, 255, 0)

    End With
    
    graph_x_cord = 0
End Sub



Private Sub Form_Load()

Dim cBar As cExplorerBar
Dim cItem As cExplorerBarItem

   With vbalExplorerBarCtl1
      .Redraw = False
      .UseExplorerStyle = False
      .BackColorStart = vbWhite
      .BackColorEnd = vbBlack
      
      Set cBar = .Bars.Add(, "camera", "Kamera")
      cBar.IsSpecial = True
      cBar.ToolTipText = "Kamera Kontrolleri"
      cBar.IconIndex = 0
      
      Set cItem = cBar.Items.Add(, "tolerans", "Hareket Toleransý", 0)
      cItem.ItemType = eItemText
      Set cItem = cBar.Items.Add(, "tolerans_kontrol", "", 0)
      cItem.ItemType = eItemControlPlaceHolder
      cItem.Control = Me.ToleranceSlider
      
      Set cItem = cBar.Items.Add(, "video_kaynagi", "Video Kaynaðý", 0)
      cItem.ItemType = eItemLink
         
      Set cItem = cBar.Items.Add(, "video_bicimi", "Video Biçimi", 0)
      cItem.ItemType = eItemLink
      
      .Redraw = True
   End With

    ToleranceSlider.Value = 700
    mdTriger = ToleranceSlider.Value
    'Temporary storage for video
    VidCap.CaptureFile = App.Path & "\capture.avi"
    
    init_graph
    

    
End Sub

Private Sub tmrMain_Timer()
    'Main loop
    'Display the current frame in the PictureBox and detect any motion
    
    If VidCap.CapSingleFrame Then
        picCamera.Cls
        VidCap.SaveDIB VidCap.CaptureFile
        Set picCamera.Picture = LoadPicture(VidCap.CaptureFile)
        
        'Detect motion
        GetMotion
        
        
        
        'Delete temporary video file
        Kill VidCap.CaptureFile
    End If
End Sub

Sub GetMotion()
    Dim ColorSumStr As String           'sum of pixel color
    Dim ColorRedStr As String           'red
    Dim ColorGreenStr As String         'green
    Dim ColorBlueStr As String          'blue
    Dim ColorRedDec As Single           'red
    Dim ColorGreenDec As Single         'green
    Dim ColorBlueDec As Single          'blue
    Dim PixX As Single                  'curent pixel X
    Dim PixY As Single                  'curent pixel Y
    Dim AveragePixel(5) As Single       'Average color from 6 pixels
    Static Counter As Single            'counter
    Dim AverageSum As Single            'Average sum of all colors
    
    Dim BoxesX As Single                'how many 'detection boxes - x axis
    Dim BoxesY As Single                'how many 'detection boxes - y axis
    Dim AveragePixelLoop As Single      'defines how many frames does this sub compare
    
    BoxesX = 16                         'from 1 to 50
    BoxesY = 16                         'from 1 to 50
    AveragePixelLoop = 30               'from 1 to 250
    
    Dim motion_evidince_count As Integer
    
    Dim Repeat As Single
    Dim Px As Single, Py As Single
    
    motion_evidince_count = -1
    
    For Px = 0 To (picCamera.Width) Step Int(picCamera.Width / BoxesX)
        For Py = 0 To (picCamera.Height) Step Int(picCamera.Height / BoxesY)
            PixX = Fix(Px / (picCamera.Width / BoxesX))
            PixY = Fix(Py / (picCamera.Height / BoxesY))
            
            For Repeat = 0 To 5
                ColorSumStr = Right$("000000" + Hex(GetPixel(picCamera.hdc, Px + Repeat, Py + Repeat)), 6)
                ColorRedStr = Mid$(ColorSumStr, 5, 2)
                ColorGreenStr = Mid$(ColorSumStr, 3, 2)
                ColorBlueStr = Mid$(ColorSumStr, 1, 2)
                ColorRedDec = Val("&H" + ColorRedStr)
                ColorGreenDec = Val("&H" + ColorGreenStr)
                ColorBlueDec = Val("&H" + ColorBlueStr)
                AveragePixel(Repeat) = ColorRedDec + ColorGreenDec + ColorBlueDec
            Next
            
            Counter = Counter + 1
            
            If Counter = AveragePixelLoop Then Counter = 1
            
            mdSample(PixX, PixY, 0) = 0
            mdSample(PixX, PixY, Counter) = 0
            
            For Repeat = 0 To 5
                mdSample(PixX, PixY, 0) = mdSample(PixX, PixY, 0) + AveragePixel(Repeat)
                mdSample(PixX, PixY, Counter) = mdSample(PixX, PixY, 0) + AveragePixel(Repeat)
            Next
            
            AverageSum = 0
            
            For Repeat = 1 To AveragePixelLoop
                AverageSum = AverageSum + mdSample(PixX, PixY, Repeat)
            Next
            
            AverageSum = AverageSum / AveragePixelLoop
            
            If Abs(mdSample(PixX, PixY, 0) - AverageSum) > mdTriger Then
                picCamera.Line (Px - 4, Py - 4)-Step((picCamera.Width / BoxesX) - 4, (picCamera.Height / BoxesY) - 4), , B
                motion_evidince_count = motion_evidince_count + 1
            End If
        Next
    Next
    
    graph_x_cord = graph_x_cord + 1
    If graph_x_cord > 100 Then
        init_graph
    End If
    graph.PlotY motion_evidince_count, 0
    graph.SetRange 0, 100, 0, 400
    
End Sub


Private Sub ToleranceSlider_Change()
mdTriger = ToleranceSlider.Value
End Sub



Private Sub vbalExplorerBarCtl1_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
If LCase(itm.Key) = "video_kaynagi" Then
    VidCap.ShowDlgVideoSource
ElseIf LCase(itm.Key) = "video_bicimi" Then
    VidCap.ShowDlgVideoFormat
End If
End Sub
