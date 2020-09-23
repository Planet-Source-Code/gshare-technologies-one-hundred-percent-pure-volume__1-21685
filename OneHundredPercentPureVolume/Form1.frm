VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Volume"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   2505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRealSlider 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   165
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   420
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   30
      Width           =   270
   End
   Begin VB.Timer picSlideTMR 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2670
      Top             =   1815
   End
   Begin VB.PictureBox picActualBackground 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   45
      Picture         =   "Form1.frx":0662
      ScaleHeight     =   420
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   30
      Width           =   2415
      Begin VB.PictureBox picBackGround 
         Height          =   495
         Left            =   1305
         ScaleHeight     =   435
         ScaleWidth      =   1575
         TabIndex        =   1
         Top             =   705
         Visible         =   0   'False
         Width           =   1635
         Begin VB.PictureBox picFakeSlider 
            Height          =   255
            Left            =   525
            ScaleHeight     =   195
            ScaleWidth      =   450
            TabIndex        =   2
            Top             =   90
            Width           =   510
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MovePicture As Boolean
Public initialPicX As Integer

Dim XPicScroll, YPicScroll, ZPicScroll As Long
Dim Ratio As Integer
' Variables for the Slider Graphics

Dim VolRX9 As Long
'The almighty volume value itself


Private Sub Form_Load()

    picBackGround.Left = picRealSlider.Width

    picBackGround.Width = picActualBackground.Width - picRealSlider.Width - picRealSlider.Width
    Ratio = picBackGround.Left + picActualBackground.Left

    picFakeSlider.Left = picRealSlider.Left + Ratio
    picFakeSlider.Width = picRealSlider.Width
    
    picRealSlider.Left = picActualBackground.Left
    
    Call InitGetVolume
    
    
End Sub

Private Sub picRealSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picSlideTMR.Enabled = True

    If Button = vbLeftButton Then
  
        MovePicture = True
        initialPicX = X
        
    End If
    
End Sub

Private Sub picRealSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If MovePicture Then
        
        picRealSlider.Left = picRealSlider.Left - (initialPicX - X)
        picFakeSlider.Left = picRealSlider.Left - Ratio
              
   End If
   
End Sub

Private Sub picRealSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   MovePicture = False
   picSlideTMR.Enabled = False
   
End Sub

Private Sub picSlideTMR_Timer()


        If picRealSlider.Left <= picActualBackground.Left Then
            
            MovePicture = False
            picSlideTMR.Enabled = False
            picRealSlider.Left = picActualBackground.Left
            
        End If
       
        
        If picRealSlider.Left >= picActualBackground.Left + picActualBackground.Width - picRealSlider.Width Then
            
            picRealSlider.Left = picActualBackground.Left + picActualBackground.Width - picRealSlider.Width
            picSlideTMR.Enabled = False
            MovePicture = False
        
        End If
        

         XPicScroll = picFakeSlider.Left + picFakeSlider.Width
         YPicScroll = picBackGround.Left + picBackGround.Width
         ZPicScroll = Int(XPicScroll / YPicScroll * 100)
        

         
         'Calculate the percentage of movement
         
         'X divided by Y * 100 where Y = the proportion
         'against it's position width of the background
         'Something like that anyway ;)
         
         
         If ZPicScroll < 0 Then ZPicScroll = 0
         
         'Just to Make Sure, ALL Languages Sometimes Mess
         'up on timed events so we hard code a "stop"
         'effect on the range :)
         
         
         If ZPicScroll > 100 Then ZPicScroll = 100
         
         'Just to Make Sure, ALL Languages Sometimes Mess
         'up on timed events so we hard code a "stop"
         'effect on the range :)
         
         Me.Caption = "Volume :" & ZPicScroll & "%"
         
         VolRX9 = CLng(ZPicScroll) * 65535 / 100
         SetVolumeControl SetVolHmixer, SetVolCtrl, VolRX9
        
  
End Sub
