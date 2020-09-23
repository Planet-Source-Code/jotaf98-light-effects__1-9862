VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "''Light'' - Made by Jotaf98"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "Light.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Randomly get a pre-defined cool color"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6840
      TabIndex        =   18
      Text            =   "40"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6840
      TabIndex        =   16
      Text            =   "15"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6840
      TabIndex        =   13
      Text            =   "5"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6840
      TabIndex        =   12
      Text            =   "50"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Text            =   "230"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6840
      TabIndex        =   1
      Text            =   "200"
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   120
      ScaleHeight     =   276
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   0
      Top             =   120
      Width           =   6225
   End
   Begin VB.Label Label15 
      Caption         =   ";)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   6090
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      X1              =   3240
      X2              =   3240
      Y1              =   4440
      Y2              =   6240
   End
   Begin VB.Label Label14 
      Caption         =   $"Light.frx":08CA
      Height          =   855
      Left            =   3480
      TabIndex        =   20
      Top             =   5400
      Width           =   3855
   End
   Begin VB.Label Label13 
      Caption         =   "Blue Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   19
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Green Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Red Brightness:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Radius:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   3225
      X2              =   3225
      Y1              =   4440
      Y2              =   6240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000016&
      X1              =   3240
      X2              =   120
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   3240
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label7 
      Caption         =   "and vote me!"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.planet-source-code.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "If you liked this effect, please go to"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Then, you can play around with them, clicking the picture again to see it in action!"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Click the picture to see the effect with the default values."
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Label Label8 
      Caption         =   "Made by Jotaf98 (João F. S. Henriques)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label Label9 
      Caption         =   "E-mail me at: jotaf98@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CoolColors(9, 2) As Byte 'Where the "cool colors" will be stored
Dim RndCoolColor As Byte '"Cool color" got at random

Private Sub RedrawPicture() 'Redraws the picture
    On Error GoTo ErrHandler
    Picture1.Picture = LoadPicture(App.Path & "\Geyser.jpg")
    Exit Sub
    
ErrHandler:
    MsgBox "There was an error displaying the image.", vbExclamation + vbOKOnly
    Picture1.Cls
End Sub

Private Sub Command1_Click()
    RndCoolColor = Rnd * UBound(CoolColors)
    
    MsgBox "Try this one:" & Chr(13) & Chr(13) & _
      "  Red: " & CoolColors(RndCoolColor, 0) & Chr(13) & _
      "  Green: " & CoolColors(RndCoolColor, 1) & Chr(13) & _
      "  Blue: " & CoolColors(RndCoolColor, 2)
End Sub

Private Sub Form_Load()
    Randomize Timer
    
    Form1.MousePointer = vbHourglass 'Change pointer to hourglass
    Form1.Caption = "''Light'' - Initializing ''Cool Colors'' databank..." 'Change caption
    
    'Initialize the "cool colors" (they're 10)
    InitCoolColor 0, 25, 15, 50
    InitCoolColor 1, 50, 20, 10
    InitCoolColor 2, 15, 40, 5
    InitCoolColor 3, 10, 20, 60
    InitCoolColor 4, 5, 15, 40
    InitCoolColor 5, 60, 15, 35
    InitCoolColor 6, 30, 20, 15
    InitCoolColor 7, 15, 0, 60
    InitCoolColor 8, 60, 15, 0
    InitCoolColor 9, 7, 20, 0
    
    'Redraw the main image
    Form1.Caption = "''Light'' - Redrawing image..." 'Change caption
    RedrawPicture
    
    Form1.Caption = "''Light'' - Made by Jotaf98" 'Change caption to default
    Form1.MousePointer = vbDefault 'Change pointer to default
End Sub

Private Sub InitCoolColor(Index, RedB, GreenB, BlueB)
    'Initializes a "cool color"
    CoolColors(Index, 0) = RedB
    CoolColors(Index, 1) = GreenB
    CoolColors(Index, 2) = BlueB
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    
    Temp = MsgBox("If you liked this effect, please go to Planet Source Code to vote me ; )" & Chr(13) & Chr(13) & "Visit ''http://www.planet-source-code.com'' now?", vbQuestion + vbYesNo)
    
    If Temp = vbYes Then
        'This code snippet opens an .url file.
        'Please give me some credit if you use it ;)
        Shell "rundll32.exe shdocvw.dll,OpenURL " & App.Path & "\Planet Source Code.url"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Sorry, ''Planet Source Code.url'' was not found.", vbExclamation + vbOKOnly
End Sub

Private Sub Label4_Click()
    On Error GoTo ErrHandler
    
    Temp = MsgBox("Visit ''http://www.planet-source-code.com'' now?", vbQuestion + vbYesNo)
    
    If Temp = vbYes Then
        'This code snippet opens an .url file.
        'Please give me some credit if you use it ;)
        Shell "rundll32.exe shdocvw.dll,OpenURL " & App.Path & "\Planet Source Code.url"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Sorry, ''Planet Source Code.url'' was not found.", vbExclamation + vbOKOnly
End Sub

Private Sub Picture1_Click()
    Form1.MousePointer = vbHourglass 'Change pointer to hourglass
    Form1.Caption = "''Light'' - Redrawing image..." 'Change caption
    RedrawPicture
    Form1.Caption = "''Light'' - Drawing light..." 'Change caption
    
    'Draw the light
    DrawLight Picture1.hdc, Int(Text1.Text), _
      Int(Text2.Text), Int(Text4.Text), _
      Int(Text5.Text), Int(Text6.Text), _
      Int(Text3.Text), Int(Text3.Text / 2)
    '                                  ^ I decided to make the number
    '                                    of segments half the radius.
    '                                    Usually higher numbers make
    '                                    better effects, but are slower.
    
    Form1.Caption = "''Light'' - Made by Jotaf98" 'Change caption to default
    Form1.MousePointer = vbDefault 'Change pointer to default
End Sub

'This is the algorythm for drawing a full circle.
'It's much faster than the Circle function!
'It is provided to you as a bonus; use it as you
'want (you don't need to give me credit for this,
'I found it in one of my Dad's books) ;)
'
'TOP SECRET TIP: If instead of "<=" you use ">="
'(without the quotes -- ""), it will draw just
'outside of the circle. And if instead of
'"If ((X - OrigX..." you use the following
'2 lines, it will draw only the border!
'
'            Temp = ((X - OrigX) * (X - OrigX)) + ((Y - OrigY) * (Y - OrigY))
'            If Temp < (R * R) + R And Temp > (R * R) - R Then
'

Private Sub DrawACircle() 'Just a placeholder, so you don't see it commented
    Dim X As Long 'X Counter
    Dim Y As Long 'Y Counter
    Dim OrigX As Long 'Original X coordinates
    Dim OrigY As Long 'Original Y coordinates
    Dim R As Long 'The circle's radius
    
    OrigX = 100
    OrigY = 200
    R = 25

    'Will draw a blue circle in Picture1 at X
    '100 and Y 200; radius 25:

    For X = OrigX - R To OrigX + R
        For Y = OrigY - R To OrigY + R
            If ((X - OrigX) * (X - OrigX)) + ((Y - OrigY) * (Y - OrigY)) <= (R * R) Then
                SetPixel Picture1.hdc, X, Y, vbBlue
            End If
        Next Y
    Next X
End Sub
