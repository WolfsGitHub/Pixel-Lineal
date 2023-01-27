VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "frmMenu"
   ClientHeight    =   3840
   ClientLeft      =   5070
   ClientTop       =   12900
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   Visible         =   0   'False
   Begin VB.PictureBox picMenuRuler 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   5940
      Picture         =   "Menu.frx":2CFA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   42
      Top             =   2040
      Width           =   270
   End
   Begin VB.PictureBox picMenuFile 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   5
      Left            =   1800
      Picture         =   "Menu.frx":2F24
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   41
      Top             =   1560
      Width           =   270
   End
   Begin VB.PictureBox picMenuFile 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   6
      Left            =   1800
      Picture         =   "Menu.frx":32AE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   40
      Top             =   1920
      Width           =   270
   End
   Begin VB.PictureBox picMenuFile 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   8
      Left            =   1800
      Picture         =   "Menu.frx":3638
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   39
      Top             =   2640
      Width           =   270
   End
   Begin VB.PictureBox picMenuFile 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   7
      Left            =   1800
      Picture         =   "Menu.frx":39C2
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   38
      Top             =   2280
      Width           =   270
   End
   Begin VB.PictureBox picMenuRuler 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   5610
      Picture         =   "Menu.frx":3D4C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   37
      Top             =   2040
      Width           =   270
   End
   Begin VB.PictureBox picMenuRuler 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   5280
      Picture         =   "Menu.frx":40D6
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   36
      Top             =   2040
      Width           =   270
   End
   Begin VB.PictureBox picMenuRuler 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   4950
      Picture         =   "Menu.frx":4460
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   35
      Top             =   2040
      Width           =   270
   End
   Begin VB.PictureBox picMenuRuler 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   4620
      Picture         =   "Menu.frx":47EA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   34
      Top             =   2040
      Width           =   270
   End
   Begin VB.PictureBox picMenuFile 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   1815
      Picture         =   "Menu.frx":4B74
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   33
      Top             =   1200
      Width           =   270
   End
   Begin VB.PictureBox picMenuFile 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   1815
      Picture         =   "Menu.frx":4EFE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   32
      Top             =   825
      Width           =   270
   End
   Begin VB.PictureBox picMenuFile 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   1815
      Picture         =   "Menu.frx":5288
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   31
      Top             =   495
      Width           =   270
   End
   Begin VB.PictureBox picMenuFile 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   1815
      Picture         =   "Menu.frx":5612
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   30
      Top             =   165
      Width           =   270
   End
   Begin VB.PictureBox picMenuPal 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   2310
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   29
      Top             =   2475
      Width           =   2415
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   19
      Left            =   5610
      Picture         =   "Menu.frx":599C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   28
      Top             =   1485
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   17
      Left            =   5610
      Picture         =   "Menu.frx":5D26
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   27
      Top             =   1155
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   16
      Left            =   5610
      Picture         =   "Menu.frx":60B0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   26
      Top             =   825
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   15
      Left            =   5610
      Picture         =   "Menu.frx":643A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   25
      Top             =   495
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   14
      Left            =   5280
      Picture         =   "Menu.frx":67C4
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   24
      Top             =   1485
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   12
      Left            =   5280
      Picture         =   "Menu.frx":6B4E
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   23
      Top             =   1155
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   11
      Left            =   5280
      Picture         =   "Menu.frx":6ED8
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   22
      Top             =   825
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   10
      Left            =   5280
      Picture         =   "Menu.frx":7262
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   21
      Top             =   495
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   9
      Left            =   4950
      Picture         =   "Menu.frx":75EC
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   20
      Top             =   1485
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   7
      Left            =   4950
      Picture         =   "Menu.frx":7976
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   19
      Top             =   1155
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   6
      Left            =   4950
      Picture         =   "Menu.frx":7D00
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   18
      Top             =   825
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   5
      Left            =   4950
      Picture         =   "Menu.frx":808A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   17
      Top             =   495
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   4620
      Picture         =   "Menu.frx":8414
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   16
      Top             =   1485
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   4620
      Picture         =   "Menu.frx":879E
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   15
      Top             =   1155
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   4620
      Picture         =   "Menu.frx":8B28
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   14
      Top             =   825
      Width           =   270
   End
   Begin VB.PictureBox picMenuArrow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   4620
      Picture         =   "Menu.frx":8EB2
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   13
      Top             =   495
      Width           =   270
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   6
      Left            =   165
      Picture         =   "Menu.frx":923C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   12
      Top             =   2475
      Width           =   1500
   End
   Begin VB.PictureBox picTearOff 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   0
      Left            =   2805
      Picture         =   "Menu.frx":95C6
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   11
      Top             =   165
      Width           =   1695
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   12
      Left            =   2280
      Picture         =   "Menu.frx":A350
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   10
      Top             =   2040
      Width           =   1995
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   11
      Left            =   2280
      Picture         =   "Menu.frx":A6DA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   9
      Top             =   1680
      Width           =   1995
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   10
      Left            =   2280
      Picture         =   "Menu.frx":AA64
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   8
      Top             =   1320
      Width           =   1995
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   9
      Left            =   2280
      Picture         =   "Menu.frx":ADEE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   7
      Top             =   960
      Width           =   1995
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   7
      Left            =   2280
      Picture         =   "Menu.frx":B178
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   6
      Top             =   480
      Width           =   1995
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   5
      Left            =   120
      Picture         =   "Menu.frx":B502
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   5
      Top             =   2040
      Width           =   1515
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   120
      Picture         =   "Menu.frx":B88C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   4
      Top             =   960
      Width           =   1515
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   120
      Picture         =   "Menu.frx":BC16
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   3
      Top             =   1680
      Width           =   1515
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   120
      Picture         =   "Menu.frx":BFA0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   2
      Top             =   1320
      Width           =   1515
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   120
      Picture         =   "Menu.frx":C32A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   1
      Top             =   480
      Width           =   1515
   End
   Begin VB.PictureBox picMenuColor 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   120
      Width           =   270
   End
   Begin VB.Menu MRuler 
      Caption         =   "MRuler"
      Begin VB.Menu mnuOrientation 
         Caption         =   "&Vertikal"
      End
      Begin VB.Menu mnuRMagGlass 
         Caption         =   "&Lupe"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuScreenShot 
         Caption         =   "&ScreenShot"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuMagColor 
         Caption         =   "&Farb-Pipette"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuMarker 
         Caption         =   "&Markierer setzen"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorCollection 
         Caption         =   "Gesammelte Farben"
         Begin VB.Menu mnuColorCollectionItems 
            Caption         =   "&&H00000000&&"
            Index           =   0
         End
      End
      Begin VB.Menu mnuColorCodes 
         Caption         =   "Farbanzeige"
         Begin VB.Menu mnuColorCode 
            Caption         =   "HTML"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuColorCode 
            Caption         =   "VB"
            Index           =   1
         End
         Begin VB.Menu mnuColorCode 
            Caption         =   "OLEColor"
            Index           =   2
         End
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Lineal Farben"
         Begin VB.Menu mnuBackColor 
            Caption         =   "Hintergrund-Farbe"
         End
         Begin VB.Menu mnuForeColor 
            Caption         =   "Vordergrund-Farbe"
         End
         Begin VB.Menu mnuFonts 
            Caption         =   "Schriften"
         End
         Begin VB.Menu mnuMarkColor 
            Caption         =   "Markierer"
         End
         Begin VB.Menu mnuLine3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRestoreSettings 
            Caption         =   "Lineal-Einstellungen zurücksetzen"
         End
      End
      Begin VB.Menu mnuTransparency 
         Caption         =   "&Transparenz"
         Begin VB.Menu mnuTransparencyRuler 
            Caption         =   "0%"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuTransparencyRuler 
            Caption         =   "25%"
            Index           =   1
         End
         Begin VB.Menu mnuTransparencyRuler 
            Caption         =   "50%"
            Index           =   2
         End
         Begin VB.Menu mnuTransparencyRuler 
            Caption         =   "75%"
            Index           =   3
         End
      End
      Begin VB.Menu mnuScale 
         Caption         =   "Einheit"
         Begin VB.Menu mnuScaleMode 
            Caption         =   "Pixel"
            Index           =   0
         End
         Begin VB.Menu mnuScaleMode 
            Caption         =   "Twips"
            Index           =   1
         End
         Begin VB.Menu mnuScaleMode 
            Caption         =   "Selbstdefiniert"
            Index           =   2
         End
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpR 
         Caption         =   "&Hilfe"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu MScreenShot 
      Caption         =   "MScreenShot"
      Begin VB.Menu mnuRuler 
         Caption         =   "&Pixel-Lineal"
      End
      Begin VB.Menu mnuSMagGlass 
         Caption         =   "&Lupe"
      End
      Begin VB.Menu mnuShowSize 
         Caption         =   "Anzeige der Fenstergröße"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFesteGroesse 
         Caption         =   "Feste Fenstergröße"
         Begin VB.Menu mnuFix 
            Caption         =   "16 × 16"
            Index           =   0
         End
         Begin VB.Menu mnuFix 
            Caption         =   "24 × 24"
            Index           =   1
         End
         Begin VB.Menu mnuFix 
            Caption         =   "32 × 32"
            Index           =   2
         End
         Begin VB.Menu mnuFix 
            Caption         =   "48 × 48"
            Index           =   3
         End
         Begin VB.Menu mnuFix 
            Caption         =   "64 × 64"
            Index           =   4
         End
         Begin VB.Menu mnuFix 
            Caption         =   "128 × 128"
            Index           =   5
         End
         Begin VB.Menu mnuFix 
            Caption         =   "256 × 256"
            Index           =   6
         End
         Begin VB.Menu mnuFix 
            Caption         =   "Selbstdefiniert"
            Index           =   7
         End
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpS 
         Caption         =   "&Hilfe"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Schließen      Esc"
      End
   End
   Begin VB.Menu MBorderStyle 
      Caption         =   "MBorderStyle"
      Begin VB.Menu mnuBorder 
         Caption         =   "Rahmen"
         Index           =   0
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss oben"
         Index           =   2
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss rechts"
         Index           =   3
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss unten"
         Index           =   4
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss links"
         Index           =   5
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss Mitte"
         Index           =   6
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Schatten"
         Index           =   7
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss oben-rechts"
         Index           =   9
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss unten-rechts"
         Index           =   10
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss unten-links"
         Index           =   11
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "Abriss oben-links"
         Index           =   12
      End
   End
   Begin VB.Menu MFile 
      Caption         =   "MFile"
      Begin VB.Menu mnuFile 
         Caption         =   "Öffnen"
         Index           =   0
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   ""
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mruFile 
            Caption         =   "-"
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileOpen 
            Caption         =   "Bild laden..."
         End
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Speichern"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Einfügen aus Datei"
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   ""
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mruPaste 
            Caption         =   "-"
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFilePaste 
            Caption         =   "Bild laden..."
         End
      End
      Begin VB.Menu MFileL1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpF 
         Caption         =   "Hilfe"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuUpdates 
         Caption         =   "Nach Online-Updates suchen"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Einstellungen zurücksetzen"
      End
      Begin VB.Menu mnuInternet 
         Caption         =   "Pixel-Lineal im Web"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
      End
   End
   Begin VB.Menu MArrow 
      Caption         =   "MArrow"
      Begin VB.Menu mnuArrow 
         Caption         =   "AL1"
         Index           =   0
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AL2"
         Index           =   1
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AL3"
         Index           =   2
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "MP1"
         Index           =   4
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AT1"
         Index           =   5
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AT2"
         Index           =   6
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AT3"
         Index           =   7
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "MP2"
         Index           =   9
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AR1"
         Index           =   10
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AR2"
         Index           =   11
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AR3"
         Index           =   12
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "MP3"
         Index           =   14
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AB1"
         Index           =   15
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AB2"
         Index           =   16
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "AB3"
         Index           =   17
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuArrow 
         Caption         =   "MP4"
         Index           =   19
      End
   End
   Begin VB.Menu MPal 
      Caption         =   "MPal"
      Begin VB.Menu mnuPal 
         Caption         =   "Standard"
         Index           =   0
      End
      Begin VB.Menu mnuPal 
         Caption         =   "Schwarz-Weiß"
         Index           =   1
      End
      Begin VB.Menu mnuPal 
         Caption         =   "Rot"
         Index           =   2
      End
      Begin VB.Menu mnuPal 
         Caption         =   "Grün"
         Index           =   3
      End
      Begin VB.Menu mnuPal 
         Caption         =   "Blau"
         Index           =   4
      End
      Begin VB.Menu mnuPal 
         Caption         =   "Gelb"
         Index           =   5
      End
      Begin VB.Menu mnuPal 
         Caption         =   "Aquamarin"
         Index           =   6
      End
      Begin VB.Menu mnuPal 
         Caption         =   "Violett"
         Index           =   7
      End
      Begin VB.Menu mnuPal 
         Caption         =   "Gesammelt"
         Index           =   8
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPopupMenuName As String
Private mPopupMenuCaption As String
Private mPopupMenuIndex As Integer
Private mPopupMenuChecked As Boolean
Private mPalette(8, 15) As Long

Public Function GetPalColors(palIndex As Integer) As Long()
Dim i As Integer
Dim palColors(15) As Long
    If palIndex = 8 Then
        For i = 0 To UBound(modMain.ColorCollection)
            palColors(i) = modMain.ColorCollection(i)
        Next i
    Else
        For i = 0 To 15
            palColors(i) = mPalette(palIndex, i)
        Next
    End If
    GetPalColors = palColors
End Function

Public Function GetPopupMenu(ByRef Sender As Form, ByVal X As Long, ByVal Y As Long, ByRef Name As String, ByRef Caption As String, ByRef Index As Integer, ByRef Checked As Boolean) As Boolean
    mPopupMenuName = ""
    mPopupMenuCaption = ""
    mPopupMenuIndex = -1
    mPopupMenuChecked = False

    Select Case Name
        Case "File":    Sender.PopupMenu MFile, vbPopupMenuLeftAlign, X, Y
        Case "Border":  Sender.PopupMenu MBorderStyle, vbPopupMenuLeftAlign, X, Y
        Case "Arrow":   Sender.PopupMenu MArrow, vbPopupMenuLeftAlign, X, Y
        Case "Palette": Sender.PopupMenu MPal, vbPopupMenuLeftAlign, X, Y
    End Select
    If Len(mPopupMenuName) > 0 Then
        Name = mPopupMenuName
        Caption = mPopupMenuCaption
        Index = mPopupMenuIndex
        Checked = mPopupMenuChecked
        GetPopupMenu = True
    End If
End Function

Public Sub ToogleMagGlass()
    On Error Resume Next
    If MagGlass Is Nothing Then
        Set MagGlass = New frmMagGlass
        With MagGlass
          .Show
          .WindowState = vbNormal
        End With
    Else
        Unload MagGlass
        Set MagGlass = Nothing
    End If
End Sub


Public Sub UpdateFileMru(FileName As String)
Dim i As Integer, j As Integer
    If mruFile(0).Caption <> FileName Then
        For i = 1 To mruFile.UBound - 1
            If FileName = mruFile(i).Caption Then
                For j = i + 1 To mruFile.UBound - 1
                    mruFile(j - 1).Caption = mruFile(j).Caption
                Next j
                mruFile(j - 1).Caption = ""
            End If
        Next i
        For i = mruFile.UBound - 1 To 0 Step -1
            If i > 0 Then
                mruFile(i).Caption = mruFile(i - 1).Caption
            Else
                mruFile(0).Caption = FileName
            End If
            SaveSetting App.Title, "Editor", "LastFile" & (i + 1), mruFile(i).Caption
            mruFile(i).Visible = (Len(mruFile(i).Caption) > 0)
        Next i
    End If
End Sub
Public Sub UpdatePasteMru(FileName As String)
Dim i As Integer, j As Integer
    If mruPaste(0).Caption <> FileName Then
        For i = 1 To mruPaste.UBound - 1
            If FileName = mruPaste(i).Caption Then
                For j = i + 1 To mruPaste.UBound - 1
                    mruPaste(j - 1).Caption = mruPaste(j).Caption
                Next j
                mruPaste(j - 1).Caption = ""
            End If
        Next i
        For i = mruPaste.UBound - 1 To 0 Step -1
            If i > 0 Then
                mruPaste(i).Caption = mruPaste(i - 1).Caption
            Else
                mruPaste(0).Caption = FileName
            End If
            SaveSetting App.Title, "Editor", "LastPaste" & (i + 1), mruPaste(i).Caption
            mruPaste(i).Visible = (Len(mruPaste(i).Caption) > 0)
        Next i
    End If

End Sub

Public Sub mnuMarker_Click()
    On Error GoTo mnuMarker_Click_Error
    If mnuMarker.Tag = "+" Then frmRuler.SetMarker Else frmRuler.RemoveMarker CInt(mnuMarker.Tag)
    mnuMarker.Tag = ""
Exit Sub

mnuMarker_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuMarker_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub


Public Sub mnuScaleMode_Click(Index As Integer)
    On Error GoTo mnuScaleMode_Click_Error
    mnuScaleMode(PL_USER).Checked = Index = PL_USER
    mnuScaleMode(PL_TWIPS).Checked = Index = PL_TWIPS
    mnuScaleMode(PL_PIXEL).Checked = Index = PL_PIXEL
    RulerScaleMode = Index
    Select Case RulerScaleMode
    Case PL_PIXEL
      RulerScaleMulti = 1
      XYFieldWidth = XYFieldMinWidth
    Case PL_TWIPS
      If frmRuler.Orientation = PL_HORIZONTAL Then
        RulerScaleMulti = LTwipsPerPixelX
      Else
        RulerScaleMulti = LTwipsPerPixelY
      End If
      XYFieldWidth = XYFieldMinWidth + Len(CStr(Round(RulerScaleMulti * 1000))) * 8
    Case PL_USER
      RulerScaleMulti = -1
    End Select
    frmRuler.ProcRefreshRuler frmRuler.Left, frmRuler.Top
    SaveSetting App.Title, "Options", "ScaleMode", RulerScaleMode
Exit Sub

mnuScaleMode_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuScaleMode_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Public Sub mnuTransparencyRuler_Click(Index As Integer)
Dim i As Integer
    On Error GoTo mnuTransparencyRuler_Click_Error
    For i = 0 To 3
      mnuTransparencyRuler(i).Checked = False
    Next i
    mnuTransparencyRuler(Index).Checked = True
    Select Case Index
      Case 0
          Call TransparencyRuler(frmRuler.hwnd, 255&)
      Case 1
          Call TransparencyRuler(frmRuler.hwnd, 192&)
      Case 2
          Call TransparencyRuler(frmRuler.hwnd, 128&)
      Case 3
          Call TransparencyRuler(frmRuler.hwnd, 64&)
    End Select
    frmRuler.Refresh
    SaveSetting App.Title, "Options", "TransparencyRuler", Index
    Exit Sub

mnuTransparencyRuler_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuTransparencyRuler_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub


Private Sub FillPalettes()
Dim i As Integer, j As Integer
Dim r As Byte, s As Byte
    mPalette(0, 0) = &H0&
    mPalette(0, 1) = &H80&
    mPalette(0, 2) = &H8000&
    mPalette(0, 3) = &H8080&
    mPalette(0, 4) = &H800000
    mPalette(0, 5) = &H800080
    mPalette(0, 6) = &H808000
    mPalette(0, 7) = &H808080
    mPalette(0, 8) = &HC0C0C0
    mPalette(0, 9) = &HFF&
    mPalette(0, 10) = &HFF00&
    mPalette(0, 11) = &HFFFF&
    mPalette(0, 12) = &HFF0000
    mPalette(0, 13) = &HFF00FF
    mPalette(0, 14) = &HFFFF00
    mPalette(0, 15) = &HFFFFFF
    For i = 1 To 7
        mPalette(i, 15) = &HFFFFFF
    Next i
    For i = 1 To 14
        r = i * 17
        mPalette(1, i) = RGB(r, r, r)
        If i <= 7 Then
            r = 45 + (i * 30)
            mPalette(2, i) = RGB(r, 0, 0)
            mPalette(3, i) = RGB(0, r, 0)
            mPalette(4, i) = RGB(0, 0, r)
            mPalette(5, i) = RGB(r, r, 0)
            mPalette(6, i) = RGB(0, r, r)
            mPalette(7, i) = RGB(r, 0, r)
        Else
            r = (i - 7) * 31.8
            mPalette(2, i) = RGB(255, r, r)
            mPalette(3, i) = RGB(r, 255, r)
            mPalette(4, i) = RGB(r, r, 255)
            mPalette(5, i) = RGB(255, 255, r)
            mPalette(6, i) = RGB(r, 255, 255)
            mPalette(7, i) = RGB(255, r, 255)
        End If
    Next
    
    For j = 0 To 7 'UBound(mPalette, 1)
        For i = 0 To 15
            picMenuPal(j).Line (i * 10, 0)-((i + 1) * 10, 18), mPalette(j, i), BF
        Next i
    Next j
    
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim h1 As Long, h2 As Long, h3 As Long

    On Error Resume Next
    h1 = Me.hwnd
    For i = 0 To 15
      Load mnuColorCollectionItems(i)
      mnuColorCollectionItems(i).Visible = False
      Load picMenuColor(i)
      If i <= 8 Then
        Load picMenuPal(i)
        If App.LogMode = 0 Then
            picMenuPal(i).Visible = True
            picMenuPal(i).Top = picMenuPal(0).Top + ((i + 1) * picMenuPal(0).Height)
        End If
        picMenuRuler(i).Picture = picMenuRuler(i).Image
      End If
      picMenuFile(i).Picture = picMenuFile(i).Image
    Next i
    Err.Clear
    
    On Error GoTo Form_Load_Error
    ColorCode = CInt(GetSetting(App.Title, "Options", "ColorCode", 0))
    Call mnuColorCode_Click(CInt(ColorCode))
    modMenuColor.Set_MenuColor nfoMenuColor, h1, &HF2F2F2, 0, True  'MRuler
    modMenuColor.Set_MenuColor nfoMenuColor, h1, &HF2F2F2, 1, True  'MScreenShot
    modMenuColor.Set_MenuColor nfoMenuColor, h1, &HF2F2F2, 2, True  'BorderStyle
    modMenuColor.Set_MenuColor nfoMenuColor, h1, &HF2F2F2, 3, True  'MFile
    modMenuColor.Set_MenuColor nfoMenuColor, h1, &HF2F2F2, 4, True  'MArrow
    modMenuColor.Set_MenuColor nfoMenuColor, h1, &HF2F2F2, 5, True  'MPal
    
    For i = mruFile.LBound To mruFile.UBound - 1
        mruFile(i).Caption = Trim$(GetSetting(App.Title, "Editor", "LastFile" & (i + 1), ""))
        mruFile(i).Visible = (Len(mruFile(i).Caption) > 0)
    Next i
    mruFile(i).Visible = mruFile(mruFile.LBound).Visible
    For i = mruPaste.LBound To mruPaste.UBound - 1
        mruPaste(i).Caption = Trim$(GetSetting(App.Title, "Editor", "LastPaste" & (i + 1), ""))
        mruPaste(i).Visible = (Len(mruPaste(i).Caption) > 0)
    Next i
    mruPaste(i).Visible = mruPaste(mruPaste.LBound).Visible
    
    h1 = GetMenu(h1)            'Hauptmenü
    h2 = GetSubMenu(h1, 0&)     'MRuler
        SetMenuItemBitmaps h2, 0, MF_BYPOSITION, picMenuRuler(1).Picture, picMenuRuler(1).Picture
        SetMenuItemBitmaps h2, 1, MF_BYPOSITION, picMenuRuler(2).Picture, picMenuRuler(2).Picture
        SetMenuItemBitmaps h2, 2, MF_BYPOSITION, picMenuRuler(3).Picture, picMenuRuler(3).Picture
        SetMenuItemBitmaps h2, 3, MF_BYPOSITION, picMenuRuler(4).Picture, picMenuRuler(4).Picture
        SetMenuItemBitmaps h2, 13, MF_BYPOSITION, picMenuFile(4).Picture, picMenuFile(4).Picture
        mnuColors_Click
        
    h2 = GetSubMenu(h1, 1&)     'MScreenShot
        SetMenuItemBitmaps h2, 0, MF_BYPOSITION, picMenuRuler(0).Picture, picMenuRuler(0).Picture
        SetMenuItemBitmaps h2, 1, MF_BYPOSITION, picMenuRuler(2).Picture, picMenuRuler(2).Picture
        SetMenuItemBitmaps h2, 5, MF_BYPOSITION, picMenuFile(4).Picture, picMenuFile(4).Picture
        
    h2 = GetSubMenu(h1, 2&)     'MBorderStyle
    For i = 0 To mnuBorder.UBound
        If mnuBorder(i).Caption <> "-" Then
            h3 = GetMenuItemID(h2, i) 'mnuBorder(i)
            picBorder(i).CurrentX = 20: picBorder(i).CurrentY = 1
            picBorder(i).Print mnuBorder(i).Caption
            picBorder(i).Picture = picBorder(i).Image
            If i = 7 Then   'Umbruch
                Call ModifyMenu(h2, i, MF_BYPOSITION Or MF_MENUBREAK Or MF_BITMAP, h3, picBorder(i).Picture.Handle)
            Else
                Call ModifyMenu(h2, h3, MF_BITMAP, h3, picBorder(i).Picture.Handle)
            End If
        End If
    Next i
    
    h2 = GetSubMenu(h1, 3&)     'MFile
    For i = 0 To picMenuFile.UBound
        If i <> 3 Then
            SetMenuItemBitmaps h2, i, MF_BYPOSITION, picMenuFile(i).Picture, picMenuFile(i).Picture
        End If
    Next

        
    h2 = GetSubMenu(h1, 4&)     'MArrow
    For i = 0 To mnuArrow.UBound
        h3 = GetMenuItemID(h2, i) 'mnuArrow(i)
        If mnuArrow(i).Caption <> "-" Then
            picMenuArrow(i).Picture = picMenuArrow(i).Image
            If i = 5 Or i = 10 Or i = 15 Then   'Umbruch
                Call ModifyMenu(h2, i, MF_BYPOSITION Or MF_MENUBREAK Or MF_BITMAP, h3, picMenuArrow(i).Picture.Handle)
            Else
                Call ModifyMenu(h2, h3, MF_BITMAP, h3, picMenuArrow(i).Picture.Handle)
            End If
        End If
    Next i
    
    Call FillPalettes
    h2 = GetSubMenu(h1, 5&)     'MPal
    For i = 0 To mnuPal.UBound
        picMenuPal(i).Picture = picMenuPal(i).Image
        SetMenuItemBitmaps h2, i, MF_BYPOSITION, picMenuPal(i).Picture, picMenuPal(i).Picture
    Next i
        
    'Abriss-Bilder
    Load picTearOff(1): Load picTearOff(2): Load picTearOff(3)
    
    picTearOff(1).Width = picTearOff(0).Height: picTearOff(1).Height = picTearOff(0).Width
    picTearOff(1).Picture = gdiplus.CopyStdPicture(picTearOff(0).Image, 3&)  'rechts
    
    picTearOff(2).Width = picTearOff(0).Width: picTearOff(2).Height = picTearOff(0).Height
    picTearOff(2).Picture = gdiplus.CopyStdPicture(picTearOff(0).Image, 2&)   'oben
    
    picTearOff(3).Width = picTearOff(0).Height: picTearOff(3).Height = picTearOff(0).Width
    picTearOff(3).Picture = gdiplus.CopyStdPicture(picTearOff(0).Image, 1&)   'links
    
    mnuShowSize.Checked = CBool(Val(GetSetting(App.Title, "Options", "ShowSize", 1)))
    
    Call modMain.CheckVersion(autoCheck:=True)
Exit Sub

Form_Load_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.Form_Load." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    Set frmMenu = Nothing
End Sub

Private Sub MRuler_Click()
    If Capture Is Nothing And isImgEditor = False Then mnuEnd.Caption = "Beenden" Else mnuEnd.Caption = "Schließen"
End Sub

Private Sub MScreenShot_Click()
    If frmRuler.Visible Or isImgEditor Then mnuClose.Caption = "Schließen" Else mnuClose.Caption = "Beenden"
End Sub

Private Function isImgEditor() As Boolean
Dim f As Form
    For Each f In Forms
        If TypeOf f Is frmImage Then
            isImgEditor = True
            Exit For
        End If
    Next
End Function



Private Sub mnuArrow_Click(Index As Integer)
    mPopupMenuName = "mnuArrow"
    mPopupMenuCaption = mnuArrow(Index).Caption
    mPopupMenuIndex = Index
    mPopupMenuChecked = True
    mnuArrow(Index).Checked = True
End Sub

Private Sub mnuBackColor_Click()
On Error GoTo errBackColor
  Dim col As Long
  col = frmRuler.picRuler.BackColor
  col = ShowColorDlg(frmRuler.hwnd, col)
  If col > -1 Then
    frmRuler.picRuler.BackColor = col
    frmRuler.ProcRefreshRuler frmRuler.Left, frmRuler.Top
    SaveSetting App.Title, "Options", "BackColor", col
  End If
  Exit Sub
errBackColor:
MsgBox "Fehler: " & Err.Number & vbCrLf & Err.Description

End Sub

Private Sub mnuBorder_Click(Index As Integer)
    mPopupMenuName = "mnuBorder"
    mPopupMenuCaption = mnuBorder(Index).Caption
    mPopupMenuIndex = Index
    mPopupMenuChecked = mnuBorder(Index).Checked
End Sub

Private Sub mnuClose_Click()
    On Error Resume Next
    Unload Capture
    Set Capture = Nothing
End Sub

Private Sub mnuColorCode_Click(Index As Integer)
    If Index = 0 Then
        mnuColorCode(0).Checked = True
        mnuColorCode(1).Checked = False
        mnuColorCode(2).Checked = False
        If Not MagGlass Is Nothing Then
            With MagGlass
                .mnuColorCode(0).Checked = True
                .mnuColorCode(1).Checked = False
                .mnuColorCode(2).Checked = False
            End With
        End If
    ElseIf Index = 1 Then
        mnuColorCode(0).Checked = False
        mnuColorCode(1).Checked = True
        mnuColorCode(2).Checked = False
        If Not MagGlass Is Nothing Then
            With MagGlass
                .mnuColorCode(0).Checked = False
                .mnuColorCode(1).Checked = True
                .mnuColorCode(2).Checked = False
            End With
        End If
    Else
        mnuColorCode(0).Checked = False
        mnuColorCode(1).Checked = False
        mnuColorCode(2).Checked = True
        If Not MagGlass Is Nothing Then
            With MagGlass
                .mnuColorCode(0).Checked = False
                .mnuColorCode(1).Checked = False
                .mnuColorCode(2).Checked = True
            End With
        End If
    End If
    ColorCode = Index
    On Error Resume Next
    SaveSetting App.Title, "Options", "ColorCode", ColorCode

End Sub

Private Sub mnuColorCollectionItems_Click(Index As Integer)
  On Error GoTo mnuColorCollectionItems_Click_Error
  Clipboard.Clear
  Clipboard.SetText mnuColorCollectionItems(Index).Caption, vbCFText
Exit Sub

mnuColorCollectionItems_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuColorCollectionItems_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuColorCollection_Click()
    On Error GoTo mnuColorCollection_Click_Error
    Call modMain.FillMenuColorCollection(Me, 6&) '6 = Position von Menü in Ruller-Menü
Exit Sub

mnuColorCollection_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuColorCollection_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuColors_Click()
Dim mnuID As Long, h1 As Long, h2 As Long, h3 As Long
    On Error GoTo mnuColors_Click_Error
    h1 = GetMenu(Me.hwnd)   'Hauptmenü
    h2 = GetSubMenu(h1, 0)  'MRuler
    h3 = GetSubMenu(h2, 8)  'Lineal-Farben / mnuColorCollection

    picMenuColor(0).Line (0, 0)-(18, 18), frmRuler.picRuler.BackColor, BF
    picMenuColor(0).Picture = picMenuColor(0).Image
    SetMenuItemBitmaps h3, 0, MF_BYPOSITION, picMenuColor(0).Picture, picMenuColor(0).Picture

    picMenuColor(1).Line (0, 0)-(18, 18), frmRuler.picRuler.ForeColor, BF
    picMenuColor(1).Picture = picMenuColor(1).Image
    SetMenuItemBitmaps h3, 1, MF_BYPOSITION, picMenuColor(1).Picture, picMenuColor(1).Picture

    picMenuColor(2).Line (0, 0)-(18, 18), frmRuler.picRuler.ForeColor, BF
    picMenuColor(2).Picture = picMenuColor(2).Image
    SetMenuItemBitmaps h3, 2, MF_BYPOSITION, picMenuColor(2).Picture, picMenuColor(2).Picture

    picMenuColor(3).Line (0, 0)-(18, 18), frmRuler.MarkerColor, BF
    picMenuColor(3).Picture = picMenuColor(3).Image
    SetMenuItemBitmaps h3, 3, MF_BYPOSITION, picMenuColor(3).Picture, picMenuColor(3).Picture
Exit Sub

mnuColors_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuColors_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuEnd_Click()
    On Error Resume Next
    If mnuEnd.Caption = "Schließen" Then frmRuler.Hide Else Unload frmRuler
End Sub

Private Sub mnuFileOpen_Click()
    mPopupMenuName = "mnuFileOpen"
    mPopupMenuCaption = mnuFileOpen.Caption
    mPopupMenuIndex = -1
    mPopupMenuChecked = mnuFileOpen.Checked
End Sub

Private Sub mnuFilePaste_Click()
    mPopupMenuName = "mnuFilePaste"
    mPopupMenuCaption = mnuFilePaste.Caption
    mPopupMenuIndex = -1
    mPopupMenuChecked = mnuFilePaste.Checked
End Sub

Private Sub mnuFileSave_Click()
    mPopupMenuName = "mnuFileSave"
    mPopupMenuCaption = mnuFileSave.Caption
    mPopupMenuIndex = -1
    mPopupMenuChecked = mnuFileSave.Checked
End Sub

Private Sub mnuFix_Click(Index As Integer)
Dim ret As String
Dim X As Long, Y As Long
Dim i As Integer
    On Error GoTo mnuFix_Click_Error
    With Capture
    Select Case Index
        Case 0: .Move .Left, .Top, 16 * LTwipsPerPixelX, 16 * LTwipsPerPixelY
        Case 1: .Move .Left, .Top, 24 * LTwipsPerPixelX, 24 * LTwipsPerPixelY
        Case 2: .Move .Left, .Top, 32 * LTwipsPerPixelX, 32 * LTwipsPerPixelY
        Case 3: .Move .Left, .Top, 48 * LTwipsPerPixelX, 48 * LTwipsPerPixelY
        Case 4: .Move .Left, .Top, 64 * LTwipsPerPixelX, 64 * LTwipsPerPixelY
        Case 5: .Move .Left, .Top, 128 * LTwipsPerPixelX, 128 * LTwipsPerPixelY
        Case 6: .Move .Left, .Top, 256 * LTwipsPerPixelX, 256 * LTwipsPerPixelY
        Case Else
            X = Capture.Width \ LTwipsPerPixelX
            Y = Capture.Height \ LTwipsPerPixelY
            ret = Trim$(InputBox("Gewünschte Abmessungen im Format x,y eingeben." & vbCrLf & "Für ein Quadrat reicht die Eingabe einer einzelnen Zahl.", "Selbstdefinierter Screenshot", X & "," & Y))
            If Len(ret) = 0 Then Exit Sub
            If IsNumeric(Left$(ret, 1)) Then
                For i = 1 To Len(ret)
                    If Not IsNumeric(Mid$(ret, i, 1)) Then Exit For
                Next
                If i > Len(ret) And IsNumeric(ret) Then
                    X = CInt(Val(ret))
                    Y = X
                Else
                    X = CInt(Val(Left$(ret, i - 1)))
                    Y = CInt(Val(Mid$(ret, i + 1)))
                End If
                If X < 12 Then X = 12
                If Y < 12 Then Y = 12
                .Move .Left, .Top, X * LTwipsPerPixelX, Y * LTwipsPerPixelY
            Else
                MsgBox ret & vbCrLf & "Wirklich?", vbExclamation, ":-("
            End If
        End Select
        End With
    Exit Sub
    
mnuFix_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMenu.mnuFix_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub mnuFonts_Click()
Dim TmpFName As String
  Dim LFnt As LOGFONT
  Dim CF_T As CHOOSEFONT_TYPE
  On Error GoTo mnuFonts_Click_Error
  ' Dialog-Eigenschaften setzen
  
      With CF_T
      .nSizeMax = 12
      .nSizeMin = 4
      .Flags = CF_SCREENFONTS Or CF_FORCEFONTEXIST Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE Or CF_NOSCRIPTSEL
      .hWndOwner = frmRuler.hwnd
      .lStructSize = Len(CF_T)
      .lpLogFont = VarPtr(LFnt)
      .hInstance = App.hInstance
      .hDC = 0
      .nFontType = SCREEN_FONTTYPE
      .rgbColors = Convert_OLEtoRBG(frmRuler.picRuler.ForeColor)
    End With
    
    TmpFName = frmRuler.picRuler.FontName
    TmpFName = StrConv(TmpFName, vbFromUnicode)
    LFnt.lfFaceName = TmpFName & vbNullChar
    With LFnt
        .lfHeight = frmRuler.picRuler.FontSize * -20 / LTwipsPerPixelY 'Alternativ: 'MM_TEXT mapping mode: lfHeight = -MulDiv(PointSize, GetDeviceCaps(hDC, LOGPIXELSY), 72);
        .lfWeight = IIf(frmRuler.picRuler.FontBold, FW_BOLD, FW_NORMAL)
        .lfItalic = Abs(frmRuler.picRuler.FontItalic)
        .lfUnderline = Abs(frmRuler.picRuler.FontUnderline)
        .lfStrikeOut = Abs(frmRuler.picRuler.FontStrikethru)
        .lfOutPrecision = OUT_TT_PRECIS
        .lfQuality = ANTIALIASED_QUALITY
        .lfCharSet = DEFAULT_CHARSET
        .lfPitchAndFamily = VARIABLE_PITCH
    End With
 
  ' Dialog aufrufen
  If ChooseFont(CF_T) = 0 Then Exit Sub
  TmpFName = StrConv(LFnt.lfFaceName, vbUnicode)

  With frmRuler.picRuler
    With .Font
        .Name = Left$(TmpFName, InStr(1, TmpFName, vbNullChar) - 1)
        .bold = CBool(LFnt.lfWeight >= FW_BOLD)
        .Italic = CBool(LFnt.lfItalic)
        .Underline = CBool(LFnt.lfUnderline)
        .Strikethrough = CBool(LFnt.lfStrikeOut)
        .Size = CF_T.iPointSize / 10
        SaveSetting App.Title, "Options", "FontBold", Abs(.bold)
        SaveSetting App.Title, "Options", "FontItalic", Abs(.Italic)
        SaveSetting App.Title, "Options", "FontUnderline", Abs(.Underline)
        SaveSetting App.Title, "Options", "FontStrikethru", Abs(.Strikethrough)
        SaveSetting App.Title, "Options", "FontSize", .Size
        SaveSetting App.Title, "Options", "FontName", .Name
    End With
    .ForeColor = CF_T.rgbColors
    SaveSetting App.Title, "Options", "ForeColor", .ForeColor
  End With
  frmRuler.ProcRefreshRuler frmRuler.Left, frmRuler.Top

Exit Sub

mnuFonts_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuFonts_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuForeColor_Click()
On Error GoTo errForeColor
  Dim col As Long
  col = frmRuler.picRuler.ForeColor
  col = ShowColorDlg(frmRuler.hwnd, col)
  If col > -1 Then
    frmRuler.picRuler.ForeColor = col
    frmRuler.ProcRefreshRuler frmRuler.Left, frmRuler.Top
    SaveSetting App.Title, "Options", "ForeColor", col
  End If
  Exit Sub
errForeColor:
MsgBox "Fehler: " & Err.Number & vbCrLf & Err.Description
End Sub



Private Sub mnuHelpF_Click()
    On Error GoTo mnuHelpF_Click_Error
    ShellExec "https://docs.ww-a.de/doku.php/pixellineal:bildeditor", vbNormalFocus
    Exit Sub
    
mnuHelpF_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMenu.mnuHelpF_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub mnuHelpR_Click()
    On Error Resume Next
    ShellExec "https://docs.ww-a.de/doku.php/pixellineal:start", vbNormalFocus
End Sub

Private Sub mnuHelpS_Click()
    On Error Resume Next
    ShellExec "https://docs.ww-a.de/doku.php/pixellineal:screenshot", vbNormalFocus
End Sub


Private Sub mnuInfo_Click()
  MsgBox modMain.GetInfo, vbInformation, "Pixel-Lineal"
End Sub

Private Sub mnuInternet_Click()
    On Error GoTo mnuInternet_Click_Error
    ShellExec "https://www.ww-a.de/pixellineal.html", vbNormalFocus
    Exit Sub
    
mnuInternet_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMenu.mnuInternet_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub mnuMagColor_Click()
Dim isMagGlass As Boolean
    On Error GoTo mnuMagColor_Click_Error
    If Not MagGlass Is Nothing Then
        isMagGlass = True
        MagGlass.Visible = False
        Set MagGlass = Nothing
    End If
    Set MagColor = New frmMagColor
    MagColor.Show vbModal, Me
    Unload MagColor
    If MagColor.PipColor <> &H1000000 Then
        If MagColor.PipColor > 0 Then CopyRGB MagColor.PipColor
    End If
    Set MagColor = Nothing
    

mnuMagColor_Click_Resume:
    On Error Resume Next
    If isMagGlass Then
        Set MagGlass = frmMagGlass
        MagGlass.Visible = True
    End If
    Exit Sub
    
mnuMagColor_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMenu.mnuMagColor_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
     Resume mnuMagColor_Click_Resume
End Sub

Private Sub mnuMarkColor_Click()
On Error GoTo errMarkColor
  Dim col As Long
  col = frmRuler.MarkerColor
  col = ShowColorDlg(frmRuler.hwnd, col)
  If col > -1 Then
    frmRuler.MarkerColor = col
    frmRuler.ProcRefreshRuler frmRuler.Left, frmRuler.Top
    SaveSetting App.Title, "Options", "MarkColor", col
  End If
  Exit Sub
errMarkColor:
MsgBox "Fehler: " & Err.Number & vbCrLf & Err.Description
End Sub

Private Sub mnuOrientation_Click()
    On Error GoTo mnuOrientation_Click_Error
    frmRuler.Orientation = Abs(frmRuler.Orientation - 1)
Exit Sub

mnuOrientation_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuOrientation_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuPal_Click(Index As Integer)
Dim i As Integer
    mPopupMenuName = "mnuPal"
    mPopupMenuIndex = Index
    mPopupMenuChecked = mnuBorder(Index).Checked
End Sub


Private Sub mnuRMagGlass_Click()
    On Error GoTo mnuRMagGlass_Click_Error
    ToogleMagGlass
    Exit Sub
    
mnuRMagGlass_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMenu.mnuRMagGlass_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub

Private Sub mnuReset_Click()
    mPopupMenuName = "mnuReset"
    mPopupMenuIndex = -1
End Sub



Private Sub mnuRestoreSettings_Click()
    On Error Resume Next
    With frmRuler
        .MarkerColor = RGB(255, 0, 0)
        With .picRuler
            With .Font
                .Name = "Arial"
                .bold = False
                .Italic = False
                .Underline = False
                .Strikethrough = False
                .Size = 6
            End With
            .ForeColor = RGB(132, 132, 132)
            .BackColor = RGB(255, 255, 231)
        End With
        .ProcRefreshRuler .Left, .Top
    End With
    DeleteSetting App.Title, "Options", "FontBold"
    DeleteSetting App.Title, "Options", "FontItalic"
    DeleteSetting App.Title, "Options", "FontName"
    DeleteSetting App.Title, "Options", "FontSize"
    DeleteSetting App.Title, "Options", "FontStrikethru"
    DeleteSetting App.Title, "Options", "FontUnderline"
    DeleteSetting App.Title, "Options", "BackColor"
    DeleteSetting App.Title, "Options", "ForeColor"
    DeleteSetting App.Title, "Options", "MarkColor"

End Sub

Private Sub mnuRuler_Click()
    On Error GoTo mnuRuler_Click_Error
    frmRuler.Show
    Exit Sub
    
mnuRuler_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMenu.mnuRuler_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub
Private Sub mnuSMagGlass_Click()
    On Error GoTo mnuSMagGlass_Click_Error
    ToogleMagGlass
    Exit Sub
    
mnuSMagGlass_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Fehler: " & Err.Number & vbCrLf & _
     "Beschreibung: " & Err.Description & vbCrLf & _
     "Quelle: frmMenu.mnuSMagGlass_Click." & Erl & vbCrLf & Err.Source, _
     vbCritical
End Sub


Private Sub mnuScreenShot_Click()
    On Error GoTo mnuScreenShot_Click_Error
    If Capture Is Nothing Then
        Set Capture = New frmCapture
        Capture.Show vbModeless, Me
    Else
        Unload Capture
        Set Capture = Nothing
    End If
Exit Sub

mnuScreenShot_Click_Error:
Screen.MousePointer = vbDefault
MsgBox "Fehler: " & Err.Number & vbCrLf & _
 "Beschreibung: " & Err.Description & vbCrLf & _
 "Quelle: frmMenu.mnuScreenShot_Click." & Erl & vbCrLf & Err.Source, _
 vbCritical
End Sub

Private Sub mnuShowSize_Click()
    mnuShowSize.Checked = Not mnuShowSize.Checked
    Capture.lblSize.Visible = mnuShowSize.Checked
    SaveSetting App.Title, "Options", "ShowSize", Abs(mnuShowSize.Checked)
End Sub



Private Sub mnuUpdates_Click()
    Call modMain.CheckVersion
End Sub

Private Sub mruFile_Click(Index As Integer)
    mPopupMenuName = "mruFile"
    mPopupMenuCaption = mruFile(Index).Caption
    mPopupMenuIndex = Index
    mPopupMenuChecked = mruFile(Index).Checked
End Sub

Private Sub mruPaste_Click(Index As Integer)
    mPopupMenuName = "mruPaste"
    mPopupMenuCaption = mruPaste(Index).Caption
    mPopupMenuIndex = Index
    mPopupMenuChecked = mruPaste(Index).Checked
End Sub


