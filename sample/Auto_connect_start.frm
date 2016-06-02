VERSION 5.00
Begin VB.Form Auto_connect_start 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "장치 통신 안내"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1644.449
   ScaleMode       =   0  '사용자
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   3840
      Top             =   1920
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Text            =   "Auto_connect_start.frx":0000
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "장비 연결중입니다.  잠시 기다려 주십시오         장비를 분리하지 마십시오."
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "---------------------------------------"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "장치 연결중 - - - - - - -"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "Auto_connect_start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Unload Me

End Sub
