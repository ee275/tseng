VERSION 5.00
Begin VB.Form Auto_connect_end 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "장비 통신 안내"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   3960
      Top             =   2040
   End
   Begin VB.Label Label3 
      Caption         =   "데이타 전송 작업 완료후 장비 리셋중입니다.                                                               장치를 분리하지 마십시오"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Auto_connect_end"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim apple As Integer
Private Sub Timer1_Timer()
    If mFlag_timer = 1 Then
        
       
    ElseIf mFlag_timer = 2 Then
        commCmd = "X"
        MainScreen.MSComm1.Output = commCmd
        MainScreen.mRefresh
        Unload Me
    End If
   
    
    
End Sub
