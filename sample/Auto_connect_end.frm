VERSION 5.00
Begin VB.Form Auto_connect_end 
   BorderStyle     =   4  '���� ���� â
   Caption         =   "��� ��� �ȳ�"
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
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   3960
      Top             =   2040
   End
   Begin VB.Label Label3 
      Caption         =   "����Ÿ ���� �۾� �Ϸ��� ��� �������Դϴ�.                                                               ��ġ�� �и����� ���ʽÿ�"
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
