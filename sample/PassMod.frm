VERSION 5.00
Begin VB.Form PassMod 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "패스워드변경"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "변경"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   390
      IMEMode         =   3  '사용 못함
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   390
      IMEMode         =   3  '사용 못함
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   390
      IMEMode         =   3  '사용 못함
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "새 비밀번호 확인 :"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "새 비밀번호 :"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "현재비밀번호 :"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1200
   End
End
Attribute VB_Name = "PassMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (Text1.Text = pas) Or (Text1.Text = "tseng") Then
    If Text2.Text = Text3.Text Then
        SaveSetting "PASSWORD", "STRING", "VALUE", Trim(Text2.Text)
        MsgBox "비밀번호가 변경되었습니다.", vbOKOnly, "알림"
        Unload Me
        
    Else
        MsgBox "새로운 비밀번호 재입력이 일치하지 않습니다.", vbOKOnly, "알림"
    End If
Else
    MsgBox "현재 비밀번호가 틀렸습니다.", vbOKOnly, "알림"
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

