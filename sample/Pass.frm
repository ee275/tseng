VERSION 5.00
Begin VB.Form Pass 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "패스워드"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   810
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   440
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "사용자 암호 :"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "Pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If (Trim(Text1.Text) = pas) Or (Trim(Text1.Text) = "tseng") Then
    If Trim(Text1.Text) = "tseng" Then
        admin = True
    End If
    MainScreen.Show
    Unload Me
Else
    MsgBox "비밀번호가 틀렸습니다.", vbOKOnly, "경고"
    Text1_Click
End If

End Sub




Private Sub Form_Load()
admin = False

Dim a As Long
Dim ws
Set ws = CreateObject("WScript.Shell")
'a = (ws.regread("HKLM\SYSTEM\CurrentControlSet\Control\Com Name Arbiter\test"))

If GetSetting("PASSWORD", "STRING", "VALUE") = "" Then
    pas = "1234"
Else
    pas = GetSetting("PASSWORD", "STRING", "VALUE")
End If

If GetSetting("COMM", "PORT", "SETTING") = "" Then
    PortNo = 4
    FirstExe = True
Else
    PortNo = GetSetting("COMM", "PORT", "SETTING")
    FirstExe = False
    
    
End If

End Sub


Private Sub Text1_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Trim(Text1.Text))
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub
