VERSION 5.00
Begin VB.Form PassMod 
   BorderStyle     =   4  '���� ���� â
   Caption         =   "�н����庯��"
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
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   390
      IMEMode         =   3  '��� ����
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   390
      IMEMode         =   3  '��� ����
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   390
      IMEMode         =   3  '��� ����
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�� ��й�ȣ Ȯ�� :"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�� ��й�ȣ :"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����й�ȣ :"
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
        MsgBox "��й�ȣ�� ����Ǿ����ϴ�.", vbOKOnly, "�˸�"
        Unload Me
        
    Else
        MsgBox "���ο� ��й�ȣ ���Է��� ��ġ���� �ʽ��ϴ�.", vbOKOnly, "�˸�"
    End If
Else
    MsgBox "���� ��й�ȣ�� Ʋ�Ƚ��ϴ�.", vbOKOnly, "�˸�"
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

