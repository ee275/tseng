VERSION 5.00
Begin VB.Form Auto_worker_add 
   Caption         =   "�۾��� ���"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command2 
      Caption         =   "������"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ϱ�"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox worker_id 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox worker_name 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label clover_serial 
      BorderStyle     =   1  '���� ����
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "�ֹε�Ϲ�ȣ        (��:8307231059475)"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "�۾��� ����"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "-------------------------------------------------"
      Height          =   135
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Ŭ�ι� �ø��� ��ȣ"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Auto_worker_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim dialog

dialog = MsgBox("������ ����� �Ϸ�Ǿ����ϴ�   Ȯ���� ������ ����Ÿ �б⸦ �����մϴ�    ���� ��ġ�� �и����� ���ʽÿ�", [vbOKCancel], [], [], [])

If dialog = vbOK Then
Auto_worker_enter.Show

End If

If dialog = vbCancel Then


End If



End Sub

Private Sub Command2_Click()
Unload Me
End Sub
