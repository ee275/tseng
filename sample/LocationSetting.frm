VERSION 5.00
Begin VB.Form LocationSetting 
   Caption         =   "�۾��� ����"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton del 
      Caption         =   "����"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton add 
      Caption         =   "�߰�"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "������ �۾���"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "�߰��� �۾���"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "LocationSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim apple As String

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Combo1.Text = ""
End Sub




Private Sub add_Click()
sql = "insert into location(�۾���) values('" & Text1.Text & "')"
If Not Run(sql) Then Exit Sub
MsgBox Text1.Text + "��(��)����Ǿ����ϴ�.", vbOKOnly, "�˸�"
 Call comboview
End Sub



Private Sub del_Click()
     
     apple = Combo1.Text
     If apple = "��ü" Then
        MsgBox "��ü ����� ����Ǽ� �����ϴ�", vbOKOnly, "�˸�"
     Else
        sql = "delete From location where �۾��� ='" & apple & "'"
        If Not Run(sql) Then Exit Sub
        MsgBox Combo1.Text + "��(��) ���� �Ǿ����ϴ�.", vbOKOnly, "�˸�"
        Call comboview
     End If
End Sub

Private Sub Form_Load()

    Call comboview

End Sub

Public Sub comboview()

 sql = "select * from location where �۾��� <> '��ü'"
    If Not Run(sql) Then Exit Sub
     
     If Not (Rs.EOF Or Rs.BOF) Then
         Rs.MoveFirst
        
         Combo1.Clear '�̰� �޺�1�� ��系���� ����� �κ��Ӵϴ�.
         Do While Not Rs.EOF
             
          With Combo1
              .AddItem Trim(Rs.Fields("�۾���"))
          End With
             Rs.MoveNext
         Loop
         Text1.Text = ""
    End If
    
End Sub
