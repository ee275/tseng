VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form All_worker 
   BorderStyle     =   4  '���� ���� â
   Caption         =   "�۾��� ����"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   9360
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Fg1 
      Height          =   8895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   15690
      _Version        =   393216
      Cols            =   7
   End
End
Attribute VB_Name = "All_worker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim iRow1 As Integer
Dim mWorker_Location As String
Dim flag





Private Sub Combo1_Click()
Fg1.Clear
'mWorker_Location = Combo1.Text
flag = 1

Call DBCall

End Sub

Private Sub Command1_Click()
On Error Resume Next
Rs.MoveFirst
If Not Rs.EOF Then
    
    Data_search.Text1 = Fg1.TextMatrix(Fg1.RowSel, 1)
    Data_search.Text2 = Fg1.TextMatrix(Fg1.RowSel, 2)
    Data_search.Text3 = Fg1.TextMatrix(Fg1.RowSel, 3)
    Unload Me
Else
    MsgBox "��ȸ�� �ڷᰡ �����ϴ�.", vbOKOnly, "�˸�"
End If
End Sub
Public Sub Worker_Location(apple)
mWorker_Location = apple

End Sub
Private Sub Command2_Click()
Unload Me

End Sub
Public Sub comboview()


    
End Sub
Private Sub Fg1_DblClick()

Data_search.Show
    iRow1 = Fg1.Row

    Data_search.Text1 = Fg1.TextMatrix(iRow1, 1)
    Data_search.Text2 = Fg1.TextMatrix(iRow1, 2)
    Data_search.Text3 = Fg1.TextMatrix(iRow1, 3)
    Data_search.Combo1 = Fg1.TextMatrix(iRow1, 6)
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim i
    Fg1.ColWidth(0) = 900
    Fg1.ColWidth(1) = 1100
    Fg1.ColWidth(2) = 1600
    Fg1.ColWidth(3) = 2100
    Fg1.ColWidth(4) = 2600
    Fg1.ColWidth(5) = 900
    Fg1.ColWidth(6) = 900
   
    flag = 0
    For i = 0 To 6
        Fg1.ColAlignment(i) = 4
    Next i
        
    Call DB_Conn_MDB
    Call DBCall
    'Combo1.Text = "��ü"
    Call comboview
   

   
End Sub
Public Sub DBCall()
On Error Resume Next
Dim cnt
If mWorker_Location = "��ü" Then
    sql = "select * from data"
Else
    sql = "select * from data  where �۾��� = '" & mWorker_Location & "'"
End If
    If Not Run(sql) Then Exit Sub
    Rs.MoveFirst
    cnt = 0
    Fg1.Clear
    
    Fg1.TextMatrix(0, 0) = "��ȣ"
    Fg1.TextMatrix(0, 1) = "����"
    Fg1.TextMatrix(0, 2) = "�ֹε�Ϲ�ȣ"
    Fg1.TextMatrix(0, 3) = "�ø����ȣ"
    Fg1.TextMatrix(0, 4) = "�����"
    Fg1.TextMatrix(0, 5) = "����"
    Fg1.TextMatrix(0, 6) = "�۾���"
    Do While Not Rs.EOF
        If flag <> 1 Then
            Fg1.Rows = Fg1.Rows + 1
        End If
        
        With Fg1
            Fg1.TextMatrix(cnt + 1, 0) = cnt + 1
            Fg1.TextMatrix(cnt + 1, 1) = Trim(Rs.Fields("�̸�"))
            Fg1.TextMatrix(cnt + 1, 2) = Trim(Rs.Fields("�ֹε�Ϲ�ȣ"))
            Fg1.TextMatrix(cnt + 1, 3) = Trim(Rs.Fields("�ø����ȣ"))
            Fg1.TextMatrix(cnt + 1, 4) = Trim(Rs.Fields("�����"))
            Fg1.TextMatrix(cnt + 1, 5) = Trim(Rs.Fields("����"))
            Fg1.TextMatrix(cnt + 1, 6) = Trim(Rs.Fields("�۾���"))
        End With
        
        cnt = cnt + 1
        Rs.MoveNext
    Loop
    
        
        
End Sub

