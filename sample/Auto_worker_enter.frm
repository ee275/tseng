VERSION 5.00
Begin VB.Form Auto_worker_enter 
   BorderStyle     =   4  '���� ���� â
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ��"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Worker_id 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Clover_sn 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Wiorker_name 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "���� �Ϸ��� �ڵ����� �˴ϴ�."
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "�ֹ� ��� ��ȣ"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "��ġ S/N"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "�۾��� ����"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "���� ���� �� ��ġ�� ����"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Auto_worker_enter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim mCmdString As String
Dim mSerial As String
Dim mWorkerid As String




Private Sub Command1_Click()

Dim mText As String '�ؽ�Ʈ �о �����ϴ� �κ� ��ü ������

 
Dim todaytotal
Dim monthtotal
Dim day
Dim mDay
Dim dd
Dim apple
Dim dmodetotal
Dim userid
Dim mapple
Dim mstring As String
Dim mmapple
    commCmd = ""
    commCmd = mCmdString
    mFlag_timer = 1
    
    Auto_connect_end.Show
    
    MainScreen.MSComm1.Output = commCmd + vbCrLf
    mSerial = Clover_sn.Text
    mWorkerid = Worker_id.Text
    


Unload Me

End Sub
Public Sub Txtinsert()

Dim mText As String '�ؽ�Ʈ �о �����ϴ� �κ� ��ü ������

 
Dim todaytotal
Dim monthtotal
Dim day
Dim mDay
Dim dd
Dim apple
Dim dmodetotal
Dim userid
Dim mapple
Dim mstring As String
Dim mmapple

    

 '�ؽ�Ʈ ���� ���� �κ�
  FreeFile
  Open "tmpData.txt" For Input As #1
    Line Input #1, mText
  Close #1



  
  Call mDb_handling(mText, mSerial)
  'sql = "select * from monthly where �ֹε�Ϲ�ȣ ='" & mWorkerid & "'" ������������
    'If Not Run(sql) Then Exit Sub
    'Rs.MoveFirst
  
  
  
  
    sql = "select * from data where �ø����ȣ = '" & mSerial & "'"
    If Not Run(sql) Then Exit Sub
    Rs.MoveFirst
     Do While Not Rs.EOF
      
               todaytotal = IIf(IsNull(Rs.Fields("���ϴ�������")), "", Trim(Rs.Fields("���ϴ�������")))
               'monthtotal = IIf(IsNull(Rs.Fields("�̹��޴�������")), "", Trim(Rs.Fields("�̹��޴�������")))
               userid = IIf(IsNull(Rs.Fields("�ֹε�Ϲ�ȣ")), "", Trim(Rs.Fields("�ֹε�Ϲ�ȣ")))

        Rs.MoveNext
        Loop
    
    
    sql = "select * from monthly where ���� = '" & Date & "' AND �ֹι�ȣ= '" & userid & "'"
  
    If Not Run(sql) Then Exit Sub
     If Not Rs.EOF Then
    Rs.MoveFirst
    
     Do While Not Rs.EOF
               
               apple = Rs.Fields("������")
               todaytotal = apple
              
                
        Rs.MoveNext
    
        Loop

             sql = "update data set ���ϴ�������='" & Format(Val(apple) / 10, "######0.0") & "'  where �ֹε�Ϲ�ȣ = '" & userid & "'"
                If Not Run(sql) Then Exit Sub
'
             sql = "update data set ���ӽð�='" & Now & "'  where �ֹε�Ϲ�ȣ = '" & userid & "'"
             If Not Run(sql) Then Exit Sub
                   
         End If
         



End Sub
Private Sub Text5_Change()

End Sub

Private Sub Text7_Change()

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Label7_Click()

End Sub


Private Sub Form_Load()
Dim a
Dim apple
a = Weekday(Now)

If a = 2 Then
    apple = 1
ElseIf a = 3 Then
    apple = 2
ElseIf a = 4 Then
     apple = 3
ElseIf a = 5 Then
     apple = 4
ElseIf a = 6 Then
     apple = 5
ElseIf a = 7 Then
    apple = 6
ElseIf a = 1 Then
     apple = 0
End If

mCmdString = "T" & Right(Year(Now), 2) & Format(Month(Now), "00") & Format(day(Now), "00") & Format(apple, "00") & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
End Sub

