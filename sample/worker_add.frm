VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form worker_add 
   BorderStyle     =   4  '���� ���� â
   Caption         =   "�۾��� ���"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6375
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "worker_add.frx":0000
      Left            =   2520
      List            =   "worker_add.frx":0002
      TabIndex        =   11
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton Connect 
      Caption         =   "����"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton exit 
      Caption         =   "������"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton add 
      Caption         =   "���"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Worker_id 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox worker_name 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox clover_serial 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5400
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   6
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.Label Label5 
      Caption         =   "�۾���"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "(��:8307231059475)"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "�ֹε�Ϲ�ȣ "
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "�۾��� ����"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "��� �ø��� ��ȣ"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "worker_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim mbuf1
Dim tmp

Private Sub add_Click()

If Trim(clover_serial.Text) = "" Or Trim(worker_name.Text) = "" Or Trim(Worker_id.Text) = "" Then
    MsgBox "����,�ø����ȣ,�ֹε�Ϲ�ȣ�� ������ �ֽ��ϴ�.", vbOKOnly, "�˸�"
Else
    If Len(Trim(Worker_id.Text)) <> 13 Then
        MsgBox "�ֹε�Ϲ�ȣ �ڸ����� Ʋ�Ƚ��ϴ�.", vbOKOnly, "�˸�"
    ElseIf Combo1.Text = "" Then
        MsgBox "�۾����� �������ּ���.", vbOKOnly, "�˸�"
    Else
        sql = "inseRt into data(�̸�,�ֹε�Ϲ�ȣ,�����,�ø����ȣ,����,�۾���,����,���ϴ�������) values('" & Trim(worker_name.Text) & "','" & Trim(Worker_id.Text) & "','" & CStr(Now) & "','" & Trim(clover_serial.Text) & "','����','" & Trim(Combo1.Text) & "','uSv','0')"
        If Not Run(sql) Then Exit Sub
       
        MsgBox "��ϵǾ����ϴ�.", vbOKOnly, "�˸�"
        MainScreen.Combo1.Text = "��ü"
        
        MainScreen.DBCall
        Unload Me
    End If
    
    
End If
MainScreen.Combo1.Text = "��ü"
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Connect_Click()
On Error GoTo err:
Dim i

If mbuf1 = "" Then
    commCmd = "S"
     If MainScreen.MSComm1.PortOpen = False Then
     MainScreen.MSComm1.PortOpen = True
        MainScreen.MSComm1.Output = commCmd
    'If MSComm1.PortOpen = False Then
'        MSComm1.PortOpen = True
'        MSComm1.Output = commCmd
'
    Else
        MainScreen.MSComm1.PortOpen = False
    End If
     
Else

End If
   
Exit Sub

err:
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call comboview
    'Combo1.Text = "��ü"
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
        
    End If
    
End Sub
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Combo1.Text = ""
End Sub
Private Sub MSComm1_OnComm()
Dim i

If MSComm1.CommEvent = comEvReceive Then
    SaveSetting "COMM", "PORT", "SETTING", MSComm1.CommPort
    mbuf1 = mbuf1 + MSComm1.Input
    
    If Left(commCmd, 1) = "S" Then
        If Len(mbuf1) > 21 Then
            For i = 1 To Len(mbuf1)
                If Not Asc(Mid(mbuf1, i, 1)) = 9 Then
                    tmp = tmp + Mid(mbuf1, i, 1)
                End If
            Next i
            sql = "select * from data where �ø����ȣ = '" & tmp & "'"
            If Not Run(sql) Then Exit Sub
            
            If Rs.BOF Then
                clover_serial.Text = tmp
                
              
            Else
                 If MsgBox("���� �����ø� ����� �����մϴ� ����� �����ϸ� ��ġ�� ������ �������ϴ�!", vbYesNo, "�˸�") = vbYes Then
                
                    Auto_worker_enter.Show
                    Auto_worker_enter.Wiorker_name = Trim(Rs.Fields("�̸�"))
                    Auto_worker_enter.Clover_sn.Text = Rs.Fields("�ø����ȣ")
                    Auto_worker_enter.Worker_id = Trim(Rs.Fields("�ֹε�Ϲ�ȣ"))
                    Auto_connect_start.Show
                 Else
                
               
                
                 End If
                
            End If
            
            tmp = ""
            mbuf1 = ""
            
        End If
    ElseIf Left(commCmd, 1) = "T" Then
        If Right(mbuf1, 1) = "R" Then
            For i = 1 To Len(mbuf1)
                If Not Asc(Mid(mbuf1, i, 1)) = 9 Then
                    tmp = tmp + Mid(mbuf1, i, 1)
                End If
            Next i
            'Auto_connect_start.Text1.Text = tmp
            tmp = ""
            mbuf1 = ""
        End If
    End If
    
    
    
    
End If
MSComm1.PortOpen = False
Exit Sub

End Sub
