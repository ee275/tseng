VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form worker_edit 
   BorderStyle     =   4  '���� ���� â
   Caption         =   "�۾��� ���� ����"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   8520
      TabIndex        =   17
      Top             =   8880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton exit 
      Caption         =   "������"
      Height          =   495
      Left            =   10440
      TabIndex        =   15
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "�۾��� ���� ����"
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   6600
      Width           =   12975
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   5760
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton edit 
         Caption         =   "��������"
         Height          =   855
         Left            =   11280
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "����"
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   7200
         TabIndex        =   7
         Top             =   480
         Width           =   3855
         Begin VB.OptionButton Option6 
            Caption         =   "������"
            Height          =   180
            Left            =   2520
            TabIndex        =   13
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "����"
            Height          =   255
            Left            =   1440
            TabIndex        =   12
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "�ް�"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "������ȯ"
            Height          =   180
            Left            =   2520
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "���"
            Height          =   180
            Left            =   1440
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "����"
            Height          =   180
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "�۾���"
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "�ֹε�Ϲ�ȣ      (��:8307231059475)"
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "�ø��� ��ȣ"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "�۾��� ����"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6015
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   10610
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "worker_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim sql1 As String
Dim i
Dim temp As String

Private Sub Command1_Click()
sql = " Delete from data where �ֹε�Ϲ�ȣ = '" + ListView1.SelectedItem.SubItems(2) + "'"

If Not Run(sql) Then Exit Sub
MsgBox "������ �����Ǿ����ϴ�.", vbOKOnly, "�˸�"

Form_Load
MainScreen.DBCall
End Sub

Private Sub edit_Click()
On Error Resume Next

If Len(Trim(Text3.Text)) <> 13 Then
        MsgBox "�ֹε�Ϲ�ȣ �ڸ����� Ʋ�Ƚ��ϴ�.", vbOKOnly, "�˸�"
        Exit Sub
End If

sql = "update data set �̸�= '" & Trim(Text1.Text) & "', �ø����ȣ='" & Trim(Text2.Text)
sql = sql & "',����='"

sql1 = "update data set �̸�= '" & Trim(Text1.Text) & "', �ø����ȣ='" & ""

If Option1.Value = True Then
    sql = sql & "����"
ElseIf Option2.Value = True Then
    sql = sql & "���"
    temp = "���"
ElseIf Option3.Value = True Then
    sql = sql & "������ȯ"
    temp = "������ȯ"
ElseIf Option4.Value = True Then
    sql = sql & "�ް�"
ElseIf Option5.Value = True Then
    sql = sql & "����"
ElseIf Option6.Value = True Then
    sql = sql & "������"
End If
sql = sql + "',�ֹε�Ϲ�ȣ='" & Trim(Text3.Text) & "',�۾���='" & Trim(Combo1.Text) & "'  where �ֹε�Ϲ�ȣ='" & ListView1.SelectedItem.SubItems(2) & "'"
sql1 = sql1 + "',�ֹε�Ϲ�ȣ='" & Trim(Text3.Text) & "'  where �ֹε�Ϲ�ȣ='" & ListView1.SelectedItem.SubItems(2) & "'"
If Not Run(sql) Then Exit Sub

    If temp = "���" Then
    
        If Not Run(sql1) Then Exit Sub
    ElseIf temp = "������ȯ" Then
        If Not Run(sql1) Then Exit Sub
    End If

temp = ""
MsgBox "������ ����Ǿ����ϴ�.", vbOKOnly, "�˸�"
Form_Load
MainScreen.DBCall

End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
If admin = True Then
    Command1.Visible = True
End If
Call DBCall
End Sub
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Combo1.Text = ""
End Sub
Public Sub DBCall()
On Error Resume Next
    sql = "select * from data'"
    
    If Not Run(sql) Then Exit Sub
    With ListView1
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.add 1, , "��ȣ", 800
        .ColumnHeaders.add 2, , "����", 1600, 2
        .ColumnHeaders.add 3, , "�ֹε�Ϲ�ȣ", 2000, 2
        .ColumnHeaders.add 4, , "�ø����ȣ", 2400, 2
        .ColumnHeaders.add 5, , "�����", 2700, 2
        .ColumnHeaders.add 6, , "����", 1600, 2
        .ColumnHeaders.add 7, , "�۾���", 1200, 2
        Call DBref
    End With
    Rs.MoveFirst
    If Rs.EOF = True Then
        edit.Enabled = False
    Else
        edit.Enabled = True
    End If
    
        
        
End Sub
Private Sub DBref()
Dim cnt
cnt = 1
    With ListView1
        .ListItems.Clear
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveFirst
            
            Do While Not Rs.EOF
                .ListItems.add , , cnt
                .ListItems(.ListItems.Count).SubItems(1) = IIf(IsNull(Rs(0).Value), "", Trim(Rs(0).Value))
                .ListItems(.ListItems.Count).SubItems(2) = IIf(IsNull(Rs(1).Value), "", Trim(Rs(1).Value))
                .ListItems(.ListItems.Count).SubItems(3) = IIf(IsNull(Rs(2).Value), "", Trim(Rs(2).Value))
                .ListItems(.ListItems.Count).SubItems(4) = IIf(IsNull(Rs(9).Value), "", Trim(Rs(9).Value))
                .ListItems(.ListItems.Count).SubItems(5) = IIf(IsNull(Rs(7).Value), "", Trim(Rs(7).Value))
                .ListItems(.ListItems.Count).SubItems(6) = IIf(IsNull(Rs(8).Value), "", Trim(Rs(8).Value))
            cnt = cnt + 1
            Rs.MoveNext
            Loop
        End If
    End With
    
    
End Sub


Private Sub ListView1_Click()
On Error Resume Next
Call comboview
Text1.Text = ListView1.SelectedItem.SubItems(1)
Text2.Text = ListView1.SelectedItem.SubItems(3)
Text3.Text = ListView1.SelectedItem.SubItems(2)
Combo1.Text = ListView1.SelectedItem.SubItems(6)
If ListView1.SelectedItem.SubItems(5) = "����" Then
    Option1.Value = True
ElseIf ListView1.SelectedItem.SubItems(5) = "���" Then
    Option2.Value = True
ElseIf ListView1.SelectedItem.SubItems(5) = "���" Then
    Option3.Value = True
ElseIf ListView1.SelectedItem.SubItems(5) = "�ް�" Then
    Option4.Value = True
ElseIf ListView1.SelectedItem.SubItems(5) = "����" Then
    Option5.Value = True
ElseIf ListView1.SelectedItem.SubItems(5) = "������" Then
    Option6.Value = True
End If




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
