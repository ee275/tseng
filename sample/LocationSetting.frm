VERSION 5.00
Begin VB.Form LocationSetting 
   Caption         =   "작업장 설정"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton del 
      Caption         =   "삭제"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton add 
      Caption         =   "추가"
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
      Caption         =   "삭제할 작업장"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "추가할 작업장"
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
sql = "insert into location(작업장) values('" & Text1.Text & "')"
If Not Run(sql) Then Exit Sub
MsgBox Text1.Text + "이(가)저장되었습니다.", vbOKOnly, "알림"
 Call comboview
End Sub



Private Sub del_Click()
     
     apple = Combo1.Text
     If apple = "전체" Then
        MsgBox "전체 목록은 지우실수 없습니다", vbOKOnly, "알림"
     Else
        sql = "delete From location where 작업장 ='" & apple & "'"
        If Not Run(sql) Then Exit Sub
        MsgBox Combo1.Text + "이(가) 삭제 되었습니다.", vbOKOnly, "알림"
        Call comboview
     End If
End Sub

Private Sub Form_Load()

    Call comboview

End Sub

Public Sub comboview()

 sql = "select * from location where 작업장 <> '전체'"
    If Not Run(sql) Then Exit Sub
     
     If Not (Rs.EOF Or Rs.BOF) Then
         Rs.MoveFirst
        
         Combo1.Clear '이건 콤보1에 모든내용을 지우는 부분임니다.
         Do While Not Rs.EOF
             
          With Combo1
              .AddItem Trim(Rs.Fields("작업장"))
          End With
             Rs.MoveNext
         Loop
         Text1.Text = ""
    End If
    
End Sub
