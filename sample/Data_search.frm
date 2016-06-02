VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Data_search 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "Data Search"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11100
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   10440
   End
   Begin VB.CommandButton Command6 
      Caption         =   "닫기"
      Height          =   375
      Left            =   9960
      TabIndex        =   13
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "프린트"
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "엑셀출력"
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   10440
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "검색결과"
      ForeColor       =   &H8000000D&
      Height          =   8055
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   11775
      Begin VB.OptionButton Option3 
         Caption         =   "Event Mode"
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "누적선량 모드"
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Count Down Mode"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid Fg1 
         Height          =   6975
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   12303
         _Version        =   393216
         Cols            =   7
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   8880
         TabIndex        =   25
         Top             =   360
         Width           =   60
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "검색조건"
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   9480
         TabIndex        =   24
         Text            =   "전체"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "목록보기"
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker Dt1 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   15
         Top             =   915
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   116850689
         CurrentDate     =   42306
      End
      Begin VB.CommandButton Command1 
         Caption         =   "검색"
         Height          =   1095
         Left            =   10800
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Dt2 
         Height          =   375
         Index           =   0
         Left            =   6360
         TabIndex        =   17
         Top             =   915
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   116850689
         CurrentDate     =   42306
      End
      Begin MSComCtl2.DTPicker Dt1 
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   18
         Top             =   915
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   116850690
         CurrentDate     =   42306
      End
      Begin MSComCtl2.DTPicker Dt2 
         Height          =   375
         Index           =   1
         Left            =   7680
         TabIndex        =   19
         Top             =   915
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   116850690
         CurrentDate     =   42306
      End
      Begin VB.Label Label6 
         Caption         =   "작업장:"
         Height          =   255
         Left            =   8760
         TabIndex        =   23
         Top             =   435
         Width           =   735
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   1200
         X2              =   9960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "날짜 및 시간조회 :"
         Height          =   180
         Left            =   1200
         TabIndex        =   14
         Top             =   1005
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "~"
         Height          =   255
         Left            =   6000
         TabIndex        =   7
         Top             =   1005
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "시리얼 번호 :"
         Height          =   180
         Left            =   5880
         TabIndex        =   3
         Top             =   435
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "주민등록번호 :"
         Height          =   180
         Left            =   3000
         TabIndex        =   2
         Top             =   435
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "성명 :"
         Height          =   180
         Left            =   1200
         TabIndex        =   1
         Top             =   435
         Width           =   480
      End
   End
End
Attribute VB_Name = "Data_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim msql As String
Dim Align As Integer
Dim codeuser
Dim Location As String
Dim OptionFlag
Public Sub Usercode(User_code)
    codeuser = User_code
    Text2.Text = codeuser
    
End Sub

Private Sub Command1_Click()

IniTial

    Call DB_Conn_MDB2
    Call DBCall

    
    

End Sub

Public Sub IniTial()

Fg1.Clear
Fg1.Rows = 2
If OptionFlag = 1 Then

Fg1.TextMatrix(0, 0) = "번호"
Fg1.TextMatrix(0, 1) = "이름"
Fg1.TextMatrix(0, 2) = "주민등록번호"
Fg1.TextMatrix(0, 3) = "접속시간"
Fg1.TextMatrix(0, 4) = "누적선량"
Fg1.TextMatrix(0, 5) = "단위"
Fg1.TextMatrix(0, 6) = "작업장"
ElseIf OptionFlag = 2 Then
Fg1.TextMatrix(0, 0) = "번호"
Fg1.TextMatrix(0, 1) = "이름"
Fg1.TextMatrix(0, 2) = "주민등록번호"
Fg1.TextMatrix(0, 3) = "접속시간"
Fg1.TextMatrix(0, 4) = "누적선량"
Fg1.TextMatrix(0, 5) = "단위"
Fg1.TextMatrix(0, 6) = "작업장"
ElseIf OptionFlag = 3 Then
Fg1.TextMatrix(0, 0) = "번호"
Fg1.TextMatrix(0, 1) = "이름"
Fg1.TextMatrix(0, 2) = "주민등록번호"
Fg1.TextMatrix(0, 3) = "접속시간"
Fg1.TextMatrix(0, 4) = "선량율"
Fg1.TextMatrix(0, 5) = "단위"
Fg1.TextMatrix(0, 6) = "작업장"
End If

End Sub

Public Sub DBCall()
On Error Resume Next
Dim cnt
Dim mcnt


If Trim(Text2.Text) = "" Then
    MsgBox "주민등록번호 입력은 필수입니다.", vbOKOnly, "알림"

Else
    msql = "select * from Data where 주민등록번호 = '" & Trim(Text2.Text) & "' "
        
        If Not Run(msql) Then Exit Sub
        Rs.MoveFirst
       
        If Rs.EOF = True Then
            Label5.Caption = "0 건의 자료가 검색되었습니다."
        Else
            Location = Rs.Fields("작업장")
            
        End If
    If Option1.Value = True Then
        If Align = 1 Then
            sql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 >= '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 <= '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            sql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            sql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량 desc"
        ElseIf Align = 4 Then
            sql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량"
        End If
        
    ElseIf Option2.Value = True Then
        If Align = 1 Then
            sql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            sql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            sql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량 desc"
        ElseIf Align = 4 Then
            sql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량"
        End If
    ElseIf Option3.Value = True Then
        If Align = 1 Then
            sql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            sql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            sql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 선량율 desc"
        ElseIf Align = 4 Then
            sql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 선량율"
        End If
    End If
    
    
    
    
    If Not Run(sql) Then Exit Sub
    Rs.MoveFirst
    cnt = 0
    If Rs.EOF = True Then
        Label5.Caption = "0 건의 자료가 검색되었습니다."
    Else
        Do While Not Rs.EOF
            Fg1.Rows = Fg1.Rows + 1
            With Fg1
                Fg1.TextMatrix(cnt + 1, 0) = cnt + 1
                Fg1.TextMatrix(cnt + 1, 1) = Trim(Rs.Fields("성명"))
                Fg1.TextMatrix(cnt + 1, 2) = Trim(Rs.Fields("주민등록번호"))
                Fg1.TextMatrix(cnt + 1, 3) = Trim(Rs.Fields("시간"))
                If Option3.Value = True Then 'Format(day(pDay), "00")
'                 mapple = Left(Format(apple, "00000000"), 3)
'               mstring = mapple
'               mapple = Mid(Format(apple, "00000000"), 4)
'               mstring = mstring + "." + mapple
'
'
                
                Fg1.TextMatrix(cnt + 1, 4) = Trim(Rs.Fields("선량율"))
                Else
                Fg1.TextMatrix(cnt + 1, 4) = Trim(Rs.Fields("누적선량")) 'Format(Trim(Rs.Fields("누적선량")), "0000000.0")
                
                End If
                
                Fg1.TextMatrix(cnt + 1, 6) = Trim(Location)
                If Option1.Value = True Then
                    Fg1.TextMatrix(cnt + 1, 5) = "uSv"
                ElseIf Option2.Value = True Then
                    Fg1.TextMatrix(cnt + 1, 5) = "uSv"
                ElseIf Option3.Value = True Then
                    Fg1.TextMatrix(cnt + 1, 5) = "uSv/h"
                End If
                
            End With
            
            cnt = cnt + 1
            Rs.MoveNext
        Loop
        Label5.Caption = cnt & " 건의 자료가 검색되었습니다."
    End If
    
    
End If
    
    
        
        
End Sub


Private Sub Command3_Click()
    Fg1.Refresh
    Call DB_Conn_MDB
    Call DBCall
   
    
End Sub

Private Sub Command4_Click()
   Call DB_Conn_MDB2
   Dim XL As Object
    PrintFACO XL
    Set XL = Nothing
End Sub
Private Sub GetExcel(obj As Object)
    Set obj = Nothing
    On Error Resume Next
    err.Clear
    Set obj = GetObject(, "Excel.Application")
  
    If err.Number <> 0 Then
        err.Clear
        Set obj = Nothing
        Set obj = CreateObject("Excel.Application")
        obj.Workbooks.Open FileName:="C:\NOW_Cop" & "\DataSearch.xlsx"
       
        
    ElseIf err.Number = 0 Then
        obj.Workbooks.Open FileName:="C:\NOW_Cop" & "\DataSearch.xlsx"
        obj.Application.Visible = True
        
    End If
   
End Sub

Private Sub PrintFACO(obj As Object)
Dim cnt
Dim ModeName

    GetExcel obj
    
    If Trim(Text2.Text) = "" Then
    MsgBox "주민등록번호 입력은 필수입니다.", vbOKOnly, "알림"

Else
    msql = "select * from Data where 주민등록번호 = '" & Trim(Text2.Text) & "' "
        
        If Not Run(msql) Then Exit Sub
        Rs.MoveFirst
       
        If Rs.EOF = True Then
            Label5.Caption = "0 건의 자료가 검색되었습니다."
        Else
            Location = Rs.Fields("작업장")
            
        End If
    If Option1.Value = True Then
        If Align = 1 Then
            sql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            sql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            sql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량 desc"
        ElseIf Align = 4 Then
            sql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량"
        End If
        
    ElseIf Option2.Value = True Then
        If Align = 1 Then
            sql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            sql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            sql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량 desc"
        ElseIf Align = 4 Then
            sql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량"
        End If
    ElseIf Option3.Value = True Then
        If Align = 1 Then
            sql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            sql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            sql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 선량율 desc"
        ElseIf Align = 4 Then
            sql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 선량율"
        End If
    End If
    
    '1,2->조회날짜1 1,4 조회날짜2
    '2,2 출력모드
    '5,1 ->이름 5,2 주민등록번호 5,4접속시간 5,5 누적선량 5,6 단위 5,7 작업장
    
    obj.Application.Sheets("Sheet1").Cells(1, 2).Value = Dt1(0).Value
    obj.Application.Sheets("Sheet1").Cells(1, 4).Value = Dt2(0).Value
    
    
    If OptionFlag = 1 Then
        ModeName = "C 모드"
    ElseIf OptionFlag = 2 Then
        ModeName = "D 모드"
    ElseIf OptionFlag = 3 Then
        ModeName = "E 모드"
    End If
    
    obj.Application.Sheets("Sheet1").Cells(2, 2).Value = ModeName




    If Not Run(sql) Then Exit Sub
    If Not Rs.EOF = True Then
        Rs.MoveFirst
    End If
    
    cnt = 5
    If Rs.EOF = True Then
        Label5.Caption = "0 건의 자료가 검색되었습니다."
    Else
        Do While Not Rs.EOF
           
           
           
                obj.Application.Sheets("Sheet1").Cells(cnt, 1).Value = Trim(Rs.Fields("성명"))
                obj.Application.Sheets("Sheet1").Cells(cnt, 2).Value = Format(Trim(Rs.Fields("주민등록번호")), "######-#######")
                obj.Application.Sheets("Sheet1").Cells(cnt, 4).Value = Trim(Rs.Fields("시간"))
                
                If Option3.Value = True Then
                    obj.Application.Sheets("Sheet1").Cells(cnt, 5).Value = Trim(Rs.Fields("선량율"))
    
                Else
                    obj.Application.Sheets("Sheet1").Cells(cnt, 5).Value = Trim(Rs.Fields("누적선량"))
                    'Fg1.TextMatrix(cnt + 1, 4) = Format(CInt(Trim(Rs.Fields("누적선량"))) / 10, "######0.0")
                End If
                
                If Option1.Value = True Then
                     obj.Application.Sheets("Sheet1").Cells(cnt, 6).Value = "uSv"
                ElseIf Option2.Value = True Then
                    obj.Application.Sheets("Sheet1").Cells(cnt, 6).Value = "uSv"
                ElseIf Option3.Value = True Then
                    obj.Application.Sheets("Sheet1").Cells(cnt, 6).Value = "uSv/h"
                End If
                                        
                obj.Application.Sheets("Sheet1").Cells(cnt, 7).Value = Trim(Location)

            cnt = cnt + 1
            Rs.MoveNext
        Loop
      
    End If
    
    
End If
    

    obj.Application.Sheets("Sheet1").SaveAs App.Path & "DataSearch" & Format(Now, "YYMMDDHHMMSS") & ".xlsx"
    obj.Application.ActiveWorkbook.Close
End Sub
Private Sub Command5_Click()
Dim conStr As String
    Dim DBPath As String
    Dim DBFile As String
    Dim MsgMake As String
    Dim strSql As String
    Dim i As Integer
   
    Dim ModeName
        strSql = ""
        Call DB_Conn_MDB2
       
       Load Data2
       
        
        
     
        If Option1.Value = True Then
        If Align = 1 Then
            strSql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            strSql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            strSql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량 desc"
        ElseIf Align = 4 Then
            strSql = "select * from dataC where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량"
        End If
        
    ElseIf Option2.Value = True Then
        If Align = 1 Then
            strSql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            strSql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            strSql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량 desc"
        ElseIf Align = 4 Then
            strSql = "select * from dataD where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 누적선량"
        End If
     
    ElseIf Option3.Value = True Then
        If Align = 1 Then
            strSql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간 desc"
        ElseIf Align = 2 Then
            strSql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 시간"
        ElseIf Align = 3 Then
            strSql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 선량율 desc"
        ElseIf Align = 4 Then
            strSql = "select * from dataE where 주민등록번호 = '" & Trim(Text2.Text) & "' and 시간 > '" & Dt1(0) & " " & Format(Dt1(1), "AMPM hh:mm:ss") & "' and 시간 < '" & Dt2(0) & " " & Format(Dt2(1), "AMPM hh:mm:ss") & "' order by 선량율"
        End If
       
    End If
    '접속 연결 문자열
    Debug.Print strSql
    If Not Run(strSql) Then Exit Sub
'    If Not (Rs.EOF Or Rs.BOF) Then
'        Rs.MoveFirst
'    Else
'        sql = ""
'    End If
'
  
    
 
  
    
    
    Set Con = New ADODB.Connection

    '파일의 경로
    'DBPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
       DBPath = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    'DBPath = IIf(Right("C:\선량계\", 1) = "\", App.Path, App.Path + "\")
    '파일의 이름
    DBFile = "Po.MDB"
    'DB 구조 불러옴
    Call DB_Info
    'Provider=Microsoft.Jet.OLEDB.4.0;Password="";User ID=admin;Data Source=App.Path+"\"+"MDB\";Persist Security Info=True;Jet OLEDB:Database Password=1234

    '연결 문자열 설정
    conStr = "Provider=Microsoft.Jet.OLEDB.4.0;"
    conStr = conStr + "Data Source=" + DBPath + DBFile + ";"
    conStr = conStr + "User ID=admin"
    conStr = conStr + ";Jet OLEDB:Database Password=1234;"

    'Data2.Connections("Connection2").ConnectionString = conStr
    
    
    

    If OptionFlag = 1 Then
        ModeName = "C 모드"
        DataReport2.Sections("Section1").Controls("Label7").Caption = "uSv"

        Data2.Commands("Command2").CommandText = strSql
        DataReport2.DataMember = "Command2"
        
        With Data2
        If .rsCommand2.State <> 0 Then .rsCommand2.Close

        End With
    ElseIf OptionFlag = 2 Then
        ModeName = "D 모드"
        DataReport2.Sections("Section1").Controls("Label7").Caption = "uSv"
      
        Data2.Commands("Command1").CommandText = strSql
        DataReport2.DataMember = "Command1"
        
        
        With Data2
        If .rsCommand1.State <> 0 Then .rsCommand1.Close

        End With

    ElseIf OptionFlag = 3 Then
        ModeName = "E 모드"
        DataReport2.Sections("Section1").Controls("Text7").DataField = "선량율"
        DataReport2.Sections("Contents").Controls("Label4").Caption = "선량율"
        DataReport2.Sections("Section1").Controls("Label7").Caption = "uSv/h"
   
        Set DataReport2.DataSource = Data2
        Data2.Commands("Command3").CommandText = strSql
        DataReport2.DataMember = "Command3"
        
         
        With Data2
        If .rsCommand3.State <> 0 Then .rsCommand3.Close
        End With

    End If
   
    'Combo1
    If Not Rs.EOF = True Then
        Rs.MoveFirst
    End If
DataReport2.Sections("Section4").Controls("PrintMode").Caption = ModeName
DataReport2.Sections("Section4").Controls("Date1").Caption = Dt1(0).Value
DataReport2.Sections("Section4").Controls("date2").Caption = Dt2(0).Value
DataReport2.Sections("Section1").Controls("Label9").Caption = Combo1.Text




'Data2.Commands("Command2").CommandText = sql
DataReport2.Show


'Set Data = Nothing


End Sub



Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
All_worker.Worker_Location (Trim(Combo1.Text))
All_worker.Show
End Sub



Private Sub Fg1_Click()

If Fg1.ColSel = 3 Then
    If Fg1.TextMatrix(Fg1.RowSel, 3) = "" Then
        If Align = 1 Then
            Align = 2
        Else
            Align = 1
        End If
        Command1_Click
    End If
ElseIf Fg1.ColSel = 4 Then
    If Fg1.TextMatrix(Fg1.RowSel, 4) = "" Then
        If Align = 3 Then
            Align = 4
        Else
            Align = 3
        End If
        Command1_Click
    End If
End If

End Sub

Private Sub Form_Load()

    Fg1.ColWidth(0) = 600
    Fg1.ColWidth(1) = 1000
    Fg1.ColWidth(2) = 2200
    Fg1.ColWidth(3) = 2700
    Fg1.ColWidth(4) = 2200
    Fg1.ColWidth(5) = 1300
    Fg1.ColWidth(6) = 1300
    
   
    Fg1.TextMatrix(0, 0) = "번호"
    Fg1.TextMatrix(0, 1) = "이름"
    Fg1.TextMatrix(0, 2) = "주민등록번호"
    Fg1.TextMatrix(0, 3) = "접속시간"
    Fg1.TextMatrix(0, 4) = "선량율"
    Fg1.TextMatrix(0, 5) = "단위"
    Fg1.TextMatrix(0, 6) = "작업장"

    Fg1.ColAlignment(0) = 4
    Fg1.ColAlignment(1) = 4
    Fg1.ColAlignment(2) = 4
    Fg1.ColAlignment(3) = 4
    Fg1.ColAlignment(4) = 4
    Fg1.ColAlignment(5) = 4
    Fg1.ColAlignment(6) = 4
    
    Dt1(0).Value = Date - 1
    Dt1(1).Value = time
    Dt2(0).Value = Date
    Dt2(1).Value = time
    Align = 1
    
    Timer1.Interval = 10
    
End Sub


Private Sub Option1_Click()
OptionFlag = 1

If Option1.Value = True Then
    Command1_Click
ElseIf Option2.Value = True Then
    Command1_Click
ElseIf Option3.Value = True Then
    Command1_Click
End If

End Sub

Private Sub Option2_Click()
OptionFlag = 2
If Option1.Value = True Then
    Command1_Click
ElseIf Option2.Value = True Then
    Command1_Click
ElseIf Option3.Value = True Then
    Command1_Click
End If

End Sub

Private Sub Option3_Click()
OptionFlag = 3
If Option1.Value = True Then
    Command1_Click
ElseIf Option2.Value = True Then
    Command1_Click
ElseIf Option3.Value = True Then
    Command1_Click
End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 0
    OptionFlag = 1

    Command1_Click
End Sub
