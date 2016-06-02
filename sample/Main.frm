VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form MainScreen 
   Caption         =   "Data Management Tool "
   ClientHeight    =   10830
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15135
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   15135
   StartUpPosition =   2  '화면 가운데
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "조회날자"
      Height          =   615
      Left            =   11160
      TabIndex        =   18
      Top             =   5520
      Width           =   3375
      Begin MSComCtl2.DTPicker DateYMD 
         Height          =   420
         Index           =   0
         Left            =   1080
         TabIndex        =   19
         Top             =   120
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   65280
         Format          =   116850688
         UpDown          =   -1  'True
         CurrentDate     =   42491
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   2640
      Top             =   10080
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   7800
      TabIndex        =   13
      Top             =   10200
      Width           =   5415
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         TabIndex        =   15
         Top             =   195
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Main.frx":030A
         Left            =   3720
         List            =   "Main.frx":030C
         TabIndex        =   14
         Text            =   "전체"
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "기관번호"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "작업장:"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3375
      Left            =   5520
      OleObjectBlob   =   "Main.frx":030E
      TabIndex        =   11
      Top             =   6240
      Width           =   9255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "장비 연결"
      Height          =   495
      Left            =   13440
      TabIndex        =   10
      Top             =   10320
      Width           =   1455
   End
   Begin SysInfoLib.SysInfo SysInf 
      Left            =   1680
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   10200
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      InBufferSize    =   2048
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   9600
      Width           =   9615
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   270
         Left            =   8520
         TabIndex        =   12
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4800
         TabIndex        =   8
         Top             =   200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Top             =   200
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "이번달 작업자 총 누적 선량"
         Height          =   255
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "/"
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "금일 작업자 총 누적 선량"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "/"
         Height          =   135
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "명"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "총 작업자"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   9340
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
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   2655
      Left            =   240
      Picture         =   "Main.frx":5132
      Top             =   6360
      Width           =   5100
   End
   Begin VB.Image Image7 
      Enabled         =   0   'False
      Height          =   3540
      Left            =   1200
      Picture         =   "Main.frx":12372
      Top             =   5880
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Image Image6 
      Enabled         =   0   'False
      Height          =   3495
      Left            =   960
      Picture         =   "Main.frx":15777
      Top             =   5760
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Image Image5 
      Enabled         =   0   'False
      Height          =   3510
      Left            =   1200
      Picture         =   "Main.frx":189AC
      Top             =   5760
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Image Image4 
      Enabled         =   0   'False
      Height          =   3750
      Left            =   960
      Picture         =   "Main.frx":1B82E
      Top             =   5760
      Visible         =   0   'False
      Width           =   3510
   End
   Begin VB.Image Image3 
      Height          =   3660
      Left            =   960
      Picture         =   "Main.frx":1EA62
      Top             =   6000
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   3870
      Left            =   960
      Picture         =   "Main.frx":21B96
      Top             =   5760
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Menu worker 
      Caption         =   "작업자등록"
      Index           =   1
      Begin VB.Menu worker_add_click 
         Caption         =   "등록"
      End
      Begin VB.Menu worker_edit_click 
         Caption         =   "수정/삭제"
      End
   End
   Begin VB.Menu state_change 
      Caption         =   "작업자 상태변경"
      Index           =   2
      Begin VB.Menu state_normal_click 
         Caption         =   "정상"
      End
      Begin VB.Menu state_vacation_click 
         Caption         =   "휴가"
      End
      Begin VB.Menu state_businesstrip_click 
         Caption         =   "출장"
      End
      Begin VB.Menu state_change_click 
         Caption         =   "업무전환"
      End
      Begin VB.Menu state_trouble_click 
         Caption         =   "장비고장"
      End
   End
   Begin VB.Menu search_click 
      Caption         =   "조회"
      Index           =   3
      Begin VB.Menu data_search_click 
         Caption         =   "데이터 검색"
      End
   End
   Begin VB.Menu p_click 
      Caption         =   "출력"
      Index           =   4
      Begin VB.Menu print_click 
         Caption         =   "인쇄"
      End
      Begin VB.Menu file_print_click 
         Caption         =   "파일출력"
      End
   End
   Begin VB.Menu number_click 
      Caption         =   "기관번호"
      Index           =   5
      Begin VB.Menu company_number_add_click 
         Caption         =   "기관번호입력"
      End
   End
   Begin VB.Menu Setting 
      Caption         =   "환경설정"
      Begin VB.Menu location_setting 
         Caption         =   "작업장등록"
      End
      Begin VB.Menu CommSet 
         Caption         =   "통신설정"
      End
      Begin VB.Menu Password 
         Caption         =   "비밀번호변경"
      End
   End
   Begin VB.Menu reset 
      Caption         =   "장비리셋"
   End
   Begin VB.Menu end_click 
      Caption         =   "종료"
      Index           =   6
      Begin VB.Menu exit 
         Caption         =   "종료하기"
      End
   End
End
Attribute VB_Name = "MainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Option Explicit

Dim sql As String
Dim i As Integer
Dim mbuf1
Dim tmp
Dim dmodetotal
Dim comboText
Dim TodayWorker
Dim monthWorker
Dim time As Integer





Private Sub Combo1_Click()
DBCall
End Sub

Private Sub Combo1_GotFocus()

Call comboview
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Combo1.Text = ""
End Sub

Private Sub Command1_Click()
On Error GoTo err:
Dim i

If mbuf1 = "" Then
    commCmd = "S"
    If MSComm1.PortOpen = False Then
        
        
        MSComm1.PortOpen = True
        MSComm1.Output = commCmd
        
        
    Else
        MSComm1.PortOpen = False
        MSComm1.PortOpen = True
        MSComm1.Output = commCmd
    End If
Else

End If
    
Exit Sub

err:
MsgBox err.Description

End Sub

Private Sub CommSet_Click()
Comm.Show
End Sub

Private Sub company_number_add_click_Click()
    company_number_add.Show
    
End Sub

Private Sub data_search_click_Click()
Data_search.Show

End Sub

Private Sub number_add_click_Click()

End Sub
Public Sub DBCall()
If Trim(Combo1.Text) = "전체" Then
    sql = "select * from data where 상태 <> '퇴사'"
Else
    sql = "select * from data where 상태 <> '퇴사' AND 작업장 = '" & Combo1.Text & "'"
End If
    
    If Not Run(sql) Then Exit Sub
    With ListView1
        .ListItems.Clear
        .ColumnHeaders.Clear
        For i = 0 To Rs.Fields.Count - 1
            If i = 0 Then
                .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 800
            ElseIf i = 1 Then
                .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 1500
            ElseIf i = 2 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 2000
            ElseIf i = 3 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 2400
            ElseIf i = 4 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 1300
            ElseIf i = 5 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 1500
            ElseIf i = 6 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 800
            ElseIf i = 7 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 800
            ElseIf i = 8 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 800
            
            ElseIf i = 9 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 2250
            ElseIf i = 10 Then
            .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 1000
                
            Else
                .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 500
                
            End If
        Next
        Call DBref
    End With
    
    
        
        
End Sub

Public Sub DBref()
Dim apple
Dim pDay
Dim nDay
Dim tmp
Dim mTmp
Dim userid As String

Dim mTotal_sql As String
Dim msql As String
Dim Monthly_total
Dim mon
Dim mNow
    With ListView1
        .ListItems.Clear
        .Sorted = False
        If Not (Rs.EOF Or Rs.BOF) Then
        Monthly_total = ""
        TodayWorker = 0
        monthWorker = 0
            Rs.MoveFirst
            
            Do While Not Rs.EOF
                .ListItems.add , , Trim(Rs(0))
                For i = 1 To Rs.Fields.Count - 1
                'if3 일때 최종접속시간이 나옴
               
                    userid = IIf(IsNull(Rs(1).Value), "", Trim(Rs(1).Value))
                
                
                If i = 3 Then
                 '"T" & Right(Year(Now), 2) & Format(Month(Now), "00") & Format(day(Now), "0
                    pDay = Rs(i).Value
                    nDay = Format(day(pDay), "00")
                    
                    If nDay <> Format(day(Now), "00") Then
                        i = 4
                    Else
                        .ListItems(.ListItems.Count).SubItems(i) = IIf(IsNull(Rs(i).Value), "", Trim(Rs(i).Value))

                    End If
                    
                
              '      month = Format(month(IIf(IsNull(Rs(i).Value), "", Trim(Rs(i).Value))))
              
                ElseIf i = 4 Then '금일누적선량
                
                mTotal_sql = "select * from monthly where 주민번호 = '" & userid & "' AND 날자 = '" & Date & "' "
                    If Not mRun(mTotal_sql) Then Exit Sub
                    
                    If Not (mRs.EOF Or mRs.BOF) Then
                        tmp = mRs.Fields("누적량")
                       .ListItems(.ListItems.Count).SubItems(i) = Format(tmp / 10, "######0.0")
                    Else
                        tmp = 0
                        .ListItems(.ListItems.Count).SubItems(i) = Format(tmp / 10, "######0.0")
                    End If
                 
                
                   TodayWorker = TodayWorker + tmp
                    
                    
                    'tmp 를업뎃하면되
                    msql = "update data set 금일누적선량='" & Format(tmp / 10, "######0.0") & "' where 주민등록번호 = '" & userid & "'"
                    If Not mmRun(msql) Then Exit Sub
'
            
                ElseIf i = 5 Then '이번달 누적선량

                        mTotal_sql = "select * from monthly where 주민번호 = '" & userid & "' "
                        If Not mRun(mTotal_sql) Then Exit Sub
    
                        If Not (mRs.EOF Or mRs.BOF) Then
                                mRs.MoveFirst
        
                               
                                Do While Not mRs.EOF
                                    mTmp = mRs.Fields("날자")
                                    mon = Year(mTmp) & Month(mTmp)
                                    mNow = Year(Now) & Month(Now)
                                    If mNow = mon Then
                                      Monthly_total = Val(Monthly_total) + Val(mRs.Fields("누적량"))
                                    End If
                                    
                                         mRs.MoveNext
                                 Loop
                        
                        Else
                               Monthly_total = ""
                        End If
                            
                        
                       If Monthly_total = "" Then
                            tmp = 0
                            .ListItems(.ListItems.Count).SubItems(i) = tmp
                       Else
                             tmp = Format(Monthly_total / 10, "######0.0")
                             .ListItems(.ListItems.Count).SubItems(i) = tmp
                       End If
                       
                       
                       monthWorker = monthWorker + tmp
                             
                       
                       msql = "update data set 이번달누적선량='" & tmp & "' where 주민등록번호 = '" & userid & "'"
                       If Not mmRun(msql) Then Exit Sub
        '
                Else '나머지 항목 채우기
                    
                
                    .ListItems(.ListItems.Count).SubItems(i) = IIf(IsNull(Rs(i).Value), "", Trim(Rs(i).Value))
                End If
                
                Next
            apple = apple + 1
            Rs.MoveNext
            Loop
            Text1.Text = apple
        End If
    End With
    
    
End Sub


Private Sub exit_Click()
End
End Sub

Sub regWriteBinary(sRegKey, sRegValue, sBinaryData)
      
      Dim oShell, oFSO, oFile, oExec
      Dim sTempFile
      Set oShell = CreateObject("WScript.Shell")
      Set oFSO = CreateObject("Scripting.FileSystemObject")
      sTempFile = oShell.ExpandEnvironmentStrings("%temp%\RegWriteBinary.reg")
      Set oFile = oFSO.CreateTextFile(sTempFile, True)
      oFile.WriteLine ("Windows Registry Editor Version 5.00")
      oFile.WriteLine ("")
      oFile.WriteLine ("[" & sRegKey & "]")
      oFile.WriteLine (Chr(34) & sRegValue & Chr(34) & "=" & "hex:" & sBinaryData)
      oFile.Close
      Set oExec = oShell.Exec("REG IMPORT " & sTempFile)
     
End Sub

Private Sub file_print_click_Click()
 Dim XL As Object
 Call DB_Conn_MDB2
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
        obj.Workbooks.Open FileName:="C:\NOW_Cop\" & "Main.xlsx"
       
        
    ElseIf err.Number = 0 Then
        obj.Workbooks.Open FileName:="C:\NOW_Cop\" & "Main.xlsx"
        obj.Application.Visible = True
        
    End If
   
End Sub

Private Sub PrintFACO(obj As Object)
Dim cnt


    GetExcel obj
    
   
    sql = "select * from data where 상태 <> '퇴사'"
        
       
        
   
    
    '1,2->조회날짜1 1,4 조회날짜2
    '2,2 출력모드
    '5,1 ->이름 5,2 주민등록번호 5,4접속시간 5,5 누적선량 5,6 단위 5,7 작업장
    
    obj.Application.Sheets("Sheet1").Cells(1, 2).Value = Now
    obj.Application.Sheets("Sheet1").Cells(2, 2).Value = Text1.Text

    If Not Run(sql) Then Exit Sub
    Rs.MoveFirst
    cnt = 5
    If Rs.EOF = True Then
      
    Else
        Do While Not Rs.EOF
           
           '1이름 2주민등록번호 3시리얼번호 4 접속시간 5금일누적선량 6 이번달누적선량 7상태 8작업장
           
                obj.Application.Sheets("Sheet1").Cells(cnt, 1).Value = Trim(Rs.Fields("이름"))
                obj.Application.Sheets("Sheet1").Cells(cnt, 2).Value = Format(Trim(Rs.Fields("주민등록번호")), "######-#######")
                obj.Application.Sheets("Sheet1").Cells(cnt, 3).Value = Trim(Rs.Fields("시리얼번호"))
                obj.Application.Sheets("Sheet1").Cells(cnt, 4).Value = Trim(Rs.Fields("접속시간"))
                obj.Application.Sheets("Sheet1").Cells(cnt, 5).Value = Trim(Rs.Fields("금일누적선량"))
                obj.Application.Sheets("Sheet1").Cells(cnt, 6).Value = Trim(Rs.Fields("이번달누적선량"))
                obj.Application.Sheets("Sheet1").Cells(cnt, 7).Value = Trim(Rs.Fields("상태"))
                obj.Application.Sheets("Sheet1").Cells(cnt, 8).Value = Trim(Rs.Fields("작업장"))
              

            cnt = cnt + 1
            Rs.MoveNext
        Loop
      
    End If
    
    
   
    
    obj.Application.Sheets("Sheet1").SaveAs App.Path & "Main" & Format(Now, "YYMMDDHHMMSS") & ".xlsx"

    
    'obj.Application.Sheets("Sheet1").SaveAs App.Path & "Main" & Format(Now, "YYMMDDHHMMSS") & ".xlsx"
    obj.Application.ActiveWorkbook.Close
End Sub
Private Sub Form_Load()
Dim cnt
Dim i
Dim j

On Error Resume Next
If App.PrevInstance Then End
    
    time = 1
    
    Combo1.Text = "전체"
    Call DB_Conn_MDB
    
    
    Call DBCall
    Text2.Text = Format(TodayWorker / 10, "######0.0")
    Text4.Text = monthWorker
    
    Rs.MoveFirst
    If Not (Rs.EOF Or Rs.BOF) Then
        Do While Not Rs.EOF
            cnt = cnt + 1
        Rs.MoveNext
        Loop
    End If
    
Text1.Text = cnt
Text3.Text = IIf(IsNull(GetSetting("COMPANY", "NO", "VALUE")), "", GetSetting("COMPANY", "NO", "VALUE"))

    If FirstExe = True Then
        MsgBox "통신설정을 먼저 해주세요", vbOKOnly, "알림"
    Else
        MSComm1.CommPort = GetSetting("COMM", "PORT", "SETTING")
    End If
        For i = 1 To 7
               For j = 1 To 5
                   MSChart1.Row = i
                   MSChart1.Column = j
                   MSChart1.Data = 0
               Next j
                
            Next i
   
    
   Call comboview
    DateYMD(0).Value = Date
    
    
    
    
End Sub

Public Sub mRefresh()
Dim cnt
Dim i
Dim j


On Error Resume Next
If App.PrevInstance Then End

    Combo1.Text = "전체"
    Call DB_Conn_MDB2
    
    
    Call DBCall
    Text2.Text = Format(TodayWorker / 10, "######0.0")
    Text4.Text = monthWorker
    
    Rs.MoveFirst
    If Not (Rs.EOF Or Rs.BOF) Then
        Do While Not Rs.EOF
            cnt = cnt + 1
        Rs.MoveNext
        Loop
    End If
    
Text1.Text = cnt
Text3.Text = IIf(IsNull(GetSetting("COMPANY", "NO", "VALUE")), "", GetSetting("COMPANY", "NO", "VALUE"))

    If FirstExe = True Then
        MsgBox "통신설정을 먼저 해주세요", vbOKOnly, "알림"
    Else
        MSComm1.CommPort = GetSetting("COMM", "PORT", "SETTING")
    End If
        For i = 1 To 7
               For j = 1 To 5
                   MSChart1.Row = i
                   MSChart1.Column = j
                   MSChart1.Data = 0
               Next j
                
            Next i
   
    
   Call comboview
    
    
    
End Sub



Public Sub comboview()
Dim mmsql As String
Dim apple As String
apple = "전체"
 mmsql = "select * from location"
    If Not Run(mmsql) Then Exit Sub

     If Not (Rs.EOF Or Rs.BOF) Then
         Rs.MoveFirst

         Combo1.Clear '이건 콤보1에 모든내용을 지우는 부분임니다.
         Do While Not Rs.EOF

          With Combo1
              .AddItem Trim(Rs.Fields("작업장"))
          End With
             Rs.MoveNext
         Loop
      Else

        mmsql = "insert into location(작업장)values('" & Trim(apple) & "')"
         If Not Run(mmsql) Then Exit Sub
    End If
    Combo1.Text = "전체"

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True

End Sub
Private Sub ListView1_Click()
    
  Dim apple
  Dim day
  Dim mDay
  Dim dd As Integer
  Dim i, j
  Dim mapple
  Dim mstring As String
  Dim yy
  Dim mm
  Dim mYear
  Dim mMont
  Dim mmYear
  Dim mmmMont
  
  Dim pie
  Dim applepie
  Call DB_Conn_MDB2
  For i = 1 To 7
               For j = 1 To 5
                   MSChart1.Row = i
                   MSChart1.Column = j
                   MSChart1.Data = 0
               Next j
                
            Next i

  
sql = "select * from data where 주민등록번호"
If Not Run(sql) Then Exit Sub

        
If (Rs.EOF) Then
    MsgBox "사용자등록후 사용해주세요", vbOKOnly, "알립"

Else
sql = "select * from monthly where 주민번호 = '" & ListView1.SelectedItem.SubItems(1) & "'"
'sql = "select * from dataD where 주민등록번호 = '" & ListView1.SelectedItem.SubItems(1) & "'  AND 시리얼번호 = '" & ListView1.SelectedItem.SubItems(2) & "'"
    dmodetotal = 0
 
    
    If Not Run(sql) Then Exit Sub
    Debug.Print Rs.RecordCount
    Debug.Print sql
    'If Rs.RecordCount > 0 Then
    
    If Not Rs.EOF Then
    
    If Not (Rs.EOF Or Rs.BOF) Then
        Rs.MoveFirst
    
    
     Do While Not Rs.EOF
               day = Rs.Fields("날자")
               mDay = Left(day, 10)
               mm = Mid(mDay, 6, 2)
               mYear = Mid(mDay, 1, 4) 'DB에서 불러온값의 년도
               mMont = Mid(mDay, 6, 2) 'DB에서 불러온값의 달
               mmYear = DateYMD(0).Value
               pie = Mid(mmYear, 1, 4) '현재의 년도 적혀있는 달
               applepie = Mid(mmYear, 6, 2) '현재의 월
               
               
               
               
               
               If mYear = pie And mMont = applepie Then
                dd = dd + Mid(mDay, 9)

               apple = Val(Format(Rs.Fields("누적량"), "00000000"))
               dmodetotal = Val(Format(dmodetotal, "00000000")) + Val(Format(Rs.Fields("누적량"), "00000000"))
                
              If dd = 1 Then
                i = 1
                ElseIf dd = 2 Then
                i = 1
                ElseIf dd = 3 Then
                i = 1
                ElseIf dd = 4 Then
                i = 1
                ElseIf dd = 5 Then
                i = 1
                ElseIf dd = 6 Then
                i = 2
                ElseIf dd = 7 Then
                i = 2
                ElseIf dd = 8 Then
                i = 2
                ElseIf dd = 9 Then
                i = 2
                ElseIf dd = 10 Then
                i = 2
                ElseIf dd = 11 Then
                i = 3
                ElseIf dd = 12 Then
                i = 3
                ElseIf dd = 13 Then
                i = 3
                ElseIf dd = 14 Then
                i = 3
                ElseIf dd = 15 Then
                i = 3
                ElseIf dd = 16 Then
                i = 4
                ElseIf dd = 17 Then
                i = 4
                ElseIf dd = 18 Then
                i = 4
                ElseIf dd = 19 Then
                i = 4
                ElseIf dd = 20 Then
                i = 4
                ElseIf dd = 21 Then
                i = 5
                ElseIf dd = 22 Then
                i = 5
                ElseIf dd = 23 Then
                i = 5
                ElseIf dd = 24 Then
                i = 5
                ElseIf dd = 25 Then
                i = 5
                ElseIf dd = 26 Then
                i = 6
                ElseIf dd = 27 Then
                i = 6
                ElseIf dd = 28 Then
                i = 6
                ElseIf dd = 29 Then
                i = 6
                ElseIf dd = 30 Then
                i = 6
                ElseIf dd = 31 Then
                i = 7
                
              End If
              
             
               MSChart1.Row = i
               If dd Mod 5 = 0 Then
                MSChart1.Column = 5
                Else
                 MSChart1.Column = dd Mod 5
               End If
               
               mapple = Left(Format(apple, "00000000"), 3)
               mstring = mapple
               mapple = Mid(Format(apple, "00000000"), 4)
               mstring = mstring + "." + mapple
               
               MSChart1.Data = Val(mstring) * 10000
               
               
               dd = 0
               End If
               
        Rs.MoveNext
        Loop
        End If
   End If
End If

         
    
   
End Sub

Private Sub ListView1_DblClick() 'listview 더블클릭시 수정
  Data_search.Show
   Data_search.Usercode (ListView1.SelectedItem.SubItems(1))
    Data_search.Text1 = ListView1.SelectedItem
    Data_search.Text2 = ListView1.SelectedItem.SubItems(1) '주민번호
    Data_search.Text3 = ListView1.SelectedItem.SubItems(2)
    Data_search.Combo1 = ListView1.SelectedItem.SubItems(8)
    
    
    

    
End Sub


Private Sub Location_add_Click()

End Sub

Private Sub location_setting_Click()
LocationSetting.Show

End Sub



Private Sub MSComm1_OnComm()
Dim i
Dim FS, FileStream, OutStream
Dim count_flag
If MSComm1.CommEvent = comEvReceive Then
    SaveSetting "COMM", "PORT", "SETTING", MSComm1.CommPort
    mbuf1 = mbuf1 + MSComm1.Input
    'Sleep 3000
    Timer1.Interval = 400
    If Left(commCmd, 1) = "S" Then
        If Len(mbuf1) = 26 Then
            For i = 1 To Len(mbuf1)
                If Not Asc(Mid(mbuf1, i, 1)) = 9 Then
                    tmp = tmp + Mid(mbuf1, i, 1)
                End If
            Next i
            sql = "select * from data where 시리얼번호 = '" & tmp & "'"
            If Not Run(sql) Then Exit Sub

            If Rs.BOF Then
                worker_add.Show
                worker_add.clover_serial.Text = tmp

            Else

                    Auto_worker_enter.Show
                    Auto_worker_enter.Wiorker_name = Trim(Rs.Fields("이름"))
                    Auto_worker_enter.Clover_sn.Text = Rs.Fields("시리얼번호")
                    Auto_worker_enter.Worker_id = Trim(Rs.Fields("주민등록번호"))
                    Auto_connect_start.Show
                    Auto_connect_start.Hide

            End If

            tmp = ""
            mbuf1 = ""
        End If
        
        ElseIf Left(commCmd, 1) = "X" Then
        If Len(mbuf1) = 3 Then
       For i = 1 To Len(mbuf1)
                If Not Asc(Mid(mbuf1, i, 1)) = 9 Then
                    tmp = tmp + Mid(mbuf1, i, 1)
                End If
            Next i
            If tmp = "00" Then
                If MsgBox("장비 리셋실패 재시도 하겠습니까?", vbYesNo, "알림") = vbYes Then
                    commCmd = "X"
                    MSComm1.Output = commCmd
                End If
            End If

            tmp = ""
            mbuf1 = ""
        End If
        
        End If
    
  
 
End If
Exit Sub

End Sub

Private Sub Password_Click()
    PassMod.Show
End Sub

Private Sub print_click_Click()
Dim conStr As String
    Dim DBPath As String
    Dim DBFile As String
    Dim MsgMake As String
    Dim strSql As String
    Dim i As Integer
    'Form_Load
    Call mRefresh
    Call DB_Conn_MDB2

  Load Data
        

    '접속 연결 문자열
    Set Con = New ADODB.Connection

    '파일의 경로
'    DBPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "MDB\"
'    '파일의 이름
'    DBFile = "Po.MDB"
'    'DB 구조 불러옴
'    Call DB_Info
'    '연결 문자열 설정
'    conStr = "Provider=Microsoft.Jet.OLEDB.4.0;"
'    conStr = conStr + "Data Source=" + DBPath + DBFile + ";"
'    conStr = conStr + "User ID=admin"
'    conStr = conStr + ";Jet OLEDB:Database Password=1234;"
'    Debug.Print conStr
Rs.MoveFirst

Data.Commands("Command1").CommandText = sql
DataReport1.Sections("section4").Controls("workerno").Caption = Trim(Text1.Text)
DataReport1.Sections("section4").Controls("datenow").Caption = Now
DataReport1.Show

        With Data
        If .rsCommand1.State <> 0 Then .rsCommand1.Close
        End With
'Set Data = Nothing


End Sub

Private Sub state_absence_click_Click()
On Error Resume Next
sql = "update data set 상태='결근' where 주민등록번호 = '" & ListView1.SelectedItem.SubItems(1) & "'"
If Not Run(sql) Then Exit Sub
MsgBox "변경되었습니다.", vbOKOnly, "알림"
Form_Load

End Sub

Private Sub reset_Click()

If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
Else
    MSComm1.PortOpen = False
    MSComm1.PortOpen = True
End If

If MsgBox("장비를 리셋 하겠습니까?", vbYesNo, "알림") = vbYes Then
    commCmd = "X"
    MSComm1.Output = commCmd

                
End If



End Sub

Private Sub state_businesstrip_click_Click()
On Error Resume Next
sql = "update data set 상태='출장' where 주민등록번호 = '" & ListView1.SelectedItem.SubItems(1) & "'"
If Not Run(sql) Then Exit Sub
MsgBox "변경되었습니다.", vbOKOnly, "알림"
Form_Load

End Sub

Private Sub state_normal_click_Click()
On Error Resume Next
sql = "update data set 상태='정상' where 주민등록번호 = '" & ListView1.SelectedItem.SubItems(1) & "'"
If Not Run(sql) Then Exit Sub
MsgBox "변경되었습니다.", vbOKOnly, "알림"
Form_Load
End Sub

Private Sub state_trouble_click_Click()
On Error Resume Next
sql = "update data set 상태='장비고장' where 주민등록번호 = '" & ListView1.SelectedItem.SubItems(1) & "'"
If Not Run(sql) Then Exit Sub
MsgBox "변경되었습니다.", vbOKOnly, "알림"
Form_Load

End Sub

Private Sub state_vacation_click_Click()
On Error Resume Next
sql = "update data set 상태='휴가' where 주민등록번호 = '" & ListView1.SelectedItem.SubItems(1) & "'"
If Not Run(sql) Then Exit Sub
MsgBox "변경되었습니다.", vbOKOnly, "알림"
Form_Load

End Sub

Private Sub Timer1_Timer()
Dim FS, FileStream, OutStream
On Error GoTo err:
Timer1.Interval = 0
If Left(commCmd, 1) = "T" Then
        If Right(mbuf1, 1) = "R" Then
            'count_flag = count_flag + 1
             Debug.Print mbuf1
            'If count_flag = 3 Then
            
                  For i = 1 To Len(mbuf1)
                
                
                 
                  
                      If Not Asc(Mid(mbuf1, i, 1)) = 9 Then
                          tmp = tmp + Mid(mbuf1, i, 1)
                      End If
                  Next i
                  'Auto_connect_start.Text1.Text = tmp
                  
                  
                  Set FS = CreateObject("Scripting.FileSystemObject")
                  Set FileStream = FS.OpenTextFile("tmpData.txt", 2, True)  '파일을 저장 용도로 엽니다.
                  FileStream.Write tmp  '파일에 Text1의 내용을 기록합니다.
                  FileStream.Close    '파일을 닫습니다.
                  Set FS = Nothing
                  
                  
                  Call Auto_worker_enter.Txtinsert
                   mbuf1 = ""
                   tmp = ""
                    mFlag_timer = 2
ElseIf Left(commCmd, 1) = "X" Then

       
         
      


            End If
            

           
           
        End If
   
Exit Sub

err:

End Sub

Private Sub Timer2_Timer()
    Dim apple
    
    apple = time Mod 7
    
    If apple = 1 Then
        Image1.Visible = True
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = False
        Image5.Visible = False
        Image6.Visible = False
        Image7.Visible = False
        
    ElseIf apple = 2 Then
        Image1.Visible = False
        Image2.Visible = True
        Image3.Visible = False
        Image4.Visible = False
        Image5.Visible = False
        Image6.Visible = False
        Image7.Visible = False
    ElseIf apple = 3 Then
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = True
        Image4.Visible = False
        Image5.Visible = False
        Image6.Visible = False
        Image7.Visible = False
    ElseIf apple = 4 Then
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = True
        Image5.Visible = False
        Image6.Visible = False
        Image7.Visible = False
    ElseIf apple = 5 Then
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = False
        Image5.Visible = True
        Image6.Visible = False
        Image7.Visible = False
    ElseIf apple = 6 Then
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = False
        Image5.Visible = False
        Image6.Visible = True
        Image7.Visible = False
    ElseIf apple = 0 Then
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = False
        Image5.Visible = False
        Image6.Visible = False
        Image7.Visible = True
    End If
    time = time + 1
    
    
End Sub

Private Sub worker_add_click_Click()
Call DB_Conn_MDB2

worker_add.Show


End Sub

Private Sub worker_edit_click_Click()
worker_edit.Show


End Sub








