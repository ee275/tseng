VERSION 5.00
Begin VB.Form Comm 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "통신설정"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "취소"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Text            =   "선택해주세요"
      Top             =   540
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "포트 :"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "Comm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum PortAttr
  PortFree = 0
  PortInUse = 1
  PortUnknown = 2
End Enum
Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Dim apple
Dim apple2

If Index = 0 Then
    Unload Me
ElseIf Index = 1 Then
    If Combo1.ListIndex = -1 Then
        MsgBox "포트를 선택해주세요.", vbOKOnly, "알림"
    Else
       
       
        apple2 = Len(Combo1.Text)
        If apple2 = 4 Then
             apple = Right(Combo1.Text, 1)
        ElseIf apple2 = 5 Then
             apple = Right(Combo1.Text, 2)
        ElseIf apple2 = 6 Then
             apple = Right(Combo1.Text, 3)
        End If
        
        SaveSetting "COMM", "PORT", "SETTING", apple
        PortNo = GetSetting("COMM", "PORT", "SETTING")
        MsgBox "통신설정이 " & Combo1.List(Combo1.ListIndex) & "로 변경되었습니다.", vbOKOnly, "알림"
        MainScreen.MSComm1.CommPort = PortNo
        Unload Me
    End If
    
End If
End Sub

Private Sub Form_Load()
On Error Resume Next


ShowPorts

End Sub
Private Sub ShowPorts()
  Dim intIndex As Integer
  Dim intPort As Integer
  Dim intFree As Integer
  On Error GoTo ErrorFound
  With MainScreen.MSComm1
    If .PortOpen Then .PortOpen = False
    intPort = .CommPort
    Combo1.Clear
    Combo1.AddItem "---- 활성화 ----"
    Combo1.ItemData(0) = -2 'not possible
    Combo1.AddItem ""
    Combo1.ItemData(1) = -2 'not possible
    intFree = 0
    For intIndex = 1 To 255
      Select Case CheckPort(intIndex)
        Case PortFree
          intFree = intFree + 1
          Combo1.AddItem "Com" & CStr(intIndex), intFree
          Combo1.ItemData(intFree) = intIndex
        Case PortInUse
          Combo1.AddItem "Com" & CStr(intIndex)
      End Select
    Next intIndex
    If .PortOpen Then .PortOpen = False
    .CommPort = intPort
    If CheckPort(intPort) = PortFree Then
      If .PortOpen = False Then .PortOpen = True
    End If
  End With 'MSComm1
Exit Sub
ErrorFound:
  MsgBox err.Description, vbCritical, "Error " & CStr(err.Number)
  On Error GoTo 0
End Sub

Private Function CheckPort(intPort As Integer) As PortAttr
  On Error GoTo ErrorFound
  With MainScreen.MSComm1
    If .PortOpen Then .PortOpen = False
    .CommPort = intPort
    .PortOpen = True
    CheckPort = PortFree
    If .PortOpen = False Then .PortOpen = True
  End With 'MSComm1
Exit Function
ErrorFound:
  Select Case err.Number
    Case 8002 'port doesnt exist
      CheckPort = PortUnknown
    Case 8005 'port already in use
      CheckPort = PortInUse
  End Select
  On Error GoTo 0
End Function

