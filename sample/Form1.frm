VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ÇÁ·Î±×·¥"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   5
      Left            =   4200
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   4
      Left            =   3480
      TabIndex        =   17
      Text            =   "Text4"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   3
      Left            =   2760
      TabIndex        =   16
      Text            =   "Text4"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   2
      Left            =   2040
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   1
      Left            =   1320
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   0
      Left            =   600
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   10920
      TabIndex        =   12
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   10920
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4560
      Width           =   12015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   12960
      TabIndex        =   9
      Text            =   "T150421132028"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   8760
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "¼ýÀÚ"
      Height          =   495
      Index           =   3
      Left            =   120
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   7
      Top             =   1575
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "¿µ¹®"
      Height          =   495
      Index           =   2
      Left            =   120
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ÇÑ±Û"
      Height          =   495
      Index           =   1
      Left            =   120
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ÀüÃ¼"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Á¾·á"
      Height          =   430
      Index           =   3
      Left            =   3840
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»èÁ¦"
      Height          =   430
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¼öÁ¤"
      Height          =   430
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ãß°¡"
      Height          =   430
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   3120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim i As Integer
Dim mbuf1

Private Sub DBCall()

    sql = "select * from data"
    If Option1(0).Value Then
        sql = sql
    ElseIf Option1(1).Value Then  'ÇÑ±Û
        sql = sql + "Where left(PB_Name,1)>='¤¡' And left(PB_Name,1)<='ÆR'"
    ElseIf Option1(2).Value Then  '¿µ¹®
        sql = sql + "Where left(PB_Name,1)>='A' And left(PB_Name,1)<='z'"
    ElseIf Option1(3).Value Then  '¼ýÀÚ
        sql = sql + "Where left(PB_Name,1)>='0' And left(PB_Name,1)<='9'"
    End If
    If Not Run(sql) Then Exit Sub
    With ListView1
        .ListItems.Clear
        For i = 0 To Rs.Fields.Count - 1
            If i = 0 Then
                .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 1600
            Else
                .ColumnHeaders.add i + 1, , Rs.Fields(i).Name, 1000, 2
            End If
        Next
        Call DBref
    End With
End Sub

Private Sub DBref()
    With ListView1
        .ListItems.Clear
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveFirst
            
            Do While Not Rs.EOF
                .ListItems.add , , Trim(Rs(0))
                For i = 1 To Rs.Fields.Count - 1
                    .ListItems(.ListItems.Count).SubItems(i) = IIf(IsNull(Rs(i).Value), "", Trim(Rs(i).Value))
                Next
            Rs.MoveNext
            Loop
        End If
    End With
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            sql = "inseRt into data(ÀÌ¸§,ÁÖ¹Îµî·Ï¹øÈ£,½Ã¸®¾ó¹øÈ£,Á¢¼Ó½Ã°£,»óÅÂ,ÀÌ¹ø´Þ´©Àû¼±·®) values ("
            sql = sql + "'" & txt(0).Text & "', "
            sql = sql + "'" & txt(1).Text & "', "
            sql = sql + "'" & txt(2).Text & "', "
            sql = sql + "'" & txt(3).Text & "', "
            sql = sql + "'" & txt(4).Text & "', "
            sql = sql + "'" & txt(5).Text & "')"
            
            Run (sql)
            
            MsgBox "ÀÔ·Â¿Ï·á"
            
            sql = "select * from data"
            Run (sql)
            Call DBCall
        Case 1
            For i = 0 To 5
                sql = "update data set " & txt(i).Tag & "='" & txt(i).Text & "' where pb_idx = "
                sql = sql & ListView1.SelectedItem.Text
                Run sql
            Next
            MsgBox "¼öÁ¤¿Ï·á"
            sql = "select * from data"
            Run (sql)
            Call DBCall
        Case 2
           If MsgBox("Á¤¸» »èÁ¦ÇÔ?", vbYesNo) = vbYes Then
                sql = "Delete From data where ÁÖ¹Îµî·Ï¹øÈ£ =" & ListView1.SelectedItem.Text
                Run sql
           End If
           sql = "select * from data"
           Run (sql)
           Call DBCall
        Case 3
            Unload Me
    End Select
End Sub

Private Sub Command2_Click()
Text2.Text = ""
MSComm1.Output = CStr(Trim(Text1.Text))

End Sub

Private Sub Command3_Click()
Text2.Text = ""
If MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
End If
End Sub

Private Sub Form_Load()
    MSComm1.CommPort = GetSetting("COMM", "PORT", "SETTING")
    Call DBCall
    'MSComm1.PortOpen = True
End Sub

Private Sub ListView1_Click()
    sql = "select * from data"
    Run sql
    If Not (Rs.BOF Or Rs.EOF) Then
         txt(0).Text = Trim(Rs("ÀÌ¸§"))
         txt(1).Text = Trim(Rs("ÁÖ¹Îµî·Ï¹øÈ£"))
         
    End If
End Sub

Private Sub MSComm1_OnComm()
Dim i

If MSComm1.CommEvent = comEvReceive Then
    SaveSetting "COMM", "PORT", "SETTING", MSComm1.CommPort
    mbuf1 = mbuf1 + MSComm1.Input
    
    
    
    If Len(mbuf1) > 21 Then
        For i = 1 To Len(mbuf1)
            If Not Asc(Mid(mbuf1, i, 1)) = 9 Then
                Text2.Text = Text2.Text + Mid(mbuf1, i, 1)
            End If
        Next i
        'Text2.Text = mbuf1
        Text3.Text = Len(mbuf1)
        mbuf1 = ""
    End If
    
    
End If
Exit Sub

End Sub

Private Sub Option1_Click(Index As Integer)
    Call DBCall
End Sub


Private Sub Timer1_Timer()
On Error GoTo err:
Dim i


If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
    'MSComm1.Output = "S"
    MSComm1.Output = CStr(Trim(Text1.Text))
Else

End If

Exit Sub

err:

End Sub
