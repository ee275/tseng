Attribute VB_Name = "Data_handling"

 Public Sub mDb_handling(mText, s_Num)
 'Dim mtext As String '텍스트 읽어서 저장하는 부분 전체 데이터
 Dim mFlag As String '들어오는 데이터 C,D,E 분류하는 Flag
 Dim mData As String '들어오는 전체 데이터 (한글자씩 쪼개논상태)
 Dim mFirstData As String '최초 가공되는 데이터 날짜+시간+누적선량
 Dim mDay As String '최초 데이터에서 날짜 + 시간부분만 띠어낸곳
 Dim time As Integer 'mData 쪼개논것 한글자씩 가공하기위해 단순 증가 변수 (for)
 Dim mNuclear As String '최초 데이터에서 누적선량만 띠어낸곳
 Dim sql As String 'sql 쿼리 문저장
 Dim mYear As String '날짜시간 부분을 나눌때 날짜부분(날짜형식으로 바꿀떄 사용)
 Dim mTime As String '날짜시간 부분을 나눌때 시간부분(시간형식으로 바꿀떄 사용)
 Dim mYyear As String '가공한 myear 를 포맷변형시켜 저장하는 변수 format(mYear)
 Dim mYtime As String '가공한 mTime 를 포맷변형시켜 저장하는 변수 format(mTime)
 Dim apple As String '전체 데이터 가공할때 한글자 넣어둔 단순 변수
 Dim mUsercode As String '주민번호 담아둘 변수
 Dim mName As String '성명 담아둘 변수
 Dim mSerial_Num As String '시리얼번호 담아둘 변수
 Dim mValue_Sql As String '시리얼번호 가지고 검색하는 sql
 Dim total_sql As String
 Dim Nuclearchange As String
 
 
 
 


'텍스트파일 열어서 문자를 하나씩 받는곳
For i = 1 To Len(mText)
        mData = Mid(mText, i, 1)
        
        If mData = "C" Then
            mFlag = mData
            
        ElseIf mData = "D" Then
            mFlag = mData
        
        ElseIf mData = "E" Then
            mFlag = mData
       
        End If


        If mFlag = "C" Then
            ',으로 나눈다
            If mData = "," Then
            
            
             time = 0
                'Text1.Text = Text1.Text + vbCrLf
                    If mFirstData <> "" Then
                    'mNuclear와 mDay로 가공한 데이터중 mDay 데이터를 받아 날짜와 시간 형태로 다시 가공
                    
                    For j = 1 To Len(mDay)
                     apple = Mid(mDay, j, 1)
                       If j = 1 Then
                            mYear = "20" + apple '2015,2016,2017 등등 20 붙이는곳
                       ElseIf j < 7 Then
                            mYear = mYear + apple 'ex)20140604
                       ElseIf j < 13 Then
                            mTime = mTime + apple 'ex)132028
'
                       End If
                       
                       If j = 2 Then
                            mYear = mYear + "-" 'ex)2014-0-604
                       ElseIf j = 4 Then
                            mYear = mYear + "-"
                       End If
                       
                       If j = 8 Then
                            mTime = mTime + ":" 'ex)13:20:28
                       ElseIf j = 10 Then
                            mTime = mTime + ":"
                       
                       End If
                       
                    Next j
                    '가공 끝
                    
                    '시리얼번호를 data 테이블 에 검색한후 성명, 주민번호,시리얼번호를 변수에 저장
                    mSerial_Num = s_Num
                    mValue_Sql = "select * from data where 시리얼번호 ='" & mSerial_Num & "'"
                    
                    If Not Run(mValue_Sql) Then Exit Sub
                      
                        
                        If Rs.EOF = True Then
                             MsgBox "DB에 저장된 시리얼 번호가 존재하지 않습니다 사용자 등록후 이용해주세요", vbOKOnly, "알림"
                             Exit Sub
                        Else
                            Do While Not Rs.EOF

                                    mName = Trim(Rs.Fields("이름"))
                                    mUsercode = Trim(Rs.Fields("주민등록번호"))
                                    

                                Rs.MoveNext
                            Loop

                        End If
                        

                    '데이터 db에 입력 부분
                    mYyear = Format(mYear, "yyyy-mm-dd")
                    mYtime = Format(mTime, "AMPM hh:mm:ss")
                     Nuclearchange = Format(mNuclear / 10, "######0.0")
                        sql = "inseRt into datac(시간,누적선량,성명,시리얼번호,주민등록번호) values('" & mYyear + " " + mYtime & "','" & Nuclearchange & "','" & mName & "','" & mSerial_Num & "','" & mUsercode & "')"
                            If Not Run(sql) Then Exit Sub
                    End If
                    
                '자료들 초기화
                
                mDay = ""
                mNuclear = ""
                mFirstData = ""
                mYear = ""
                mTime = ""
            
            

            'C나 R은 텍스트 박스에 표시 안한다
            
            ElseIf mData = "C" Then
                'Text1.Text = Text1.Text
            ElseIf mData = "R" Then
                'Text1.Text = Text1.Text
            
            'C 나 R 이 아닌것은 텍스트에 뿌려주고 변수에 넣어서 , 가 나오기전까지 조합한다
            Else
                'Text1.Text = Text1.Text + mData
                mFirstData = mFirstData + mData
                    '자리수에 의해서 시간과 누적 선량으로 분류한다
                    If time < 12 Then
                        mDay = mDay + mData
                    ElseIf time > 11 Then
                        mNuclear = mNuclear + mData
                    End If
                                
            End If
                
        ElseIf mFlag = "D" Then
        
            If mData = "," Then

             time = 0
               ' Text2.Text = Text2.Text + vbCrLf
                    If mFirstData <> "" Then
                    
                    
                    For j = 1 To Len(mDay)
                     apple = Mid(mDay, j, 1)
                       If j = 1 Then
                            mYear = "20" + apple
                       ElseIf j < 7 Then
                            mYear = mYear + apple
                       ElseIf j < 13 Then
                            mTime = mTime + apple
'
                       End If
                       
                       If j = 2 Then
                            mYear = mYear + "-"
                       ElseIf j = 4 Then
                            mYear = mYear + "-"
                       End If
                       
                       If j = 8 Then
                            mTime = mTime + ":"
                       ElseIf j = 10 Then
                            mTime = mTime + ":"
                       
                       End If
                       
                    Next j
                    mSerial_Num = s_Num
                    mValue_Sql = "select * from data where 시리얼번호 ='" & mSerial_Num & "'"
                    
                    If Not Run(mValue_Sql) Then Exit Sub
                      
                        
                        If Rs.EOF = True Then
                             MsgBox "DB에 저장된 시리얼 번호가 존재하지 않습니다 사용자 등록후 이용해주세요", vbOKOnly, "알림"
                             Exit Sub
                        Else
                            Do While Not Rs.EOF

                                    mName = Trim(Rs.Fields("이름"))
                                    mUsercode = Trim(Rs.Fields("주민등록번호"))
                                    

                                Rs.MoveNext
                            Loop

                        End If
                        
                    
                    mYyear = Format(mYear, "yyyy-mm-dd")
                    mYtime = Format(mTime, "hh:mm:ss AMPM")
                    Nuclearchange = Format(mNuclear / 10, "######0.0")
                    
                        sql = "inseRt into datad(시간,누적선량,성명,시리얼번호,주민등록번호) values('" & mYyear + " " + mYtime & "','" & Nuclearchange & "','" & mName & "','" & mSerial_Num & "','" & mUsercode & "')"
                            If Not Run(sql) Then Exit Sub
                    End If
                    
                       ' sql = "select * from dataD where 시리얼번호 = '" & mSerial & "' AND 주민등록번호= '" & userid & "'"
                    total_sql = "select * from monthly where 주민번호 = '" & mUsercode & "' AND 날자='" & mYyear & "'"
                    If Not Run(total_sql) Then Exit Sub
                 
                    If Rs.EOF = True Then
                        total_sql = "inseRt into monthly (주민번호,날자,누적량) values('" & mUsercode & "','" & mYyear & "','" & mNuclear & "') "
                        If Not Run(total_sql) Then Exit Sub
                        
                    Else
                    
                    End If
                    
                
                 mDay = ""
                mNuclear = ""
                mFirstData = ""
                mYear = ""
                mTime = ""
            
            
            
            ElseIf mData = "D" Then
                'Text2.Text = Text2.Text
            ElseIf mData = "R" Then
                'Text2.Text = Text2.Text
            Else
                'Text2.Text = Text2.Text + mData
                mFirstData = mFirstData + mData
                
                    If time < 12 Then
                        mDay = mDay + mData
                    ElseIf time > 11 Then
                        mNuclear = mNuclear + mData
                    End If
                                
            End If

        ElseIf mFlag = "E" Then
        
            If mData = "," Then
                time = 0
                'Text3.Text = Text3.Text + vbCrLf
                    If mFirstData <> "" Then
                    
                    
                    For j = 1 To Len(mDay)
                     apple = Mid(mDay, j, 1)
                       If j = 1 Then
                            mYear = "20" + apple
                       ElseIf j < 7 Then
                            mYear = mYear + apple
                       ElseIf j < 13 Then
                            mTime = mTime + apple
'
                       End If
                       
                       If j = 2 Then
                            mYear = mYear + "-"
                       ElseIf j = 4 Then
                            mYear = mYear + "-"
                       End If
                       
                       If j = 8 Then
                            mTime = mTime + ":"
                       ElseIf j = 10 Then
                            mTime = mTime + ":"
                       
                       End If
                       
                    Next j
                    mSerial_Num = s_Num
                    mValue_Sql = "select * from data where 시리얼번호 ='" & mSerial_Num & "'"
                    
                    If Not Run(mValue_Sql) Then Exit Sub
                        
                        
                        If Rs.EOF = True Then
                             MsgBox "DB에 저장된 시리얼 번호가 존재하지 않습니다 사용자 등록후 이용해주세요", vbOKOnly, "알림"
                             Exit Sub
                        Else
                            Do While Not Rs.EOF

                                    mName = Trim(Rs.Fields("이름"))
                                    mUsercode = Trim(Rs.Fields("주민등록번호"))
                                    

                                Rs.MoveNext
                            Loop

                        End If
                        
                    
                    mYyear = Format(mYear, "yyyy-mm-dd")
                    mYtime = Format(mTime, "hh:mm:ss AMPM")
                     Nuclearchange = Format(mNuclear / 10, "######0.0")
                        sql = "inseRt into datae(시간,선량율,성명,시리얼번호,주민등록번호) values('" & mYyear + " " + mYtime & "','" & Nuclearchange & "','" & mName & "','" & mSerial_Num & "','" & mUsercode & "')"
                            If Not Run(sql) Then Exit Sub
                    End If
                
                mDay = ""
                mNuclear = ""
                mFirstData = ""
                mYear = ""
                mTime = ""
            ElseIf mData = "E" Then
                'Text3.Text = Text3.Text
            ElseIf mData = "R" Then
                'Text3.Text = Text3.Text
            Else
                'Text3.Text = Text3.Text + mData
                mFirstData = mFirstData + mData
                
                    If time < 12 Then
                        mDay = mDay + mData
                    ElseIf time > 11 Then
                        mNuclear = mNuclear + mData
                    End If
                                
            End If

        End If

        
 If mData <> "R" Then
     time = time + 1
 
     If mData = "," Then
        time = 0
     ElseIf mData = "D" Then
        time = 0
     ElseIf mData = "E" Then
        time = 0
     ElseIf mData = "C" Then
        time = 0
     End If
 End If
 
Next i





 
    
End Sub


