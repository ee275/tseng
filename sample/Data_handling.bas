Attribute VB_Name = "Data_handling"

 Public Sub mDb_handling(mText, s_Num)
 'Dim mtext As String '�ؽ�Ʈ �о �����ϴ� �κ� ��ü ������
 Dim mFlag As String '������ ������ C,D,E �з��ϴ� Flag
 Dim mData As String '������ ��ü ������ (�ѱ��ھ� �ɰ������)
 Dim mFirstData As String '���� �����Ǵ� ������ ��¥+�ð�+��������
 Dim mDay As String '���� �����Ϳ��� ��¥ + �ð��κи� ����
 Dim time As Integer 'mData �ɰ���� �ѱ��ھ� �����ϱ����� �ܼ� ���� ���� (for)
 Dim mNuclear As String '���� �����Ϳ��� ���������� ����
 Dim sql As String 'sql ���� ������
 Dim mYear As String '��¥�ð� �κ��� ������ ��¥�κ�(��¥�������� �ٲ܋� ���)
 Dim mTime As String '��¥�ð� �κ��� ������ �ð��κ�(�ð��������� �ٲ܋� ���)
 Dim mYyear As String '������ myear �� ���˺������� �����ϴ� ���� format(mYear)
 Dim mYtime As String '������ mTime �� ���˺������� �����ϴ� ���� format(mTime)
 Dim apple As String '��ü ������ �����Ҷ� �ѱ��� �־�� �ܼ� ����
 Dim mUsercode As String '�ֹι�ȣ ��Ƶ� ����
 Dim mName As String '���� ��Ƶ� ����
 Dim mSerial_Num As String '�ø����ȣ ��Ƶ� ����
 Dim mValue_Sql As String '�ø����ȣ ������ �˻��ϴ� sql
 Dim total_sql As String
 Dim Nuclearchange As String
 
 
 
 


'�ؽ�Ʈ���� ��� ���ڸ� �ϳ��� �޴°�
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
            ',���� ������
            If mData = "," Then
            
            
             time = 0
                'Text1.Text = Text1.Text + vbCrLf
                    If mFirstData <> "" Then
                    'mNuclear�� mDay�� ������ �������� mDay �����͸� �޾� ��¥�� �ð� ���·� �ٽ� ����
                    
                    For j = 1 To Len(mDay)
                     apple = Mid(mDay, j, 1)
                       If j = 1 Then
                            mYear = "20" + apple '2015,2016,2017 ��� 20 ���̴°�
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
                    '���� ��
                    
                    '�ø����ȣ�� data ���̺� �� �˻����� ����, �ֹι�ȣ,�ø����ȣ�� ������ ����
                    mSerial_Num = s_Num
                    mValue_Sql = "select * from data where �ø����ȣ ='" & mSerial_Num & "'"
                    
                    If Not Run(mValue_Sql) Then Exit Sub
                      
                        
                        If Rs.EOF = True Then
                             MsgBox "DB�� ����� �ø��� ��ȣ�� �������� �ʽ��ϴ� ����� ����� �̿����ּ���", vbOKOnly, "�˸�"
                             Exit Sub
                        Else
                            Do While Not Rs.EOF

                                    mName = Trim(Rs.Fields("�̸�"))
                                    mUsercode = Trim(Rs.Fields("�ֹε�Ϲ�ȣ"))
                                    

                                Rs.MoveNext
                            Loop

                        End If
                        

                    '������ db�� �Է� �κ�
                    mYyear = Format(mYear, "yyyy-mm-dd")
                    mYtime = Format(mTime, "AMPM hh:mm:ss")
                     Nuclearchange = Format(mNuclear / 10, "######0.0")
                        sql = "inseRt into datac(�ð�,��������,����,�ø����ȣ,�ֹε�Ϲ�ȣ) values('" & mYyear + " " + mYtime & "','" & Nuclearchange & "','" & mName & "','" & mSerial_Num & "','" & mUsercode & "')"
                            If Not Run(sql) Then Exit Sub
                    End If
                    
                '�ڷ�� �ʱ�ȭ
                
                mDay = ""
                mNuclear = ""
                mFirstData = ""
                mYear = ""
                mTime = ""
            
            

            'C�� R�� �ؽ�Ʈ �ڽ��� ǥ�� ���Ѵ�
            
            ElseIf mData = "C" Then
                'Text1.Text = Text1.Text
            ElseIf mData = "R" Then
                'Text1.Text = Text1.Text
            
            'C �� R �� �ƴѰ��� �ؽ�Ʈ�� �ѷ��ְ� ������ �־ , �� ������������ �����Ѵ�
            Else
                'Text1.Text = Text1.Text + mData
                mFirstData = mFirstData + mData
                    '�ڸ����� ���ؼ� �ð��� ���� �������� �з��Ѵ�
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
                    mValue_Sql = "select * from data where �ø����ȣ ='" & mSerial_Num & "'"
                    
                    If Not Run(mValue_Sql) Then Exit Sub
                      
                        
                        If Rs.EOF = True Then
                             MsgBox "DB�� ����� �ø��� ��ȣ�� �������� �ʽ��ϴ� ����� ����� �̿����ּ���", vbOKOnly, "�˸�"
                             Exit Sub
                        Else
                            Do While Not Rs.EOF

                                    mName = Trim(Rs.Fields("�̸�"))
                                    mUsercode = Trim(Rs.Fields("�ֹε�Ϲ�ȣ"))
                                    

                                Rs.MoveNext
                            Loop

                        End If
                        
                    
                    mYyear = Format(mYear, "yyyy-mm-dd")
                    mYtime = Format(mTime, "hh:mm:ss AMPM")
                    Nuclearchange = Format(mNuclear / 10, "######0.0")
                    
                        sql = "inseRt into datad(�ð�,��������,����,�ø����ȣ,�ֹε�Ϲ�ȣ) values('" & mYyear + " " + mYtime & "','" & Nuclearchange & "','" & mName & "','" & mSerial_Num & "','" & mUsercode & "')"
                            If Not Run(sql) Then Exit Sub
                    End If
                    
                       ' sql = "select * from dataD where �ø����ȣ = '" & mSerial & "' AND �ֹε�Ϲ�ȣ= '" & userid & "'"
                    total_sql = "select * from monthly where �ֹι�ȣ = '" & mUsercode & "' AND ����='" & mYyear & "'"
                    If Not Run(total_sql) Then Exit Sub
                 
                    If Rs.EOF = True Then
                        total_sql = "inseRt into monthly (�ֹι�ȣ,����,������) values('" & mUsercode & "','" & mYyear & "','" & mNuclear & "') "
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
                    mValue_Sql = "select * from data where �ø����ȣ ='" & mSerial_Num & "'"
                    
                    If Not Run(mValue_Sql) Then Exit Sub
                        
                        
                        If Rs.EOF = True Then
                             MsgBox "DB�� ����� �ø��� ��ȣ�� �������� �ʽ��ϴ� ����� ����� �̿����ּ���", vbOKOnly, "�˸�"
                             Exit Sub
                        Else
                            Do While Not Rs.EOF

                                    mName = Trim(Rs.Fields("�̸�"))
                                    mUsercode = Trim(Rs.Fields("�ֹε�Ϲ�ȣ"))
                                    

                                Rs.MoveNext
                            Loop

                        End If
                        
                    
                    mYyear = Format(mYear, "yyyy-mm-dd")
                    mYtime = Format(mTime, "hh:mm:ss AMPM")
                     Nuclearchange = Format(mNuclear / 10, "######0.0")
                        sql = "inseRt into datae(�ð�,������,����,�ø����ȣ,�ֹε�Ϲ�ȣ) values('" & mYyear + " " + mYtime & "','" & Nuclearchange & "','" & mName & "','" & mSerial_Num & "','" & mUsercode & "')"
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


