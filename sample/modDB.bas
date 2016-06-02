Attribute VB_Name = "modDB"
Option Explicit

'MDB 연결문자
'Provider=Microsoft.Jet.OLEDB.4.0;
'Data Source=D:\VisualBasic Data\아르바이트\양호상\cmt.mdb;
'Persist Security Info=False
'MSSQL 연결문자
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=192.168.200.6
Public Con As ADODB.Connection
Public Rs As ADODB.Recordset
Public mRs As ADODB.Recordset
Public mmRs As ADODB.Recordset
Public MDB_Make As ADOX.Catalog  'MDB 파일을 만들기 위한 참조
Public mFlag_timer As Integer

Type DB_Type
    DB_Table As String
    DB_Field As String
End Type

Dim sql As String
Public pas As String
Public PortNo As String
Public DB_Gujo(100) As DB_Type
Public commCmd As String
Public admin As Boolean
Public FirstExe As Boolean
'Dim Pass


' 레지스트리 보안 옵션...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 레지스트리 키 ROOT 형식...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode null 종료 문자열
Const REG_DWORD = 4                      ' 32비트 숫자

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'디비를 번호별로 추가가 가능하게끔 설정하는 부분
Public Sub DB_Info()
    ' 1번 브랜드 관리 테이블
    DB_Gujo(1).DB_Table = "Data"
    DB_Gujo(1).DB_Field = _
            " [이름] char(20) , " & _
            " [주민등록번호] Char(20) Identity Primary Key Not Null, " & _
            " [시리얼번호] Char(30), " & _
            " [접속시간] Char(30), " & _
            " [금일누적선량] Char(30), " & _
            " [이번달누적선량] Char(30), " & _
            " [단위] Char(30), " & _
            " [상태] Char(30), " & _
            " [작업장] Char(30), " & _
            " [등록일] Char(30) "
            
            
    DB_Gujo(2).DB_Table = "DataC"
    DB_Gujo(2).DB_Field = _
            " [시간] char(40), " & _
            " [누적선량] Char(20), " & _
            " [성명] Char(20), " & _
            " [시리얼번호] Char(20), " & _
            " [주민등록번호] Char(30) "
    DB_Gujo(3).DB_Table = "DataD"
    DB_Gujo(3).DB_Field = _
            " [시간] char(40), " & _
            " [누적선량] Char(20), " & _
            " [성명] Char(20), " & _
            " [시리얼번호] Char(20), " & _
            " [주민등록번호] Char(30) "
    DB_Gujo(4).DB_Table = "DataE"
    DB_Gujo(4).DB_Field = _
            " [시간] char(40), " & _
            " [선량율] Char(20), " & _
            " [성명] Char(20), " & _
            " [시리얼번호] Char(20), " & _
            " [주민등록번호] Char(30) "
    DB_Gujo(5).DB_Table = "Location"
    DB_Gujo(5).DB_Field = _
            " [작업장] char(40) "
            
    DB_Gujo(6).DB_Table = "Monthly"
    DB_Gujo(6).DB_Field = _
            " [주민번호] Char(15), " & _
            " [날자] Char(10), " & _
            " [누적량] char(20) "
            
'    ' 2번 대분류 관리 테이블
'    DB_Gujo(2).DB_Table = "bType"
'    DB_Gujo(2).DB_Field = _
'            " [bType_idx] Long Identity Primary Key Not Null, " & _
'            " [bType_Code] Char(20), " & _
'            " [bType_Name] Char(50)"
'    ' 3번 소분류 관리 테이블
'    DB_Gujo(3).DB_Table = "sType"
'    DB_Gujo(3).DB_Field = _
'            " [sType_idx] Long Identity Primary Key Not Null, " & _
'            " [sType_Code] Char(20), " & _
'            " [sType_Name] Char(50)"
'    ' 4번 구매/판매 테이블
'    DB_Gujo(4).DB_Table = "ItemsIm"
'    DB_Gujo(4).DB_Field = _
'            " [Im_idx] Long Identity Primary Key Not Null, " & _
'            " [Im_Times] Long, " & _
'            " [Im_Date] Char(10), " & _
'            " [Im_Brand] Long, " & _
'            " [Im_bType] Long, " & _
'            " [Im_sType] Long, " & _
'            " [Im_Sex] Char(1) Default 'F', " & _
'            " [Im_Image] Char(255), " & _
'            " [Im_iCount] Long Default 0, " & _
'            " [Im_Value1] Single Default 0, " & _
'            " [Im_Value2] Single Default 0 "
'    DB_Gujo(5).DB_Table = "ItemsEx"
'    DB_Gujo(5).DB_Field = _
'            " [Ex_idx] Long Identity Primary Key Not Null, " & _
'            " [Ex_Im] Long, " & _
'            " [Ex_Ex] Char(1) Default 'N', " & _
'            " [Ex_Times] Long, " & _
'            " [Ex_Date] Char(10), " & _
'            " [Ex_Brand] Long, " & _
'            " [Ex_bType] Long, " & _
'            " [Ex_vCount] Long Default 0, " & _
'            " [Ex_sType] Long, " & _
'            " [Ex_Sex] Char(1) Default 'F', " & _
'            " [Ex_Image] Char(255), " & _
'            " [Ex_iCount] Long Default 0, " & _
'            " [Ex_Value1] Single Default 0, " & _
'            " [Ex_Value2] Single Default 0, " & _
'            " [Ex_Numbering] Long Default 0, " & _
'            " [Ex_Code] Char(" + CStr(2 + 2 + 2 + 2 + 3 + 3 + 2 + 3) + ") Default ''"
'            '월/년/브랜드id/타입1id/타입2총수량/타입2순번/타입2id/회차별번호 총 19자리
End Sub

'DB_info 에서 저장한 번호를 index 값으로 주면 해당 테이블을 만듦.
Public Sub Create_Table(Index As Integer)
    Dim strSql As String
    '로그인 테이블 생성 (기본 1번 테이블)
    strSql = "CREATE Table " + DB_Gujo(Index).DB_Table + "(" + DB_Gujo(Index).DB_Field + ")"
    '쿼리문을 실행
    Call Con.Execute(strSql)
End Sub

Public Function DB_Conn_MDB() As Boolean
    Dim conStr As String
    Dim DBPath As String
    Dim DBFile As String
    Dim DBPathtxt As String
    Dim DBFiletxt As String
    Dim MsgMake As String
    Dim strSql As String
    Dim i As Integer
    Dim apple
    DB_Conn_MDB = False
    On Error GoTo ConnectError
    '접속 연결 문자열
    Set Con = New ADODB.Connection
    
    '파일의 경로
    DBPath = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    
    DBPathtxt = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    'IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    
    'IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    'apple = IIf(Right("C:\", 1) = ":\", "C:\", "C:\")
    '파일의 이름
    DBFile = "Data.MDB"
    DBFiletxt = "tmpData.txt"
    'DB 구조 불러옴
    Call DB_Info
    '연결 문자열 설정
    conStr = "Provider=Microsoft.Jet.OLEDB.4.0;"
    conStr = conStr + "Data Source=" + DBPath + DBFile + ";"
    conStr = conStr + "User ID=admin"
    conStr = conStr + ";Jet OLEDB:Database Password=1234;"
    Data.Connection1.ConnectionString = conStr
    Data2.Connection2.ConnectionString = conStr
    
    '파일이 있는지 검사
 
    If Not Dir(DBPathtxt + DBFiletxt) = "" Then
        Debug.Print "dd"
      Else
        Open "tmpData.txt" For Output As #1
        Print #1, ""
    Close #1

    End If
    
    If Not Dir(DBPath + DBFile) = "" Then
        '있으면 DB 연결
        Con.Open conStr
    Else
        '없으면 생성
        If Not MsgBox("DB 파일이 존재하지 않습니다." + vbCrLf _
            + "DB파일을 새로 만드시겠습니까?", vbQuestion + vbYesNo, "DB 생성 확인") = vbYes Then
            MsgBox "DB 파일이 없어서 작업을 계속 진행할 수 없습니다." + vbCrLf _
                + "DB 파일을 확인하신 후 다시 시도하십시오.", vbCritical + vbOKOnly, "DB 오류"
            Exit Function
        End If
        '디렉터리 검사
        If Dir(DBPath, vbDirectory) = "" Then
            '디렉터리 생성
            Call MkDir(DBPath)
        End If
        'DB 파일 생성
        Set MDB_Make = New ADOX.Catalog
        Call MDB_Make.Create(conStr)
        Set MDB_Make = Nothing
        'DB 파일에 연결
        Con.Open conStr
        'Table 생성 및 Field 생성
        For i = 1 To UBound(DB_Gujo)
            If Not Len(DB_Gujo(i).DB_Table) = 0 Then
                Call Create_Table(i)
            End If
        Next
    End If
    On Error GoTo 0
    DB_Conn_MDB = True
    Exit Function
ConnectError:
    DB_Conn_MDB = False
    
    MsgMake = "DB에 문제가 발생하였습니다." + vbCrLf + vbCrLf _
        + "DB : " + DBPath + DBFile + vbCrLf _
        + "Error : "
    For i = 0 To Con.Errors.Count - 1
        MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
    Next
    MsgMake = MsgMake + vbCrLf + vbCrLf _
        + "DB를 점검후에 다시 시도 하세요."
    MsgBox MsgMake, vbCritical + vbOKOnly, "프로그램 오류"
End Function

Public Function DB_Conn_MDB2() As Boolean
    Dim conStr As String
    Dim DBPath As String
    Dim DBFile As String
    Dim DBPathtxt As String
    Dim DBFiletxt As String
    Dim MsgMake As String
    Dim strSql As String
    Dim i As Integer
    
    DB_Conn_MDB2 = False
    On Error GoTo ConnectError
    '접속 연결 문자열
    Set Con = New ADODB.Connection
    
    '파일의 경로
    DBPath = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    'DBPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    DBPathtxt = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    'DBPathtxt = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    '파일의 이름
    DBFile = "Data.MDB"
    DBFiletxt = "tmpData.txt"
    'DB 구조 불러옴
    Call DB_Info
    '연결 문자열 설정
    conStr = "Provider=Microsoft.Jet.OLEDB.4.0;"
    conStr = conStr + "Data Source=" + DBPath + DBFile + ";"
    conStr = conStr + "User ID=admin"
    conStr = conStr + ";Jet OLEDB:Database Password=1234;"
    
   
        Con.Open conStr
    
    DB_Conn_MDB2 = True
    Exit Function
ConnectError:
    DB_Conn_MDB2 = False
    
    MsgMake = "DB에 문제가 발생하였습니다." + vbCrLf + vbCrLf _
        + "DB : " + DBPath + DBFile + vbCrLf _
        + "Error : "
    For i = 0 To Con.Errors.Count - 1
        MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
    Next
    MsgMake = MsgMake + vbCrLf + vbCrLf _
        + "DB를 점검후에 다시 시도 하세요."
    MsgBox MsgMake, vbCritical + vbOKOnly, "프로그램 오류"
End Function
Public Function mRun(nSql As String, Optional Message As Boolean) As Boolean
    Dim MsgMake As String
    Dim i As Integer
    
    On Error GoTo ErrSql
    Set mRs = Con.Execute(nSql)
    On Error GoTo 0
    mRun = True
    Exit Function
ErrSql:
    mRun = False
    If Message Then
        
        MsgMake = "DB에 문제가 발생하였습니다." + vbCrLf + vbCrLf _
            + "Sql : " _
            + nSql + vbCrLf _
            + "Error : "
        For i = 0 To Con.Errors.Count - 1
            MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
        Next
        MsgMake = MsgMake + vbCrLf + vbCrLf _
            + "DB를 점검후에 다시 시도 하세요."
        MsgBox MsgMake, vbCritical + vbOKOnly, "프로그램 오류"
    End If
    MsgBox err.Description
End Function
Public Function mmRun(nSql As String, Optional Message As Boolean) As Boolean
    Dim MsgMake As String
    Dim i As Integer
    
    On Error GoTo ErrSql
    Set mRs = Con.Execute(nSql)
    On Error GoTo 0
    mmRun = True
    Exit Function
ErrSql:
    mmRun = False
    If Message Then
        
        MsgMake = "DB에 문제가 발생하였습니다." + vbCrLf + vbCrLf _
            + "Sql : " _
            + nSql + vbCrLf _
            + "Error : "
        For i = 0 To Con.Errors.Count - 1
            MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
        Next
        MsgMake = MsgMake + vbCrLf + vbCrLf _
            + "DB를 점검후에 다시 시도 하세요."
        MsgBox MsgMake, vbCritical + vbOKOnly, "프로그램 오류"
    End If
    MsgBox err.Description
End Function
Public Function Run(nSql As String, Optional Message As Boolean) As Boolean
    Dim MsgMake As String
    Dim i As Integer
    
    On Error GoTo ErrSql
    Set Rs = Con.Execute(nSql)
    On Error GoTo 0
    Run = True
    Exit Function
ErrSql:
    Run = False
    If Message Then
        
        MsgMake = "DB에 문제가 발생하였습니다." + vbCrLf + vbCrLf _
            + "Sql : " _
            + nSql + vbCrLf _
            + "Error : "
        For i = 0 To Con.Errors.Count - 1
            MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
        Next
        MsgMake = MsgMake + vbCrLf + vbCrLf _
            + "DB를 점검후에 다시 시도 하세요."
        MsgBox MsgMake, vbCritical + vbOKOnly, "프로그램 오류"
    End If
    MsgBox err.Description
End Function

Public Function DBField_Type(Index As Integer) As String
    Select Case Index
        Case adArray
            DBField_Type = "adArray"
        Case adBigInt
            DBField_Type = "adBigInt"
        Case adBinary
            DBField_Type = "adBinary"
        Case adBoolean
            DBField_Type = "adBoolean"
        Case adBSTR
            DBField_Type = "adBSTR"
        Case adChapter
            DBField_Type = "adChapter"
        Case adCurrency
            DBField_Type = "adCurrency"
        Case adChar
            DBField_Type = "adChar"
        Case adDate
            DBField_Type = "adDate"
        Case adDBDate
            DBField_Type = "adDBDate"
        Case adDBTime
            DBField_Type = "adDBTime"
        Case adDBTimeStamp
            DBField_Type = "adDBTimeStamp"
        Case adDecimal
            DBField_Type = "adDecimal"
        Case adDouble
            DBField_Type = "adDouble"
        Case adEmpty
            DBField_Type = "adEmpty"
        Case adError
            DBField_Type = "adError"
        Case adFileTime
            DBField_Type = "adFileTime"
        Case adGUID
            DBField_Type = "adGUID"
        Case adIDispatch
            DBField_Type = "adIDispatch"
        Case adInteger
            DBField_Type = "adInteger"
        Case adIUnknown
            DBField_Type = "adIUnknown"
        Case adLongVarBinary
            DBField_Type = "adLongVarBinary"
        Case adLongVarChar
            DBField_Type = "adLongVarChar"
        Case adLongVarWChar
            DBField_Type = "adLongVarWChar"
        Case adNumeric
            DBField_Type = "adNumeric"
        Case adPropVariant
            DBField_Type = "adPropVariant"
        Case adSingle
            DBField_Type = "adSingle"
        Case adSmallInt
            DBField_Type = "adSmallInt"
        Case adTinyInt
            DBField_Type = "adTinyInt"
        Case adUnsignedBigInt
            DBField_Type = "adUnsignedBigInt"
        Case adUnsignedInt
            DBField_Type = "adUnsignedInt"
        Case adUserDefined
            DBField_Type = "adUserDefined"
        Case adVarBinary
            DBField_Type = "adVarBinary"
        Case adVarChar
            DBField_Type = "adVarChar"
        Case adVariant
            DBField_Type = "adVariant"
        Case adVarNumeric
            DBField_Type = "adVarNumeric"
        Case adVarWChar
            DBField_Type = "adVarWChar"
        Case adWChar
            DBField_Type = "adWChar"
        Case Else
            DBField_Type = "알수없음."
    End Select
End Function
