Attribute VB_Name = "modDB"
Option Explicit

'MDB ���Ṯ��
'Provider=Microsoft.Jet.OLEDB.4.0;
'Data Source=D:\VisualBasic Data\�Ƹ�����Ʈ\��ȣ��\cmt.mdb;
'Persist Security Info=False
'MSSQL ���Ṯ��
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=192.168.200.6
Public Con As ADODB.Connection
Public Rs As ADODB.Recordset
Public mRs As ADODB.Recordset
Public mmRs As ADODB.Recordset
Public MDB_Make As ADOX.Catalog  'MDB ������ ����� ���� ����
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


' ������Ʈ�� ���� �ɼ�...
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
                     
' ������Ʈ�� Ű ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode null ���� ���ڿ�
Const REG_DWORD = 4                      ' 32��Ʈ ����

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'��� ��ȣ���� �߰��� �����ϰԲ� �����ϴ� �κ�
Public Sub DB_Info()
    ' 1�� �귣�� ���� ���̺�
    DB_Gujo(1).DB_Table = "Data"
    DB_Gujo(1).DB_Field = _
            " [�̸�] char(20) , " & _
            " [�ֹε�Ϲ�ȣ] Char(20) Identity Primary Key Not Null, " & _
            " [�ø����ȣ] Char(30), " & _
            " [���ӽð�] Char(30), " & _
            " [���ϴ�������] Char(30), " & _
            " [�̹��޴�������] Char(30), " & _
            " [����] Char(30), " & _
            " [����] Char(30), " & _
            " [�۾���] Char(30), " & _
            " [�����] Char(30) "
            
            
    DB_Gujo(2).DB_Table = "DataC"
    DB_Gujo(2).DB_Field = _
            " [�ð�] char(40), " & _
            " [��������] Char(20), " & _
            " [����] Char(20), " & _
            " [�ø����ȣ] Char(20), " & _
            " [�ֹε�Ϲ�ȣ] Char(30) "
    DB_Gujo(3).DB_Table = "DataD"
    DB_Gujo(3).DB_Field = _
            " [�ð�] char(40), " & _
            " [��������] Char(20), " & _
            " [����] Char(20), " & _
            " [�ø����ȣ] Char(20), " & _
            " [�ֹε�Ϲ�ȣ] Char(30) "
    DB_Gujo(4).DB_Table = "DataE"
    DB_Gujo(4).DB_Field = _
            " [�ð�] char(40), " & _
            " [������] Char(20), " & _
            " [����] Char(20), " & _
            " [�ø����ȣ] Char(20), " & _
            " [�ֹε�Ϲ�ȣ] Char(30) "
    DB_Gujo(5).DB_Table = "Location"
    DB_Gujo(5).DB_Field = _
            " [�۾���] char(40) "
            
    DB_Gujo(6).DB_Table = "Monthly"
    DB_Gujo(6).DB_Field = _
            " [�ֹι�ȣ] Char(15), " & _
            " [����] Char(10), " & _
            " [������] char(20) "
            
'    ' 2�� ��з� ���� ���̺�
'    DB_Gujo(2).DB_Table = "bType"
'    DB_Gujo(2).DB_Field = _
'            " [bType_idx] Long Identity Primary Key Not Null, " & _
'            " [bType_Code] Char(20), " & _
'            " [bType_Name] Char(50)"
'    ' 3�� �Һз� ���� ���̺�
'    DB_Gujo(3).DB_Table = "sType"
'    DB_Gujo(3).DB_Field = _
'            " [sType_idx] Long Identity Primary Key Not Null, " & _
'            " [sType_Code] Char(20), " & _
'            " [sType_Name] Char(50)"
'    ' 4�� ����/�Ǹ� ���̺�
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
'            '��/��/�귣��id/Ÿ��1id/Ÿ��2�Ѽ���/Ÿ��2����/Ÿ��2id/ȸ������ȣ �� 19�ڸ�
End Sub

'DB_info ���� ������ ��ȣ�� index ������ �ָ� �ش� ���̺��� ����.
Public Sub Create_Table(Index As Integer)
    Dim strSql As String
    '�α��� ���̺� ���� (�⺻ 1�� ���̺�)
    strSql = "CREATE Table " + DB_Gujo(Index).DB_Table + "(" + DB_Gujo(Index).DB_Field + ")"
    '�������� ����
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
    '���� ���� ���ڿ�
    Set Con = New ADODB.Connection
    
    '������ ���
    DBPath = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    
    DBPathtxt = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    'IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    
    'IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    'apple = IIf(Right("C:\", 1) = ":\", "C:\", "C:\")
    '������ �̸�
    DBFile = "Data.MDB"
    DBFiletxt = "tmpData.txt"
    'DB ���� �ҷ���
    Call DB_Info
    '���� ���ڿ� ����
    conStr = "Provider=Microsoft.Jet.OLEDB.4.0;"
    conStr = conStr + "Data Source=" + DBPath + DBFile + ";"
    conStr = conStr + "User ID=admin"
    conStr = conStr + ";Jet OLEDB:Database Password=1234;"
    Data.Connection1.ConnectionString = conStr
    Data2.Connection2.ConnectionString = conStr
    
    '������ �ִ��� �˻�
 
    If Not Dir(DBPathtxt + DBFiletxt) = "" Then
        Debug.Print "dd"
      Else
        Open "tmpData.txt" For Output As #1
        Print #1, ""
    Close #1

    End If
    
    If Not Dir(DBPath + DBFile) = "" Then
        '������ DB ����
        Con.Open conStr
    Else
        '������ ����
        If Not MsgBox("DB ������ �������� �ʽ��ϴ�." + vbCrLf _
            + "DB������ ���� ����ðڽ��ϱ�?", vbQuestion + vbYesNo, "DB ���� Ȯ��") = vbYes Then
            MsgBox "DB ������ ��� �۾��� ��� ������ �� �����ϴ�." + vbCrLf _
                + "DB ������ Ȯ���Ͻ� �� �ٽ� �õ��Ͻʽÿ�.", vbCritical + vbOKOnly, "DB ����"
            Exit Function
        End If
        '���͸� �˻�
        If Dir(DBPath, vbDirectory) = "" Then
            '���͸� ����
            Call MkDir(DBPath)
        End If
        'DB ���� ����
        Set MDB_Make = New ADOX.Catalog
        Call MDB_Make.Create(conStr)
        Set MDB_Make = Nothing
        'DB ���Ͽ� ����
        Con.Open conStr
        'Table ���� �� Field ����
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
    
    MsgMake = "DB�� ������ �߻��Ͽ����ϴ�." + vbCrLf + vbCrLf _
        + "DB : " + DBPath + DBFile + vbCrLf _
        + "Error : "
    For i = 0 To Con.Errors.Count - 1
        MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
    Next
    MsgMake = MsgMake + vbCrLf + vbCrLf _
        + "DB�� �����Ŀ� �ٽ� �õ� �ϼ���."
    MsgBox MsgMake, vbCritical + vbOKOnly, "���α׷� ����"
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
    '���� ���� ���ڿ�
    Set Con = New ADODB.Connection
    
    '������ ���
    DBPath = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    'DBPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    DBPathtxt = IIf(Right("C:\NOW_Cop\", 1) = ":\", "C:\NOW_Cop\", "C:\NOW_Cop\")
    'DBPathtxt = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    '������ �̸�
    DBFile = "Data.MDB"
    DBFiletxt = "tmpData.txt"
    'DB ���� �ҷ���
    Call DB_Info
    '���� ���ڿ� ����
    conStr = "Provider=Microsoft.Jet.OLEDB.4.0;"
    conStr = conStr + "Data Source=" + DBPath + DBFile + ";"
    conStr = conStr + "User ID=admin"
    conStr = conStr + ";Jet OLEDB:Database Password=1234;"
    
   
        Con.Open conStr
    
    DB_Conn_MDB2 = True
    Exit Function
ConnectError:
    DB_Conn_MDB2 = False
    
    MsgMake = "DB�� ������ �߻��Ͽ����ϴ�." + vbCrLf + vbCrLf _
        + "DB : " + DBPath + DBFile + vbCrLf _
        + "Error : "
    For i = 0 To Con.Errors.Count - 1
        MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
    Next
    MsgMake = MsgMake + vbCrLf + vbCrLf _
        + "DB�� �����Ŀ� �ٽ� �õ� �ϼ���."
    MsgBox MsgMake, vbCritical + vbOKOnly, "���α׷� ����"
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
        
        MsgMake = "DB�� ������ �߻��Ͽ����ϴ�." + vbCrLf + vbCrLf _
            + "Sql : " _
            + nSql + vbCrLf _
            + "Error : "
        For i = 0 To Con.Errors.Count - 1
            MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
        Next
        MsgMake = MsgMake + vbCrLf + vbCrLf _
            + "DB�� �����Ŀ� �ٽ� �õ� �ϼ���."
        MsgBox MsgMake, vbCritical + vbOKOnly, "���α׷� ����"
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
        
        MsgMake = "DB�� ������ �߻��Ͽ����ϴ�." + vbCrLf + vbCrLf _
            + "Sql : " _
            + nSql + vbCrLf _
            + "Error : "
        For i = 0 To Con.Errors.Count - 1
            MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
        Next
        MsgMake = MsgMake + vbCrLf + vbCrLf _
            + "DB�� �����Ŀ� �ٽ� �õ� �ϼ���."
        MsgBox MsgMake, vbCritical + vbOKOnly, "���α׷� ����"
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
        
        MsgMake = "DB�� ������ �߻��Ͽ����ϴ�." + vbCrLf + vbCrLf _
            + "Sql : " _
            + nSql + vbCrLf _
            + "Error : "
        For i = 0 To Con.Errors.Count - 1
            MsgMake = MsgMake + "(" + CStr(Con.Errors(i).Number) + ") " + Con.Errors(i).Description
        Next
        MsgMake = MsgMake + vbCrLf + vbCrLf _
            + "DB�� �����Ŀ� �ٽ� �õ� �ϼ���."
        MsgBox MsgMake, vbCritical + vbOKOnly, "���α׷� ����"
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
            DBField_Type = "�˼�����."
    End Select
End Function
