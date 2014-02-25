Attribute VB_Name = "modPFun"
Private Declare Function MultiByteToWideChar Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Const CP_ACP = 0 ' default to ANSI code page
Private Const CP_UTF8 = 65001 ' default to UTF-8 code page

'*********************************************************������������****************************************************
Function ParamsEncode(InParam As String) As String
    InParam = Replace(InParam, ":", "%3A")
    InParam = Replace(InParam, "+", "%2B")
    InParam = Replace(InParam, "=", "%3D")
    InParam = Replace(InParam, "/", "%2F")
    ParamsEncode = InParam
End Function

Function BytesToBstr(body, Optional btype As Integer = 1) As String

On Error GoTo berrfix
    Dim objstream As New ADODB.Stream
    objstream.Type = 1 'adTypeBinary
    objstream.Mode = 3
    objstream.Open
    objstream.Write body
    objstream.Position = 0
    objstream.Type = 2
    If btype = 2 Then
        objstream.Charset = "utf-8"
    Else
        objstream.Charset = "GB2312"
    End If
    
    BytesToBstr = objstream.ReadText
    objstream.Close
    Set objstream = Nothing
    Exit Function
        
berrfix:
  SavePage "[" & Now() & "]" & Err.Description, "error"
  Err.Clear
  BytesToBstr = ""
End Function


'�ַ�ת UTF8
Public Function EncodeToBytes(ByVal sData As String) As Byte() ' Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sData), -1, 0, 0, 0, 0) - 1
    If nSize = 0 Then Exit Function
    ReDim aRetn(0 To nSize - 1) As Byte
    WideCharToMultiByte CP_UTF8, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize, 0, 0
    EncodeToBytes = aRetn
    Erase aRetn
End Function

 'UTF8 ת�ַ�
Public Function DecodeToBytes(ByVal sData As String) As Byte() ' Note: Len(sData) > 0
Dim aRetn() As Byte
Dim nSize As Long
nSize = MultiByteToWideChar(CP_UTF8, 0, StrPtr(sData), -1, 0, 0) - 1
If nSize = 0 Then Exit Function
ReDim aRetn(0 To 2 * nSize - 1) As Byte
MultiByteToWideChar CP_UTF8, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize
DecodeToBytes = aRetn
Erase aRetn
End Function

Function U8Decode(enStr)
      '����һ����%�ָ����ַ������ȷֳ����飬����utf8�������жϲ������
      '����:�� E5 85 B3  ��  E9 94 AE ��   E5 AD 97
      '���:�� B9D8  ��  BCFC ��   D7D6
      Dim c, i, i2, v, deStr, WeiS

      For i = 1 To Len(enStr)
        c = Mid(enStr, i, 1)
        If c = "%" Then
          v = c16to2(Mid(enStr, i + 1, 2))
          '�жϵ�һ�γ���0��λ�ã�
          '������1(���ֽ�)��3(3-1�ֽ�)��4��5��6��7��������2�ʹ���7
          '�����ϵ�7��ʵ�ʲ��ᳬ��3��
          WeiS = InStr(v, "0")
          v = Right(v, Len(v) - WeiS) '��һ��ȥ������ߵ�WeiS��
          i = i + 3
          For i2 = 2 To WeiS - 1
            c = c16to2(Mid(enStr, i + 1, 2))
            c = Right(c, Len(c) - 2) '����ȥ������ߵ�����
            v = v & c
            i = i + 3
          Next
          If Len(c2to16(v)) = 4 Then
            deStr = deStr & ChrW(c2to10(v))
          Else
            deStr = deStr & Chr(c2to10(v))
          End If
          i = i - 1
        Else
          If c = "+" Then
            deStr = deStr & " "
          Else
            deStr = deStr & c
          End If
        End If
      Next
      U8Decode = deStr
    End Function

    Function c16to2(x)
     '�������������ת��16���Ƶ�2���Ƶģ��������κγ��ȵģ�һ��ת��UTF-8��ʱ�����������ȣ�����A9
     '���磺���롰C2����ת���ɡ�11000010��,����1100��"c"��10���Ƶ�12��1100������ô2��10������4λҪ����ɣ�0010����
     Dim tempStr
     Dim i: i = 0 '��ʱ��ָ��

     For i = 1 To Len(Trim(x))
      tempStr = c10to2(CInt(Int("&h" & Mid(x, i, 1))))
      Do While Len(tempStr) < 4
       tempStr = "0" & tempStr '�������4λ��ô����4λ��
      Loop
      c16to2 = c16to2 & tempStr
     Next
    End Function

    Function c2to16(x)
      '2���Ƶ�16���Ƶ�ת����ÿ4��0��1ת����һ��16������ĸ�����볤�ȵ�Ȼ�����ܲ���4�ı�����

      Dim i: i = 1 '��ʱ��ָ��
      For i = 1 To Len(x) Step 4
       c2to16 = c2to16 & Hex(c2to10(Mid(x, i, 4)))
      Next
    End Function

    Function c2to10(x)
      '������2���Ƶ�10���Ƶ�ת����������ת16��������Ҫ��4λǰ�㲹�롣
      '��Ϊ������������ã��Ժ�Ҳ���õ�������ͨѶ��Ӳ������Ӧ��֪����
      '�������ַ������������
       c2to10 = 0
       If x = "0" Then Exit Function '�����0�Ļ�ֱ�ӵ�0������
       Dim i: i = 0 '��ʱ��ָ��
       For i = 0 To Len(x) - 1 '��������8421����㣬��������ʼѧ�������ʱ��ͻᣬ�û���������ǵ�л��������������
        If Mid(x, Len(x) - i, 1) = "1" Then c2to10 = c2to10 + 2 ^ (i)
       Next
    End Function

    Function c10to2(x)
    '10���Ƶ�2���Ƶ�ת��
      Dim sign, Result
      Result = ""
      '����
      sign = Sgn(x)
      x = Abs(x)
      If x = 0 Then
        c10to2 = 0
        Exit Function
      End If
      Do Until x = "0"
        Result = Result & (x Mod 2)
        x = x \ 2
      Loop
      Result = StrReverse(Result)
      If sign = -1 Then
        c10to2 = "-" & Result
      Else
        c10to2 = Result
      End If
    End Function



'********************************************
'��������UTF2GB()
'��  �ã�'UTF8תGB2312
'����ֵ: ��
'��  ��: UTFStr
'��  ��: 2008��06��01��
'��  ��: zxw
'********************************************'
  Function UTF2GB(UTFStr)
    For Dig = 1 To Len(UTFStr)
    If Mid(UTFStr, Dig, 1) = "%" Then
    If Len(UTFStr) >= Dig + 8 Then
    GBStr = GBStr & ConvChinese(Mid(UTFStr, Dig, 9))
    Dig = Dig + 8
    Else
    GBStr = GBStr & Mid(UTFStr, Dig, 1)
    End If
    Else
    GBStr = GBStr & Mid(UTFStr, Dig, 1)
    End If
    Next
    UTF2GB = GBStr
  End Function
   Function ConvChinese(x)
    a = Split(Mid(x, 2), "%")
    i = 0
    j = 0
      
    For i = 0 To UBound(a)
    a(i) = binNumber(Oct(a(i)))
    Next
      
    For i = 0 To UBound(a) - 1
    DigS = InStr(a(i), "0")
    Unicode = ""
    For j = 1 To DigS - 1
    If j = 1 Then
    a(i) = Right(a(i), Len(a(i)) - DigS)
    Unicode = Unicode & a(i)
    Else
    i = i + 1
    a(i) = Right(a(i), Len(a(i)) - 2)
    Unicode = Unicode & a(i)
    End If
    Next
    If Len(Hex(decNumber(Unicode))) = 4 Then
    ConvChinese = ConvChinese & ChrW(Int("&H" & Hex(decNumber(Unicode))))
    Else
    ConvChinese = ConvChinese & Chr(Int("&H" & Hex(decNumber(Unicode))))
    End If
    Next
  End Function

'********************************************
'��������URLEncoding()
'��  �ã�����Post �������ݱ��뺯��
'����ֵ: ��
'��  ��:vstrIn
'��  ��: 2008��06��01��
'��  ��: zxw
'********************************************'
Function URLEncoding(ByVal vstrIn As String)

    Dim i
    strReturn = ""
    For i = 1 To Len(vstrIn)
        ThisChr = Mid(vstrIn, i, 1)
        If Abs(Asc(ThisChr)) < &HFF Then
            strReturn = strReturn & ThisChr
        Else
            innerCode = Asc(ThisChr)
            If innerCode < 0 Then
                innerCode = innerCode + &H10000
            End If
            Hight8 = (innerCode And &HFF00) \ &HFF
            Low8 = innerCode And &HFF
            strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
        End If
    Next
    strReturn = Replace(strReturn, Chr(32), "%20")
    URLEncoding = strReturn
End Function

'********************************************
'��������URLEncodeUTF8()
'��  �ã�����Post �������ݱ��뺯��
'����ֵ: ��
'��  ��:vstrIn
'��  ��: 2008��06��01��
'��  ��: zxw
'********************************************'
Function URLEncodeUTF8(ByVal strin As String)

    Dim i As Long, vstrIn() As Byte
    strReturn = ""
    vstrIn = EncodeToBytes(strin)
    For i = 0 To UBound(vstrIn)
        ThisChr = vstrIn(i)
        If Abs(ThisChr) < &HFF Then
            strReturn = strReturn & "%" & Hex(ThisChr)
        Else
            innerCode = ThisChr
            If innerCode < 0 Then
                innerCode = innerCode + &H10000
            End If
            Hight8 = (innerCode And &HFF00) \ &HFF
            Low8 = innerCode And &HFF
            strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
        End If
    Next
    strReturn = Replace(strReturn, Chr(32), "%20")
    URLEncodeUTF8 = strReturn
End Function


'********************************************
'��������URLDecode()
'��  �ã�����Post �������ݽ��뺯��
'����ֵ: ��
'��  ��: enStr
'��  ��: 2008��06��01��
'��  ��: zxw
'********************************************'
Public Function URLDecode(sEncodedURL As String) As String
    On Error GoTo Catch
    
    Dim iLoop As Integer
    Dim sRtn As String
    Dim sTmp As String

    If Len(sEncodedURL) > 0 Then
        ' Loop through each char


        For iLoop = 1 To Len(sEncodedURL)
            sTmp = Mid(sEncodedURL, iLoop, 1)
            sTmp = Replace(sTmp, "+", " ")
            ' If char is % then get next two chars
            ' and convert from HEX to decimal


            If sTmp = "%" And Len(sEncodedURL) + 1 > iLoop + 2 Then
                sTmp = Mid(sEncodedURL, iLoop + 1, 2)
                sTmp = Chr(CDec("&H" & sTmp))
                ' Increment loop by 2
                iLoop = iLoop + 2
            End If
            sRtn = sRtn & sTmp
        Next iLoop
        URLDecode = sRtn
    End If
Finally:
    Exit Function
Catch:
    URLDecode = ""
    Resume Finally
End Function

'��post���ݹ����Ĳ�����urldecode���봦��(�벻Ҫ�޸�)
Public Function URLDecodeing(strURL As String) As String
    Dim strChar As String
    Dim strText As String
    Dim strTemp As String
    Dim strRet As String
    Dim LngNum As Long

    Dim i As Integer

    For i = 1 To Len(strURL)

        strChar = Mid(strURL, i, 1)

        Select Case strChar
          Case "+"
            strText = strText & " "
          Case "%"
            strTemp = Mid(strURL, i + 1, 2) '��ʱȡ2λ

            LngNum = Val("&H" & strTemp)

            '>127��Ϊ����
            If LngNum < 128 Then
                strRet = Chr(LngNum)
                i = i + 2
            Else
                strTemp = strTemp & Mid(strURL, i + 4, 2)
                strRet = Chr(Val("&H" & strTemp))
                i = i + 5
            End If
            strText = strText & strRet

          Case Else
            strText = strText & strChar
        End Select
    Next

    URLDecodeing = strText
End Function


'********************************************
'��������mysubstr()
'��  �ã��ַ�����ȡ����
'����ֵ: ��
'��  ��:s, sstart, sstop
'��  ��: 2008��06��01��
'��  ��: zxw
'********************************************'
Function mySubstr(s, sStart, sstop)
    src = s
    p1 = InStr(src, sStart)
    If p1 = 0 Then
       mySubstr = ""
       Exit Function
    End If
    src = Mid(src, p1 + Len(sStart), Len(src) - Len(sStart) - p1 + 1)
    p1 = InStr(src, sstop)
    src = Mid(src, 1, p1 - 1)
    mySubstr = src
End Function


'����ļ��Ƿ����
Function FileExists(ByVal filename As String) As Boolean
On Error Resume Next
    Dim Fso As New FileSystemObject
    
    If Fso.FileExists(filename) = True Then
        FileExists = True
    Else
        FileExists = False
    End If
    Set Fso = Nothing
    
    If Err Then
        MsgBox Err.Description, vbCritical, "����"
        FileExists = False
        Err.Clear
    End If
    
End Function

'10����ת2����
Function binNumber(ByVal sNumber As Long) As String
Dim BinaryStr As String
BinaryStr = ""
Do
   BinaryStr = sNumber Mod 2 & BinaryStr
   sNumber = sNumber \ 2
Loop Until sNumber = 0
binNumber = BinaryStr
End Function

'2�����ַ���ת10����
Function decNumber(ByVal dNumber As String) As Long
 Dim DecStr As Long
  DecStr = 0
 Dim q As String
 Dim i As Integer
 Dim w As Integer
  w = Len(dNumber)
  i = 1
  Do
  q = Mid(dNumber, i, 1)
  DecStr = Val(q) * (2 ^ (w - 1)) + DecStr
  i = i + 1
  w = w - 1
  Loop Until i >= Len(dNumber) + 1
  decNumber = DecStr
End Function

'��ⰴ��
Public Sub EnterToTab(Keyasc As Integer)
    If Keyasc = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

'=======================================================================
'��������:Fun_SaveImgToFile(data, filename)
'��    ��:�洢��֤��ͼƬ
'��    ��:
'         1:StrAllResponseHeaders
'         2:Key
'��    ��:ZXX
'����:2008-07-17
'========================================================================
Function Fun_SaveImgToFile(data, filename, FilePath) As String
    Dim objstream As New ADODB.Stream  'server.CreateObject("ADODB.Stream")
    objstream.Type = 1
    objstream.Open
    objstream.Write (data)
    
    Dim Fso As New Scripting.FileSystemObject  '= server.CreateObject("Scripting.FileSystemObject")
    '�ж�Ŀ¼�Ƿ����
    If Not Fso.FolderExists(FilePath) Then
      Fso.CreateFolder (FilePath)
    End If
    If Fso.FileExists(FilePath & filename) = True Then
        Fso.DeleteFile (FilePath & filename)
    End If
    Set Fso = Nothing
    objstream.SaveToFile (FilePath & filename)
    objstream.Close
    Set objstream = Nothing
    Fun_SaveImgToFile = FilePath & filename
        
End Function

'********************************************************************
'��������:SaveZcPage()
'��    ��:��¼�ٷ���ֵ�����ҳ��
'��    ��:bodyHtml ҳ������  RndTemp������������ļ�������
'�� �� ֵ:��
'��    ��:ZXX
'����:2008-11-11
'********************************************************************

Sub SavePage(bodyHtml, RndTemp)
    '----step1:��¼ҳ��-----------
    Dim FilePath As String
    FilePath = Replace(App.Path, "\", "/") & "/temphtml/" & Format(Now(), "yy_mm_dd") & "/"
    
    Dim objFso As New Scripting.FileSystemObject
    '�ж�Ŀ¼�Ƿ����
    If Not objFso.FolderExists(FilePath) Then
      objFso.CreateFolder (FilePath)
    End If
    logfilename = RndTemp & ".log"
    logFilePath = FilePath & logfilename
    If objFso.FileExists(logFilePath) Then
        Set logFile = objFso.OpenTextFile(logFilePath, 8)
        logFile.WriteLine (bodyHtml)
        logFile.Close
    Else
        Set logFile = objFso.CreateTextFile(logFilePath)
        logFile.WriteLine (bodyHtml)
        logFile.Close
    End If
    Set objFso = Nothing
    
End Sub


Public Function MyReadFile(FilePath As String, FileBuffer() As Byte)
    Dim FileSize As Long '�ļ�����
    FileSize = FileLen(FilePath) '��ȡ�ļ�����
    ReDim FileBuffer(FileSize - 1) As Byte
    
    Open FilePath For Binary As #1
        Get #1, , FileBuffer
    Close
End Function

'����JSON indexΪ0��������,����ֵ�����ַ���
Public Function parseJSON(ByVal str As String, ByVal key As String, ByVal Index As Integer) As String()
    'ȥ[]
    If Left(str, 1) = "[" Then str = Right(str, Len(str) - 1)
    If Right(str, 1) = "]" Then str = Left(str, Len(str) - 1)
    
    ':->=    "->null
    str = Replace(str, "},{", "==|==")
    str = Replace(str, "{", "")
    str = Replace(str, "}", "")
    str = Replace(str, ":", "=")
    str = Replace(str, """", "")
    
    
    '����
    Dim tmpArr1() As String, tmpArr2() As String, tmpArr3(1) As String, returnStr As String
    
    tmpArr1 = Split(str, "==|==")
    
    'û��ָ���ؼ��־ͷ���ȫ������
    If key = "" Then
    
        'indexΪ0��������
        If Index = 0 Then
            parseJSON = tmpArr1
            Exit Function
           
        '����ֵ�����ַ���
        Else
        
            '�Ƿ�����err
            If Index < 0 Or Index > UBound(tmpArr1) + 1 Then
                tmpArr3(0) = ""
                parseJSON = tmpArr3
                Exit Function
            Else
                tmpArr3(0) = tmpArr1(Index - 1)
                parseJSON = tmpArr3
                Exit Function
            End If
        End If
    
    'ָ��key����valuseֵ������
    Else
    
        '�������ź���ȡֵ��
        For i = 0 To UBound(tmpArr1)
            tmpArr1(i) = tmpArr1(i) & ","
        Next
    
        'indexΪ0��������
        If Index = 0 Then
            For i = 0 To UBound(tmpArr1)
                tmpArr2(i) = mySubstr(tmpArr1(i), key & "=", ",")
            Next
            parseJSON = tmpArr2
            Exit Function
           
        '����ֵ�����ַ���
        Else
            
            '�Ƿ�����err
            If Index < 0 Or Index > UBound(tmpArr1) + 1 Then
                tmpArr3(0) = ""
                parseJSON = tmpArr3
                Exit Function
            Else
                tmpArr3(0) = mySubstr(tmpArr1(Index - 1), key & "=", ",")
                parseJSON = tmpArr3
                Exit Function
            End If
        End If
        

    End If
    
End Function

'����JSON�ܼ�¼����
Public Function lenJSON(ByVal JSONstr As String) As Integer
    lenJSON = UBound(Split(JSONstr, "},{")) + 1
End Function


'��������ֵ
Function CheckFunRe(ByVal restr As String, otype As Integer) As String
    If InStr(restr, "|") = 0 Then
       CheckFunRe = ""
       Exit Function
    End If
    
    If otype = 1 Then
        CheckFunRe = Left(restr, 1)
        Exit Function
    ElseIf otype = 2 Then
        CheckFunRe = Right(restr, Len(restr) - 2)
        Exit Function
    Else
        CheckFunRe = Right(restr, Len(restr) - 2)
        Exit Function
    End If
End Function
