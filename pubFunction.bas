Attribute VB_Name = "modPFun"
Private Declare Function MultiByteToWideChar Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Const CP_ACP = 0 ' default to ANSI code page
Private Const CP_UTF8 = 65001 ' default to UTF-8 code page

'*********************************************************辅助函数区域****************************************************
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


'字符转 UTF8
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

 'UTF8 转字符
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
      '输入一堆有%分隔的字符串，先分成数组，根据utf8规则来判断补齐规则
      '输入:关 E5 85 B3  键  E9 94 AE 字   E5 AD 97
      '输出:关 B9D8  键  BCFC 字   D7D6
      Dim c, i, i2, v, deStr, WeiS

      For i = 1 To Len(enStr)
        c = Mid(enStr, i, 1)
        If c = "%" Then
          v = c16to2(Mid(enStr, i + 1, 2))
          '判断第一次出现0的位置，
          '可能是1(单字节)，3(3-1字节)，4，5，6，7不可能是2和大于7
          '理论上到7，实际不会超过3。
          WeiS = InStr(v, "0")
          v = Right(v, Len(v) - WeiS) '第一个去掉最左边的WeiS个
          i = i + 3
          For i2 = 2 To WeiS - 1
            c = c16to2(Mid(enStr, i + 1, 2))
            c = Right(c, Len(c) - 2) '其余去掉最左边的两个
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
     '这个函数是用来转换16进制到2进制的，可以是任何长度的，一般转换UTF-8的时候是两个长度，比如A9
     '比如：输入“C2”，转化成“11000010”,其中1100是"c"是10进制的12（1100），那么2（10）不足4位要补齐成（0010）。
     Dim tempStr
     Dim i: i = 0 '临时的指针

     For i = 1 To Len(Trim(x))
      tempStr = c10to2(CInt(Int("&h" & Mid(x, i, 1))))
      Do While Len(tempStr) < 4
       tempStr = "0" & tempStr '如果不足4位那么补齐4位数
      Loop
      c16to2 = c16to2 & tempStr
     Next
    End Function

    Function c2to16(x)
      '2进制到16进制的转换，每4个0或1转换成一个16进制字母，输入长度当然不可能不是4的倍数了

      Dim i: i = 1 '临时的指针
      For i = 1 To Len(x) Step 4
       c2to16 = c2to16 & Hex(c2to10(Mid(x, i, 4)))
      Next
    End Function

    Function c2to10(x)
      '单纯的2进制到10进制的转换，不考虑转16进制所需要的4位前零补齐。
      '因为这个函数很有用！以后也会用到，做过通讯和硬件的人应该知道。
      '这里用字符串代表二进制
       c2to10 = 0
       If x = "0" Then Exit Function '如果是0的话直接得0就完事
       Dim i: i = 0 '临时的指针
       For i = 0 To Len(x) - 1 '否则利用8421码计算，这个从我最开始学计算机的时候就会，好怀念当初教我们的谢道建老先生啊！
        If Mid(x, Len(x) - i, 1) = "1" Then c2to10 = c2to10 + 2 ^ (i)
       Next
    End Function

    Function c10to2(x)
    '10进制到2进制的转换
      Dim sign, Result
      Result = ""
      '符号
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
'函数名：UTF2GB()
'作  用：'UTF8转GB2312
'返回值: 无
'参  数: UTFStr
'日  期: 2008年06月01日
'作  者: zxw
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
'函数名：URLEncoding()
'作  用：中文Post 传输数据编码函数
'返回值: 无
'参  数:vstrIn
'日  期: 2008年06月01日
'作  者: zxw
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
'函数名：URLEncodeUTF8()
'作  用：中文Post 传输数据编码函数
'返回值: 无
'参  数:vstrIn
'日  期: 2008年06月01日
'作  者: zxw
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
'函数名：URLDecode()
'作  用：中文Post 传输数据解码函数
'返回值: 无
'参  数: enStr
'日  期: 2008年06月01日
'作  者: zxw
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

'对post传递过来的参数作urldecode编码处理(请不要修改)
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
            strTemp = Mid(strURL, i + 1, 2) '暂时取2位

            LngNum = Val("&H" & strTemp)

            '>127即为汉字
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
'函数名：mysubstr()
'作  用：字符串截取函数
'返回值: 无
'参  数:s, sstart, sstop
'日  期: 2008年06月01日
'作  者: zxw
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


'检测文件是否存在
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
        MsgBox Err.Description, vbCritical, "错误"
        FileExists = False
        Err.Clear
    End If
    
End Function

'10进制转2进制
Function binNumber(ByVal sNumber As Long) As String
Dim BinaryStr As String
BinaryStr = ""
Do
   BinaryStr = sNumber Mod 2 & BinaryStr
   sNumber = sNumber \ 2
Loop Until sNumber = 0
binNumber = BinaryStr
End Function

'2进制字符串转10进制
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

'检测按键
Public Sub EnterToTab(Keyasc As Integer)
    If Keyasc = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

'=======================================================================
'过程名称:Fun_SaveImgToFile(data, filename)
'作    用:存储验证码图片
'参    数:
'         1:StrAllResponseHeaders
'         2:Key
'作    者:ZXX
'日期:2008-07-17
'========================================================================
Function Fun_SaveImgToFile(data, filename, FilePath) As String
    Dim objstream As New ADODB.Stream  'server.CreateObject("ADODB.Stream")
    objstream.Type = 1
    objstream.Open
    objstream.Write (data)
    
    Dim Fso As New Scripting.FileSystemObject  '= server.CreateObject("Scripting.FileSystemObject")
    '判断目录是否存在
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
'过程名称:SaveZcPage()
'作    用:记录官方充值的最后页面
'参    数:bodyHtml 页面内容  RndTemp随机码用作于文件名命名
'返 回 值:无
'作    者:ZXX
'日期:2008-11-11
'********************************************************************

Sub SavePage(bodyHtml, RndTemp)
    '----step1:记录页面-----------
    Dim FilePath As String
    FilePath = Replace(App.Path, "\", "/") & "/temphtml/" & Format(Now(), "yy_mm_dd") & "/"
    
    Dim objFso As New Scripting.FileSystemObject
    '判断目录是否存在
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
    Dim FileSize As Long '文件长度
    FileSize = FileLen(FilePath) '获取文件长度
    ReDim FileBuffer(FileSize - 1) As Byte
    
    Open FilePath For Binary As #1
        Get #1, , FileBuffer
    Close
End Function

'解析JSON index为0返回数组,其他值返回字符串
Public Function parseJSON(ByVal str As String, ByVal key As String, ByVal Index As Integer) As String()
    '去[]
    If Left(str, 1) = "[" Then str = Right(str, Len(str) - 1)
    If Right(str, 1) = "]" Then str = Left(str, Len(str) - 1)
    
    ':->=    "->null
    str = Replace(str, "},{", "==|==")
    str = Replace(str, "{", "")
    str = Replace(str, "}", "")
    str = Replace(str, ":", "=")
    str = Replace(str, """", "")
    
    
    '分组
    Dim tmpArr1() As String, tmpArr2() As String, tmpArr3(1) As String, returnStr As String
    
    tmpArr1 = Split(str, "==|==")
    
    '没有指定关键字就返回全部数据
    If key = "" Then
    
        'index为0返回数组
        If Index = 0 Then
            parseJSON = tmpArr1
            Exit Function
           
        '其他值返回字符串
        Else
        
            '非法返回err
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
    
    '指定key返回valuse值或数组
    Else
    
        '补个逗号后面取值用
        For i = 0 To UBound(tmpArr1)
            tmpArr1(i) = tmpArr1(i) & ","
        Next
    
        'index为0返回数组
        If Index = 0 Then
            For i = 0 To UBound(tmpArr1)
                tmpArr2(i) = mySubstr(tmpArr1(i), key & "=", ",")
            Next
            parseJSON = tmpArr2
            Exit Function
           
        '其他值返回字符串
        Else
            
            '非法返回err
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

'解析JSON总记录条数
Public Function lenJSON(ByVal JSONstr As String) As Integer
    lenJSON = UBound(Split(JSONstr, "},{")) + 1
End Function


'解析返回值
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
