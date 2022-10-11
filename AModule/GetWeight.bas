Attribute VB_Name = "GetWeightString"
Public WeightStr As String

'**********************************
'接收模块
'**********************************

Public bytReceiveByte() As Byte     '接收到的字节
Public intReceiveLen As Integer     '接收到的字节数
Public timbangData     As String

Public intHexWidth As Integer       '显示列数（电子秤的输出位数小循环）

Public strAscii As String           '电子秤输出的ASCII码
Public Function GetWeightA1(RedData As String) As String
    'RedData取值“+004038216..+004038216..+004038216..+00403”表示为40.38(XK3190-A9型电子秤)
    'RedData取值“+00025011D..+00025011D..+00025011D..+00025011”表示为25.01(XK3190-A1+型电子秤)
    '实际上500kg量程的XK3190-A1+电子秤的分辨率设为0.1kg所以读数为25.0kg
    Dim FFirst, FSecond, Fthird, i, FLength As Integer
    Dim OriStr1 As String
    Dim OriStr2 As String
    Dim OriStr3 As String
    
   If RedData = "" Then
        GetWeightA1 = ""
        Exit Function
   End If


        FFirst = InStr(1, RedData, "+")
        FSecond = InStr(FFirst + 1, RedData, "+")
        Fthird = InStr(FSecond + 1, RedData, "+")

        If FFirst >= FSecond Then
            GetWeightA1 = ""
            Exit Function
        End If

        OriStr1 = Mid(RedData, FFirst + 1, 6)
        OriStr2 = Mid(RedData, FSecond + 1, 6)
        OriStr3 = Mid(RedData, Fthird + 1, 6)

        If OriStr1 = OriStr2 And OriStr2 = OriStr3 Then
        
        'GetWeightA9 = AddZero(Str(Val(OriStr1) / 100))
        'XK3190-A1电子秤实际情况为:
        'GetWeightA1 = AddZero(Str(Val(OriStr1) / 10))
        'GetWeightA1 = AddZero(Str(Val(OriStr1) / 100))
        GetWeightA1 = str(Val(OriStr1))
        Else
        
        GetWeightA1 = ""
        End If

End Function

Public Function AddZero(FOriStr As String) As String
    '为重量加上二位有效数字

    Dim PointPostion As Integer
    Dim s As String
    s = Trim(FOriStr)
    If Left(s, 1) = "." Then
        s = Replace(s, ".", "0.")
    End If
        PointPostion = InStr(s, ".")
        If PointPostion = 0 Then
            AddZero = s + ".00"
        ElseIf Len(Mid(s, PointPostion + 1, Len(s) - PointPostion)) = 1 Then
            AddZero = Mid(s, 1, PointPostion) + Mid(s, PointPostion + 1, Len(s) - PointPostion) + "0"
            
        Else
            AddZero = s
        End If

End Function


'**********************************
'输入处理
'处理接收到的字节流,并保存在全局变量
'bytReceiveRyte()
'**********************************
Public Sub InputManage(bytInput() As Byte, intInputLenth As Integer)
On Error Resume Next
    Dim n As Integer                                '定义变量及初始化
    
    ReDim Preserve bytReceiveByte(intReceiveLen + intInputLenth)

    For n = 1 To intInputLenth Step 1
        bytReceiveByte(intReceiveLen + n - 1) = bytInput(n - 1)
    Next n
    
    intReceiveLen = intReceiveLen + intInputLenth
    
End Sub

'***********************************
'为输出准备文本
Public Sub GetDisplayText(ByRef sText As Label)

    Dim n As Integer
    Dim intValue As Integer
    Dim intHighHex As Integer
    Dim intLowHex As Integer
    Dim strSingleChr As String * 1
    
    strAscii = ""            '设置初值

    
    For n = 1 To intReceiveLen
        intValue = bytReceiveByte(n - 1)
        
        If intValue < 32 Or intValue > 128 Then         '处理非法字符
            strSingleChr = Chr(46)                      '对于不能显示的ASCII码,
        Else                                            '用"."表示
            strSingleChr = Chr(intValue)
        End If
        
        strAscii = strAscii + strSingleChr
        
        intHighHex = intValue \ 16
        intLowHex = intValue - intHighHex * 16
        
        If intHighHex < 10 Then
            intHighHex = intHighHex + 48
        Else
            intHighHex = intHighHex + 55
        End If
        If intLowHex < 10 Then
            intLowHex = intLowHex + 48
        Else
            intLowHex = intLowHex + 55
        End If
        
         If (n Mod intHexWidth) = 0 Then                 '设置换行
            'strAscii = strAscii + Chr$(13) + Chr$(10)
            If Len(strAscii) > 60 Then
              WeightStr = Right(strAscii, 60)
              sText.Caption = GetWeightA1(WeightStr)
            End If
            '------------------------------------------------
        End If
        If intReceiveLen > 524 Then
        Call ClearWeight
        Exit Sub
        End If
    Next n
End Sub

Private Sub ClearWeight()

    Dim bytTemp(0) As Byte
    
    ReDim bytReceiveByte(0)
    
    intReceiveLen = 0
    
    Call InputManage(bytTemp, 0)
    
End Sub
