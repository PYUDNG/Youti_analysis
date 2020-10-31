Dim FSO, ws, SA, ADO, wn
Dim SelfFolderPath, UserName, Self
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")
Set SA = CreateObject("Shell.Application")
'Set ADO = CreateObject("ADODB.STREAM")
'Set wn = CreateObject("Wscript.Network")

Call GetUAC(1, False)

SelfFolderPath = FormatPath(FSO.GetFile(WScript.ScriptFullName).ParentFolder.Path)
'UserName = wn.UserName
'Self = FSO.OpenTextFile(Wscript.ScriptFullName).ReadAll

Dim Account, Phone, Password
Account = InputBox("输入账号：", WScript.ScriptName, ""): If IsEmpty(Account) Then WScript.Quit
Phone = InputBox("输入手机号：", WScript.ScriptName, Account): If IsEmpty(Phone) Then WScript.Quit
Password = InputBox("输入密码：", WScript.ScriptName, ""): If IsEmpty(Password) Then WScript.Quit

Dim SavePath, OAS
SavePath = SelfFolderPath & "MyAccount.json"
Set OAS = New AccountSaver
OAS.Password_Length = 2048

Call OAS.SaveAccount(Account, Phone, Password, SavePath)

Dim ReadResult: ReadResult = OAS.ReadAccount(SavePath)

MsgBox "账户信息已加密保存" & vbCrLf & _
       "账号：" & ReadResult(0) & vbCrLf & _
       "手机号：" & ReadResult(1) & vbCrLf & _
       "密码：" & ReadResult(2) & vbCrLf, 64, WScript.ScriptName

Class AccountSaver
	Private Sub Class_Initialize()
		Set VJ = New VbsJson
		Set VG = New Vigenere
		
		PasswordKeyLength = 1024
	End Sub
	
	Private Sub Class_Terminate()
		Set VJ = Nothing
		Set VG = Nothing
	End Sub
	
	Private VJ, VG
	Private PasswordKeyLength
	
	Property Get Password_Length()
		Password_Length = PasswordKeyLength
	End Property
	
	Property Let Password_Length(ByVal NewLength)
		PasswordKeyLength = NewLength
	End Property
	
	Public Function ReadAccount(ByVal FilePath)
		''' 从文件读取加密的账号信息，返回[Account, Phone, Password]明文数组 '''
		If Not FSO.FileExists(FilePath) Then
			ReadAccount = -1
			Exit Function
		End If
		
		Dim Json
		Set Json = VJ.Decode(FSO.OpenTextFile(FilePath).ReadAll())
		
		Dim p
		Dim Enc_Account, Enc_Phone, Enc_Password
		
		Enc_Account = Json("account")
		Enc_Phone = Json("phone")
		Enc_Password = Json("password")
		p = Json("p")
		
		ReadAccount = Array( _
			VG.Compatible_Vigenere(Enc_Account, p, -1), _
			VG.Compatible_Vigenere(Enc_Phone, p, -1), _
			VG.Compatible_Vigenere(Enc_Password, p, -1) _
		)
	End Function
	
	Public Function SaveAccount(ByVal Account, ByVal Phone, ByVal Password, ByVal SavePath)
		''' 加密储存账号信息到文件 '''
		Dim p
		Dim Enc_Account, Enc_Phone, Enc_Password
		p = CreateRandomizedText(PasswordKeyLength)
		Enc_Account = VG.Compatible_Vigenere(Account, p, 1)
		Enc_Phone = VG.Compatible_Vigenere(Phone, p, 1)
		Enc_Password = VG.Compatible_Vigenere(Password, p, 1)
		
		Dim Json, JsonText
		Set Json = CreateObject("Scripting.Dictionary")
		Json("account") = Enc_Account
		Json("phone") = Enc_Phone
		Json("password") = Enc_Password
		Json("p") = p
		JsonText = VJ.Encode(Json)
		
		FSO.CreateTextFile(SavePath, True).Write JsonText
	End Function
End Class

Function GetUAC(ByVal Host, ByVal Hide)
    Dim HostName, Hidden, Args, i
    If Not Hide Then Hidden = 1
    If Host = 1 Then HostName = "wscript.exe"
    If Host = 2 Then HostName = "cscript.exe"
    If WScript.Arguments.Count > 0 Then
        For i = 0 To WScript.Arguments.Count - 1
            If Not(i = 0 And (WScript.Arguments(i) = "uac" Or WScript.Arguments(i) = "uacHidden")) Then Args = Args & " " & Chr(34) & WScript.Arguments(i) & Chr(34)
        Next
    End If
    If WScript.Arguments.Count = 0 Then
        SA.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & chr(34) & " uac" & Args, "", "runas", 1
        WScript.Quit
    ElseIf LCase(Right(WScript.FullName, 12)) <> "\" & HostName Or WScript.Arguments(0) <> "uacHidden" Then
        ws.Run HostName & " """ & WScript.ScriptFullName & """ uacHidden" & Args, Hidden, False
        WScript.Quit
    End If
    If Host = 2 Then ExecuteGlobal "Dim SI, SO: Set SI = Wscript.StdIn: Set SO = Wscript.StdOut"
End Function

Function FormatPath(ByVal Path)
    If Not Right(Path, 1) = "\" Then
        Path = Path & "\"
    End If
    FormatPath = Path
End Function

Function CreateTempPath(ByVal IsFolder)
    Dim TempPath
    TempPath = FSO.GetSpecialFolder(2) & "\" & FSO.GetTempName()
    If IsFolder Then TempPath = FormatPath(TempPath)
    CreateTempPath = TempPath
End Function

Function CreateRandomizedText(ByVal pwdlen)
    On Error Resume Next
    Dim LLength, rdnum
    Execute "Dim Final(" & CStr(pwdlen-1) & ")"
    Const Letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    LLength = Len(Letters)
    For i = 1 To pwdlen
        Randomize
        rdnum = Int(LLength * Rnd) + 1
        Final(i-1) = Mid(Letters, rdnum, 1)
    Next
    CreateRandomizedText = Join(Final, "")
End Function

Class Vigenere
	' This Class is written by: PY_DNG                / 此类作者：PY_DNG
	' Using is free, but you must not clear this mark / 可以任意使用，但必须保留这几行注释
	' Prohibited For commercial use and illegal use   / 禁止用于商业以及非法用途
    Private Sub Class_Initialize()
    End Sub
    
    Private Sub Class_Terminate()
    End Sub
    
    Public Function Base64ToText(ByVal Content)
        ' BASE64_Dencode a text content with BASE64.
        ' Creates the file with the Dencoded binary content
        Dim ADODB, ContentBytes, oXML, oNode, t
        Set oXML = CreateObject("Msxml2.DOMDocument")
        Set oNode = oXML.CreateElement("base64")
        oNode.DataType = "bin.base64"
        oNode.Text = Content
        ContentBytes = oNode.nodeTypedValue
        Base64ToText = Byte_String(2, ContentBytes)
    End Function
    
    Public Function TextToBase64(ByVal Content)
        ' BASE64_Dencode a text content with BASE64.
        ' Creates the file with the Dencoded binary content
        Dim ADODB, ContentBytes, oXML, oNode, Text
        Set oXML = CreateObject("Msxml2.DOMDocument")
        Set oNode = oXML.CreateElement("base64")
        oNode.DataType = "bin.base64"
        ContentBytes = Byte_String(1, Content)
        oNode.nodeTypedValue = ContentBytes
        Text = oNode.text: Text = Replace(Text, vbLf, "")
        TextToBase64 = Text
    End Function
    
    ' 来自vbs吧老先辈的函数，十分感谢，这段代码对于我来说真是及时雨啊
    Public Function Byte_String(ByVal Mode, ByVal Data)
        Dim CharSetArr, ADO
        '字节数组与字符串相互转换
        'By:雷克斯.派
        'Mode参数 限定1或2,否则返回False 1=字符to字节/2=字节to字符
        'Data参数 Mode 为1时,请传入字符串类型/为2时,请传入字节数组类型
        Select Case Mode
            Case 1, 2
            CharSetArr = Array("UniCode", "UTF-8", "UniCode")
            Set ADO = CreateObject("ADODB.Stream")
            With ADO
                .Open
                If Mode = 1 Then .Type = 2 Else .Type = 1
                If Mode = 1 Then .WriteText Data Else .Write Data
                .Position = 0
                .Type = Mode
                Select Case Mode
                    Case 1
                    Byte_String = .Read '(.size)
                    .Position = 2
                    Case 2
                    Byte_String = .ReadText
                End Select
                .Close
            End With
            Set ADO = Nothing
            Case Else
            Byte_String = False
        End Select
    End Function
    
    Public Function Vigenere(ByVal Text, ByVal Password, ByVal MoveStep)
        ''' Vigenere Encrypt/Decrypt; Can only Encrypt/Decrypt chars which are in the variant "Letters"; Use "Compatible_Vigenere" To Encrypt/Decrypt All kinds of chars & strings'''
        On Error Resume Next
        Dim LLength, PLength, TLength, RLength, EnCodeCtr
        Dim Char, Move
        Dim i, j
        Const Letters = "aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ0123456789+=/" ' 按照BASE64码表，方便BASE64文本的加密
        LLength = Len(Letters)
        PLength = Len(Password)
        TLength = Len(Text)
        RLength = PLength mod TLength
        EnCodeCtr = ((PLength - RLength) / TLength)
        If RLength = 0 Then
            EnCodeCtr = EnCodeCtr - 1
        End If
        '    MsgBox CStr(EnCodeCtr)
        Execute "Dim Final(" & CStr(TLength-1) & ")"
        ' 加密方法：对Text里每个字符单独加密，一次直接加密为最终结果，不需要多次遍历
        For i = 1 To TLength
            Char = Mid(Text, i, 1)
            Move = InStr(1, Letters, Char)
            For j = i To EnCodeCtr * TLength + i Step TLength
                If j <= PLength Then Move = Move + InStr(1, Letters, Mid(Password, j, 1)) * MoveStep Else If j Mod PLength <> 0 Then Move = Move + InStr(1, Letters, Mid(Password, j Mod PLength, 1)) * MoveStep Else Move = Move + InStr(1, Letters, Mid(Password, PLength, 1)) * MoveStep
            Next
            Move = Move Mod LLength
            If Move <= 0 Then Move = LLength + Move
            Final(i-1) = Mid(Letters, Move, 1)
        Next
        Vigenere = Join(Final, "")
    End Function
    
    Public Function Compatible_Vigenere(ByVal Text, ByVal Password, ByVal Mode)
        ''' Mode: 1-Encrypt, -1-Decrypt; You can also use any non-negative integer for encrypt and its negative ones to decrypt '''
        Dim BASE64_Text
        If Mode > 0 Then
            BASE64_Text = TextToBase64(Text)
            Password = TextToBase64(Password)
            Compatible_Vigenere = Vigenere(BASE64_Text, Password, Mode)
        Else
            Password = TextToBase64(Password)
            BASE64_Text = Vigenere(Text, Password, Mode)
            Text = Base64ToText(BASE64_Text)
            Compatible_Vigenere = Text
        End If
    End Function
End Class

Class VbsJson
	'Author: Demon
	'Date: 2012/5/3
	'Website: http://demon.tw
	Private Whitespace, NumberRegex, StringChunk
	Private b, f, r, n, t
	
	Private Sub Class_Initialize
		Whitespace = " " & vbTab & vbCr & vbLf
		b = ChrW(8)
		f = vbFormFeed
		r = vbCr
		n = vbLf
		t = vbTab
		
		Set NumberRegex = New RegExp
		NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
		NumberRegex.Global = False
		NumberRegex.MultiLine = True
		NumberRegex.IgnoreCase = True
		
		Set StringChunk = New RegExp
		StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
		StringChunk.Global = False
		StringChunk.MultiLine = True
		StringChunk.IgnoreCase = True
	End Sub
	
	'Return a JSON string representation of a VBScript data structure
	'Supports the following objects and types
	'+-------------------+---------------+
	'| VBScript          | JSON          |
	'+===================+===============+
	'| Dictionary        | object        |
	'+-------------------+---------------+
	'| Array             | array         |
	'+-------------------+---------------+
	'| String            | string        |
	'+-------------------+---------------+
	'| Number            | number        |
	'+-------------------+---------------+
	'| True              | true          |
	'+-------------------+---------------+
	'| False             | false         |
	'+-------------------+---------------+
	'| Null              | null          |
	'+-------------------+---------------+
	Public Function Encode(ByRef obj)
		Dim buf, i, c, g
		Set buf = CreateObject("Scripting.Dictionary")
		Select Case VarType(obj)
			Case vbNull
			buf.Add buf.Count, "null"
			Case vbBoolean
			If obj Then
				buf.Add buf.Count, "true"
			Else
				buf.Add buf.Count, "false"
			End If
			Case vbInteger, vbLong, vbSingle, vbDouble
			buf.Add buf.Count, obj
			Case vbString
			buf.Add buf.Count, """"
			For i = 1 To Len(obj)
				c = Mid(obj, i, 1)
				Select Case c
					Case """" buf.Add buf.Count, "\"""
					Case "\"  buf.Add buf.Count, "\\"
					Case "/"  buf.Add buf.Count, "/"
					Case b    buf.Add buf.Count, "\b"
					Case f    buf.Add buf.Count, "\f"
					Case r    buf.Add buf.Count, "\r"
					Case n    buf.Add buf.Count, "\n"
					Case t    buf.Add buf.Count, "\t"
					Case Else
					If AscW(c) >= 0 And AscW(c) <= 31 Then
						c = Right("0" & Hex(AscW(c)), 2)
						buf.Add buf.Count, "\u00" & c
					Else
						buf.Add buf.Count, c
					End If
				End Select
			Next
			buf.Add buf.Count, """"
			Case vbArray + vbVariant
			g = True
			buf.Add buf.Count, "["
			For Each i In obj
				If g Then g = False Else buf.Add buf.Count, ","
				buf.Add buf.Count, Encode(i)
			Next
			buf.Add buf.Count, "]"
			Case vbObject
			If TypeName(obj) = "Dictionary" Then
				g = True
				buf.Add buf.Count, "{"
				For Each i In obj
					If g Then g = False Else buf.Add buf.Count, ","
					buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
				Next
				buf.Add buf.Count, "}"
			Else
				Err.Raise 8732,,"None dictionary object"
			End If
			Case Else
			buf.Add buf.Count, """" & CStr(obj) & """"
		End Select
		Encode = Join(buf.Items, "")
	End Function
	
	'Return the VBScript representation of ``str(``
	'Performs the following translations in decoding
	'+---------------+-------------------+
	'| JSON          | VBScript          |
	'+===============+===================+
	'| object        | Dictionary        |
	'+---------------+-------------------+
	'| array         | Array             |
	'+---------------+-------------------+
	'| string        | String            |
	'+---------------+-------------------+
	'| number        | Double            |
	'+---------------+-------------------+
	'| true          | True              |
	'+---------------+-------------------+
	'| false         | False             |
	'+---------------+-------------------+
	'| null          | Null              |
	'+---------------+-------------------+
	Public Function Decode(ByRef str)
		Dim idx
		idx = SkipWhitespace(str, 1)
		
		If Mid(str, idx, 1) = "{" Then
			Set Decode = ScanOnce(str, 1)
		Else
			Decode = ScanOnce(str, 1)
		End If
	End Function
	
	Private Function ScanOnce(ByRef str, ByRef idx)
		Dim c, ms
		
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "{" Then
			idx = idx + 1
			Set ScanOnce = ParseObject(str, idx)
			Exit Function
		ElseIf c = "[" Then
			idx = idx + 1
			ScanOnce = ParseArray(str, idx)
			Exit Function
		ElseIf c = """" Then
			idx = idx + 1
			ScanOnce = ParseString(str, idx)
			Exit Function
		ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ScanOnce = Null
			Exit Function
		ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ScanOnce = True
			Exit Function
		ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
			idx = idx + 5
			ScanOnce = False
			Exit Function
		End If
		
		Set ms = NumberRegex.Execute(Mid(str, idx))
		If ms.Count = 1 Then
			idx = idx + ms(0).Length
			ScanOnce = CDbl(ms(0))
			Exit Function
		End If
		
		Err.Raise 8732,,"No JSON object could be ScanOnced"
	End Function
	
	Private Function ParseObject(ByRef str, ByRef idx)
		Dim c, key, value
		Set ParseObject = CreateObject("Scripting.Dictionary")
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "}" Then
			idx = idx + 1
			Exit Function
		ElseIf c <> """" Then
			WScript.Echo "ParseObject: Error Out Of Loop"
			WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
			Err.Raise 8732,,"Expecting property name"
		End If
		
		idx = idx + 1
		
		Do
			key = ParseString(str, idx)
			
			idx = SkipWhitespace(str, idx)
			If Mid(str, idx, 1) <> ":" Then
				WScript.Echo "ParseObject: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Expecting : delimiter"
			End If
			
			idx = SkipWhitespace(str, idx + 1)
			If Mid(str, idx, 1) = "{" Then
				Set value = ScanOnce(str, idx)
			Else
				value = ScanOnce(str, idx)
			End If
			ParseObject.Add key, value
			
			idx = SkipWhitespace(str, idx)
			c = Mid(str, idx, 1)
			If c = "}" Then
				Exit Do
			ElseIf c <> "," Then
				WScript.Echo "ParseObject: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Expecting , delimiter"
			End If
			
			idx = SkipWhitespace(str, idx + 1)
			c = Mid(str, idx, 1)
			If c <> """" Then
				WScript.Echo "ParseObject: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Expecting property name"
			End If
			
			idx = idx + 1
		Loop
		
		idx = idx + 1
	End Function
	
	Private Function ParseArray(ByRef str, ByRef idx)
		Dim c, values, value
		Set values = CreateObject("Scripting.Dictionary")
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "]" Then
			ParseArray = values.Items
			idx = idx + 1
			Exit Function
		End If
		
		Do
			idx = SkipWhitespace(str, idx)
			If Mid(str, idx, 1) = "{" Then
				Set value = ScanOnce(str, idx)
			Else
				value = ScanOnce(str, idx)
			End If
			values.Add values.Count, value
			
			idx = SkipWhitespace(str, idx)
			c = Mid(str, idx, 1)
			If c = "]" Then
				Exit Do
			ElseIf c <> "," Then
				WScript.Echo "ParseArray: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Expecting , delimiter"
			End If
			
			idx = idx + 1
		Loop
		
		idx = idx + 1
		ParseArray = values.Items
	End Function
	
	Private Function ParseString(ByRef str, ByRef idx)
		Dim chunks, content, terminator, ms, esc, char
		Set chunks = CreateObject("Scripting.Dictionary")
		
		Do
			Set ms = StringChunk.Execute(Mid(str, idx))
			If ms.Count = 0 Then
				WScript.Echo "ParseString: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Unterminated string starting"
			End If
			
			content = ms(0).Submatches(0)
			terminator = ms(0).Submatches(1)
			If Len(content) > 0 Then
				chunks.Add chunks.Count, content
			End If
			
			idx = idx + ms(0).Length
			
			If terminator = """" Then
				Exit Do
			ElseIf terminator <> "\" Then
				WScript.Echo "ParseString: Error In Loop"
				WScript.Echo "c = """ & c & """, Asc(c) = " & CStr(Asc(c))
				Err.Raise 8732,,"Invalid control character"
			End If
			
			esc = Mid(str, idx, 1)
			
			If esc <> "u" Then
				Select Case esc
					Case """" char = """"
					Case "\"  char = "\"
					Case "/"  char = "/"
					Case "b"  char = b
					Case "f"  char = f
					Case "n"  char = n
					Case "r"  char = r
					Case "t"  char = t
					Case Else Err.Raise 8732,,"Invalid escape"
				End Select
				idx = idx + 1
			Else
				char = ChrW("&H" & Mid(str, idx + 1, 4))
				idx = idx + 5
			End If
			
			chunks.Add chunks.Count, char
		Loop
		
		ParseString = Join(chunks.Items, "")
	End Function
	
	Private Function SkipWhitespace(ByRef str, ByVal idx)
		Do While idx <= Len(str) And _
			InStr(Whitespace, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		SkipWhitespace = idx
	End Function
	
End Class