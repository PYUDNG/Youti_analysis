'VbsBeautifier 1.0 by Demon
'HTTP://demon.tw
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

'Call Import("VBS_JSON_By_Demon")

Dim F, Text, YAI, YPI, AnswerText

Dim DonePaper, UndoPaper, StopPaper, DoneExam, UndoExam, StopExam
Dim Data, Homework, Paper, Paper_name, Paper_id, Homework_id, Choice, Special
Dim PaperJson, FP
Set YAI = New YoutiAPI
Set UndoPaper = YAI.Youti_GetHomeworkList("homework", 1)
Set DonePaper = YAI.Youti_GetHomeworkList("homework", 2)
Set StopPaper = YAI.Youti_GetHomeworkList("homework", 3)
Set UndoExam = YAI.Youti_GetHomeworkList("exam", 1)
Set DoneExam = YAI.Youti_GetHomeworkList("exam", 2)
Set StopExam = YAI.Youti_GetHomeworkList("exam", 3)
'Msgbox DonePaper("data")(0)("homework")(0)("paper_name")
MsgBox "����ɵ���ҵ��" & CStr(UBound(DonePaper("data"))+1) & "��" & vbCrLf & _
       "δ��ɵ���ҵ��" & CStr(UBound(UndoPaper("data"))+1) & "��" & vbCrLf & _
       "�ѹ��ڵ���ҵ��" & CStr(UBound(StopPaper("data"))+1) & "��" & vbCrLf & _
       "����ɵĿ�����" & CStr(UBound(DoneExam("data"))+1) & "��" & vbCrLf & _
       "δ��ɵĿ�����" & CStr(UBound(UndoExam("data"))+1) & "��" & vbCrLf & _
       "�ѹ��ڵĿ�����" & CStr(UBound(StopExam("data"))+1) & "��", 64 + 2048, "����������ͳ��"

Set YPI = New YoutiPaperInformation
Choice = 1: Special = ""
For Each Data In UndoPaper("data")
	For Each Homework In Data("homework")
		Paper_id = Homework("paper_id")
		Homework_id = Homework("homework_id")
		Paper_name = Homework("paper_name")
		' ��ȡ����������Json
		Set Paper = YAI.Youti_GetPaper(Homework_id, Paper_id, Choice, Special)
		' ���ɴ��ı�
		YPI.LoadPaper(Paper)
		FP = SelfFolderPath & CStr(Paper_id) & "-" & Paper_name & " " & "�𰸽���.txt"
		FSO.CreateTextFile(FP, True).Write YPI.Paper_Anwser
		ws.Run """" & FP & """"
	Next
Next

WScript.Quit

'Set YPI = New YoutiPaperInformation
'YPI.LoadPaper(Text)
'AnswerText = YPI.Paper_Anwser
'Dim FP
'FP = SelfFolderPath & "�𰸽���.txt"
'FSO.CreateTextFile(FP, True).Write AnswerText
'ws.Run FP






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
		''' ���ļ���ȡ���ܵ��˺���Ϣ������[Account, Phone, Password]�������� '''
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
		''' ���ܴ����˺���Ϣ���ļ� '''
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


Class YoutiAPI
    Private Sub Class_Initialize()
    	' Init Variables
    	Host = "https://www.youti99.com/"
    	SelfFolderPath = FSO.GetParentFolderName(WScript.ScriptFullName)
    	If Right(SelfFolderPath, 1) <> "\" Then SelfFolderPath = SelfFolderPath & "\"
    	CertificateFilePath = SelfFolderPath & "Certificate.json"
    	AccountFilePath = SelfFolderPath & "MyAccount.json"
    	LogFilePath = SelfFolderPath & "Youti_Log.log"
    	Set VJ = New VbsJson
    	Set VG = New Vigenere
    	
    	' Import Logging Class "CLogger", For More Information, See ./CLogger/README.md
    	Dim CLoggerModel
    	CLoggerModel = SelfFolderPath & "CLogger\CLogger.vbs"
    	ExecuteGlobal FSO.OpenTextFile(CLoggerModel).ReadAll()
    	Set CL = New CLogger
    	CL.Debug = True
    	CL.IncludeTimestamp = True
    	CL.LogFile = LogFilePath
    	CL.LogToConsole = False
    	
    	' Get Certificate Token
    	If Read(CertificateFilePath) <> 0 Then Call Youti_Auto_Login(AccountFilePath): Call Save(CertificateFilePath)
	End Sub
	
	Private Sub Class_Terminate()
		Set VJ = Nothing
		Set VG = Nothing
		Set CL = Nothing
	End Sub
	
	Private MyAccount
	Private SelfFolderPath, AccountFilePath, CertificateFilePath, LogFilePath
	Private Host, sessionid, authorization
	Private VJ, VG, CL
	
	Public Function Youti_GetPaper(ByVal homework_id, ByVal paper_id, ByVal choice, ByVal special)
		''' ��ȡ����������JSON '''
		
		CL.LogInfo "Call Youti_GetPaper"
		
		Dim JSON, Response
		Set JSON = CreateObject("Scripting.Dictionary")
		If Homework_id	<> "" Then JSON("homework_id")	= homework_id
		If Paper_id		<> "" Then JSON("paper_id")		= paper_id
		If choice		<> "" Then JSON("choice")		= choice
		If special		<> "" Then JSON("special")		= special
		Set Response = YoutiApiRequest("api/homework/v1/student/dw/do", JSON)
		Set Youti_GetPaper = VJ.Decode(Response.ResponseText)
	End Function
	
	Public Function Youti_GetHomeworkList(ByVal PaperType, ByVal Status)
		''' ��ȡ��������ҵ/�����б� '''
		' ����: 
		' 		PaperType:
		' 			"homework" - ��ҵ
		' 			"exam" - ����
		' 		Status: 
		' 			1 - δ���
		' 			2 - �����
		' 			3 - �ѹ���
		
		CL.LogInfo "Call Youti_GetHomeworkList"
		
		' �������
		Const AllowTypes = "homework|exam"
		PaperType = LCase(Trim(PaperType))
		If InStr(1, AllowTypes, PaperType) = 0 Then Err.Raise 9001, """" & PaperType & """ is not a valid PaperType"
		
		' ����YoutiApi
		Dim Response, ResponseJson
		Set Response = YoutiApiRequest("api/homework/v1/student/homework/list/" & CStr(Status) & "?cat=" & PaperType, Null)
		Set ResponseJson = VJ.Decode(Response.responseText)
		
		' �ж�Token�Ƿ���Ч
		Select Case ResponseJson("code")
		Case 200 ' һ����������������
		    Set Youti_GetHomeworkList = ResponseJson
		Case Else ' TokenʧЧ���磺600�����˺�����һ���ط���½�����ٴλ�ȡToken������
			CL.LogDebug "Token Expired, Code " & CStr(ResponseJson("code")) & ""
		    Call Youti_Auto_Login(AccountFilePath)
		    Set Youti_GetHomeworkList = Youti_GetHomeworkList(PaperType, Status)
		    Call Save(CertificateFilePath)
		End Select
	End Function
	
	Private Function Youti_Auto_Login(ByVal AccountFilePath)
		''' �Զ���ȡ���ܵ��˺���Ϣ������Youti_Login��ȡ�û�sessionid��authorization '''
		
		CL.LogInfo "Call Youti_Auto_Login"
		
		Dim OAS, AccountInformation
		Set OAS = New AccountSaver
		AccountInformation = OAS.ReadAccount(AccountFilePath)
		
		If VarType(AccountInformation) = 2 Then
			MsgBox "����û�е�¼���������õ�¼���û������ֻ��ź����룡", 64, WScript.ScriptName
			If Not FSO.FileExists(SelfFolderPath & "�˺�������ܴ���.vbs") Then 
				MsgBox "��ʧ [�˺�������ܴ���.vbs]�����������ش��ļ������ڱ�����ͬһĿ¼�£�"
				Youti_Auto_Login = -1
				Exit Function
			End If
			ws.Run "wscript.exe """ & SelfFolderPath & "�˺�������ܴ���.vbs" & """ uacHidden", 1, True
			Youti_Auto_Login = Youti_Auto_Login(AccountFilePath)
			Exit Function
		End If
		
		Call Youti_Login(AccountInformation(0), AccountInformation(1), AccountInformation(2))
	End Function
	
	Private Function Youti_Login(ByVal Account, ByVal Phone, ByVal Password)
		''' ========= �˺�����Ҫ��ʽ���ã���Ϊ�˺���Ϣ��ȫ����ʹ��Youti_Auto_Login�������� ======== '''
		''' ��������¼����ȡ�û�sessionid��authorization '''
		
		CL.LogInfo "Call Youti_Login"
		
		Dim HTTP
		Dim cookies, cookie
		
		' ����Youti_Login_Api��ȡResponse
	    Set HTTP = Youti_Login_Api(Account, Phone, Password)
	    
	    ' ��ȡcookies
	    cookies = Split(HTTP.getResponseHeader("Set-Cookie"), ";")
	    
	    ' ��cookies�����ȡsessionid
	    For Each cookie In cookies
	        If Split(cookie, "=")(0) = "sessionid" Then
	            sessionid = Split(cookie, "=")(1)
	            Exit For
	        End If
	    Next
	    
	    ' ��ȡauthorization
	    authorization = VJ.Decode(HTTP.responseText)("data")("authorization")
	End Function
	
	Private Function Youti_Login_Api(ByVal Account, ByVal Phone, ByVal Password)
		''' ��������¼(API����Youti_Login�ⲻҪֱ�ӵ���) '''
		' ����ֵ: HTTP����
		
		CL.LogInfo "Call Youti_Login_Api"
		
		Dim JSON, Response, ResponseJson
		Set JSON = CreateObject("Scripting.Dictionary")
		JSON.Add "account", CStr(Account)
		JSON.Add "phone", CStr(Phone)
		JSON.Add "passwd", CStr(Password)
		Set Youti_Login_Api = YoutiApiRequest("api/edu/v1/student/login", JSON)
	End Function
	
	Private Function YoutiApiRequest(ByVal API, ByVal Json)
	    ''' ������������ʽ��API���� '''
	    ' ������
	    '		API: API��ַ����������������
	    '		Json: Post��Json��JsonText,�����Json��Ӧ��VbsJson��;�����Null����Get����
	    ' ����ֵ: HTTP����
	    
		CL.LogInfo "Call YoutiApiRequest"
	    
	    Dim HTTP
        Set HTTP = CreateObject("Msxml2.ServerXMLHTTP")
        
	    ' ��ʽ������
	    Dim URL, JsonText
	    URL = LCase(Trim(API))
	    If Left(URL, 8) <> Host Then URL = Host & URL
	    If VarType(Json) <> 8 Then JsonText = VJ.Encode(Json) Else JsonText = Json
	    
	    ' ����HEADER
	    If IsNull(Json) Then HTTP.open "get", URL, False Else HTTP.open "post", URL, False
        HTTP.setRequestHeader "Accept", "application/json, text/plain, */*"
        HTTP.setRequestHeader "Accept-Language", "zh-CN"
        HTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) youti/2.3.1 Chrome/69.0.3497.128 Electron/4.2.12 Safari/537.36"
        If sessionid <> "" Then HTTP.setRequestHeader "cookie", "sessionid=" & sessionid
        If authorization <> "" Then HTTP.setRequestHeader "Authorization", authorization
	    
	    ' ��������
	    If IsNull(Json) Then HTTP.send() Else HTTP.send JsonText
	    
	    ' ���ؽ��
	    Set YoutiApiRequest = HTTP
	End Function
	
	Private Function Save(ByVal FP)
		
		CL.LogInfo "Call Save"
		
		Dim Json, P
		Set Json = CreateObject("Scripting.Dictionary")
		P = CreateRandomizedText(32)
		Json("Encrypted_SessionID") = VG.Compatible_Vigenere(sessionid, P, 1)
		Json("Encrypted_Authorization") = VG.Compatible_Vigenere(authorization, P, 1)
		Json("Password") = P
		FSO.CreateTextFile(FP, True).Write VJ.Encode(Json)
	End Function
	
	Private Function Read(ByVal FP)
	    ''' ���ļ��ж�ȡƾ�� '''
	    ' ����ֵ: 0 - �ɹ�; -1 - �ļ�������
	    
	    CL.LogInfo "Call Read"
	    
	    If Not FSO.FileExists(FP) Then
	        Read = -1
	        Exit Function
	    End If
	    
		Dim Json, P
		Set Json = VJ.Decode(FSO.OpenTextFile(FP).ReadAll())
		P = Json("Password")
		sessionid = VG.Compatible_Vigenere(Json("Encrypted_SessionID"), P, -1)
		authorization = VG.Compatible_Vigenere(Json("Encrypted_Authorization"), P, -1)
		
'		' ȥ���ӽ������ɵķǷ��ַ����е�СBUG��
'		Dim re
'		Set re = New RegExp
'		re.Pattern = "[^a-zA-Z0-9]"
'		sessionid = re.Replace(sessionid, "")
'		authorization = re.Replace(authorization, "")
		
		'FSO.CreateTextFile(SelfFolderPath & "sa.txt", True).WriteLine "authorization = " & sessionid & vbCrLf & "authorization = " & authorization
	End Function
End Class

Class YoutiPaperInformation
	Private Sub Class_Initialize()
		Set JSON = New VbsJson
	End Sub
	
	Private Sub Class_Terminate()
		' Do Nothing
	End Sub
	
	Private JSON, Paper
	Private PID, HID, Mid, PNM, TIM, VER, PCT, CID
	Private AnswerText
	
	Public Function LoadPaper(ByVal PaperJson)
		''' ���������������Ծ� '''
		If VarType(PaperJson) = 8 Then Set Paper = JSON.Decode(PaperJson) Else Set Paper = PaperJson
	    Call GetPaperInformation()
	    AnswerText = GetPaperAnswer()
	End Function
	
	Private Function GetPaperInformation()
		''' ͨ���Ծ�Json��ȡ�Ծ������Ϣ '''
		PID = Paper("data")("paper_id")
		HID = Paper("data")("homework_id")
		PNM = Paper("data")("data")("paper_name")
		Mid = Paper("data")("data")("manager_id")
		TIM = Paper("data")("data")("time")
		VER = Paper("data")("data")("version")
		PCT = Paper("data")("data")("cat")
		CID = Paper("data")("data")("controller_id")
	End Function
	
	Property Get Paper_id()
		Paper_id = PID
	End Property
	
	Property Get Homework_id()
		Homework_id = HID
	End Property
	
	Property Get Paper_name()
		Paper_name = PNM
	End Property
	
	Property Get Manager_id()
		Manager_id = Mid
	End Property
	
	Property Get Time()
		Time = TIM
	End Property
	
	Property Get Version()
		Version = VER
	End Property
	
	Property Get Cat()
		Cat = PCT
	End Property
	
	Property Get Controller_id()
		Controller_id = CID
	End Property
	
	Property Get Paper_Anwser()
		Paper_Anwser = AnswerText
	End Property
	
	Private Function GetPaperAnswer()
		Dim i
		i = 0
		Dim Questions, Question, Infos, QuestionType, Info, Items, Item
		Dim Answers, Answer, AnswerSort, AnswerForNewItem, CorrectAnswer
		Dim DisplayText()
		ReDim DisplayText(0)
		AnswerForNewItem = True
		AnswerSort =  - 1
		
		' ÿһ�����⣨���飩
		Questions = Paper("data")("data")("questions")
		For Each Question In Questions
			' ��Ӵ������
			ReDim Preserve DisplayText(i)
			DisplayText(i) = Question("content")
			If i > 0 Then DisplayText(i) = vbCrLf & vbCrLf & String(200, "-") & vbCrLf & vbCrLf & DisplayText(i)
			i = i + 1
			' ÿһ��С��
			Infos = Question("infos")
			For Each Info In Infos
				' ���С�����
				ReDim Preserve DisplayText(i)
				DisplayText(i) = vbCrLf & Info("content")
				i = i + 1
				' ������Ŀ����
				QuestionType = Info("type")
				' ÿһС��
				Items = Info("items")
				For Each Item In Items
					Answers = Item("answers")
					ReDim Preserve DisplayText(i)
					DisplayText(i) = Item("content")
					i = i + 1
					' ÿһ��ѡ��/��
					For Each Answer In Answers
						' �жϵ�ǰ�Ĵ���һ������Ĵ𰸻�����һ�������һ����ȷ��
						AnswerForNewItem = (AnswerSort <> Answer("sort"))
						AnswerSort = Answer("sort")
						Select Case QuestionType
							Case 1 ' ѡ����
							If Answer("is_right") Then CorrectAnswer = Answer("content") Else CorrectAnswer = ""
							Case 3 ' �����¼
							CorrectAnswer = Answer("content")
							Case 4 ' ��ͷ�ش�����
							CorrectAnswer = Answer("content")
							Case 5 ' ת��
							CorrectAnswer = Answer("content")
							Case 6 ' �ʶ�
							CorrectAnswer = Answer("content")
							Case Else ' �����ݲ�֧�֣�û�н�����
							CorrectAnswer = "����Ŀ�����ݲ�֧�ֽ����𰸣�"
						End Select
						' �����ı���ʾ
						If AnswerForNewItem And CorrectAnswer <> "" Then
							ReDim Preserve DisplayText(i)
							DisplayText(i) = "����𰸣�" & CorrectAnswer
							i = i + 1
						ElseIf CorrectAnswer <> "" Then
							DisplayText(i - 1) = DisplayText(i - 1) & " �� " & CorrectAnswer
						End If
					Next
					AnswerSort =  - 1
				Next
			Next
		Next
		GetPaperAnswer = Join(DisplayText, vbCrLf)
	End Function
End Class

Class Vigenere
	' This Class is written by: PY_DNG                / �������ߣ�PY_DNG
	' Using is free, but you must not clear this mark / ��������ʹ�ã������뱣���⼸��ע��
	' Prohibited For commercial use and illegal use   / ��ֹ������ҵ�Լ��Ƿ���;
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
    
    ' ����vbs�����ȱ��ĺ�����ʮ�ָ�л����δ����������˵���Ǽ�ʱ�갡
    Public Function Byte_String(ByVal Mode, ByVal Data)
        Dim CharSetArr, ADO
        '�ֽ��������ַ����໥ת��
        'By:�׿�˹.��
        'Mode���� �޶�1��2,���򷵻�False 1=�ַ�to�ֽ�/2=�ֽ�to�ַ�
        'Data���� Mode Ϊ1ʱ,�봫���ַ�������/Ϊ2ʱ,�봫���ֽ���������
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
        Const Letters = "aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ0123456789+=/" ' ����BASE64�������BASE64�ı��ļ���
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
        ' ���ܷ�������Text��ÿ���ַ��������ܣ�һ��ֱ�Ӽ���Ϊ���ս��������Ҫ��α���
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
		SA.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac" & Args, "", "runas", 1
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

Function FileInput(ByVal StandardExt, ByVal Filter)
	''' ��ȡ�ļ����룬����������ļ��б��Զ��ж��Ƿ����Ϸ��ļ����룬û�о��ֶ�ѡ���ļ��������в�����һλ����GetUAC����ʹ�ã�����Ϊ������ļ�·�� '''
	''' ������StandardExt - Ҫ�����չ�������û��ѡ�����չ�����ļ��ͽ���ȷ�ϣ�ע���Ϸŵ��ļ����������չ��ȷ�ϣ���֧�ֶ��������չ����д��Ϊ".Ext1|.Ext2"��Array(".Ext1", ".Ext2")�����ջ���".*"����������չ�� '''
	'''       Filter - ���˷�ʽ��'''
	'''           0. �������д�����ļ� '''
	'''           1. ֻ������չ����ȷ���ļ� '''
	Dim FP, Ext, SI, SO
	Dim i, j
	Set SI = WScript.StdIn
	Set SO = WScript.StdOut
	If IsArray(StandardExt) Then StandardExt = Join(StandardExt, "|")
	StandardExt = LCase(StandardExt)
	If WScript.Arguments.Count = 1 Then
		ReDim FP(0)
		Do
			FP(0) = ws.Exec("mshta vbscript:""<input type=file id=f><script>f.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(f.value)[close()];</script>""").StdOut.ReadAll
			Ext = "." & LCase(FSO.GetExtensionName(FP(0)))
			If FP(0) = "" Then FileInput = Empty
			Exit Function
			If Ext <> "." And InStr(1, StandardExt, Ext) <> 0 Then Exit Do
			If InStr(1, StandardExt, ".*") <> 0 Or StandardExt = "" Then Exit Do
			Select Case Filter
				Case 0
				SO.Write "��ע�⣡��ѡ����ļ���չ�������ڱ�׼��չ��""" & StandardExt & """�У���ȷ����û��ѡ���ļ������ȷ�ϼ�������ֱ�Ӱ��»س��������Ҫѡ�������ļ������������ַ����ٰ��»س�: "
				If SI.ReadLine = "" Then Exit Do
				Case 1
				SO.WriteLine "ѡ���ļ�ֻ��ѡ����չ��Ϊ""" & StandardExt & """�е���չ��֮һ���ļ���������ѡ��: "
			End Select
		Loop
	Else
		j =  - 1
		For i = 1 To WScript.Arguments.Count - 1
			Ext = "." & LCase(FSO.GetExtensionName(WScript.Arguments(i)))
			If (Filter = 0 Or InStr(1, StandardExt, Ext) <> 0) And FSO.FileExists(WScript.Arguments(i)) Then
				j = j + 1
				If j = 0 Then ReDim FP(j) Else ReDim Preserve FP(j)
				FP(j) = WScript.Arguments(i)
			End If
		Next
		If j =  - 1 Then FileInput = Empty
		Exit Function
	End If
	FileInput = FP
End Function

Function Import(ByVal Parameter)
	''' ��������ģ�飬����Python��from FP import *����ͬ���ǣ�������ֻ�ᵼ��sub��function��class����������������ᵼ�� '''
	' ����ģ��ʽ�Զ�ִ��ģ���е�ImportExecute����������еĻ��������е�ImportExecute������������ᱻ����
	' ����Parameter�����������ַ���ģ��·����Ҳ���Խ���ImportInfoVariant��������Դ��ݲ���
	'On Error Resume Next
	Select Case VarType(Parameter)
		' ����ͬ���͵Ĳ���
		' ���۲�����ʲô���ͣ����յĸ�ʽӦ����Ϊ��
		' - ImportInfos: ���ݸ�IE�Ļ�����Ϣ
		' ����һ������
		Case 8 ' �ַ���������·��
		ImportInfos.ModelFullPath = Parameter
		Case 9 ' ���󣬴���ImportInfoVariant�����
		' Do Nothing
	End Select
	If Not FSO.FileExists(ImportInfos.ModelFullPath) Then ImportInfos.ModelFullPath = ImportInfos.ModelFullPath & ".vbs"
	If Not FSO.FileExists(ImportInfos.ModelFullPath) Then
		Import =  - 1
		Exit Function
	End If
	Dim CodeAll
	Dim re
	Dim Funcs, IECode, Code
	Dim OldImportInfos
	CodeAll = FSO.OpenTextFile(ImportInfos.ModelFullPath).ReadAll()
	' ��ʼ��������ʽ
	Set re = New RegExp
	re.Global = True
	re.IgnoreCase = True
	re.Multiline = True
	' ƥ�亯�����
	re.Pattern = "^(function|sub|class) +.+\(.*\)(.*\s*)*?^end +(function|sub|class)"
	Set Funcs = re.Execute(CodeAll)
	' ƥ��ImportExecute
	re.Pattern = "^(function|sub) +ImportExecute\(.*\)(.*\s*)*?^end +(function|sub)"
	Set IECode = re.Execute(CodeAll)
	If IECode.Count > 0 Then IECode = IECode(IECode.Count - 1) Else IECode = ""
	' ȥ��IECode��Function/Sub����
	re.Pattern = "^(function|sub) +ImportExecute\(.*\)"
	IECode = re.Replace(IECode, "")
	re.Pattern = "^End +(function|sub)"
	IECode = re.Replace(IECode, "")
	' ִ�����е��뺯��
	For Each Code In Funcs
		If Code <> IECode Then ExecuteGlobal Code
	Next
	' ִ��ImportExecute
	ExecuteGlobal IECode
	' �ָ�ImportInfos����Ϣ
	ImportInfos.ModelFullPath = WScript.ScriptFullName
	Import = 0
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