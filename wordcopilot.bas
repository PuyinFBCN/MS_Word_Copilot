Attribute VB_Name = "Module2"
Private Const CP_UTF8 = 65001

  Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long
    
  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long


Sub copilot()
    Dim aitext As String
    Dim rng
    DoEvents
    aitext = generate_text2()
    aitext = URLDecode(Replace(aitext, "\n", Chr(10)))
    Set rng = Selection.Range
    rng.text = aitext
    
    'ActiveDocument.Content.InsertAfter aitext
End Sub

Function generate_text2() As String
    Dim request As Object
    Dim response As String
    Dim url As String
    Dim api_key As String
    Dim prompt As String
    Dim mytext As String
    
    prompt = URLEncode("Please continue to generate text after :   " & Right(Left(ActiveDocument.Range.text, Selection.Range.Start), 1000), True)
    url = "https://api.openai.com/v1/engines/davinci-codex/completions"
    url = "https://api.openai.com/v1/engines/text-davinci-003/completions"
    url = "https://api.openai.com/v1/completions"
    url = "https://api.openai.com/v1/chat/completions"
    api_key = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    
    Set request = CreateObject("MSXML2.XMLHTTP")
    
    request.Open "POST", url, False
    request.setRequestHeader "Content-Type", "application/json"
    request.setRequestHeader "Authorization", "Bearer " & api_key
    AI_input = ActiveDocument.BuiltInDocumentProperties("Comments")
    
    system_content = ""
    If AI_input <> "" Then
        system_content = "{""role"": ""system"", ""content"": """ & AI_input & """},"
    End If
    
    'request.send "{ ""model"": ""text-davinci-003"", ""prompt"": """ & prompt & """, ""max_tokens"": 500, ""temperature"": 0.7}"
    request.send "{ ""model"": ""gpt-3.5-turbo"", ""messages"": [" & system_content & "{""role"": ""user"", ""content"": """ & prompt & """}], ""max_tokens"": 500, ""temperature"": 0.7}"
    

    
    
    mytext = "{ ""model"": ""gpt-3.5-turbo"", ""messages"": \[\{""role"": ""user"", ""content"": ""Hello!""\}\]  }"
    response = request.responseText

    If InStr(response, "gpt-3.5-turbo-0301") Then
        mytext = Mid(response, InStr(response, """choices"":[{""message"":{""role"":""assistant"",""content"":""") + 53)
    Else
        mytext = Mid(response, InStr(response, """choices"":[{""message"":{""role"":""assistant"",""content"":""") + 53)
    
    End If
    mytext = Left(mytext, InStr(mytext, """}") - 1)
    generate_text2 = mytext


'curl https://api.openai.com/v1/chat/completions \
'  -H "Content-Type: application/json" \
'  -H "Authorization: Bearer $OPENAI_API_KEY" \
'  -d '{
'    "model": "gpt-3.5-turbo",
'    "messages": [{"role": "user", "content": "Hello!"}]
'  }'
    
End Function


Public Function UTF16To8(ByVal UTF16 As String) As String
Dim sBuffer As String
Dim lLength As Long
If UTF16 <> "" Then
    #If VBA7 Then
        lLength = WideCharToMultiByte(CP_UTF8, 0, CLngPtr(StrPtr(UTF16)), -1, 0, 0, 0, 0)
    #Else
        lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, 0, 0, 0, 0)
    #End If
    sBuffer = Space$(lLength)
    #If VBA7 Then
        lLength = WideCharToMultiByte(CP_UTF8, 0, CLngPtr(StrPtr(UTF16)), -1, CLngPtr(StrPtr(sBuffer)), LenB(sBuffer), 0, 0)
    #Else
        lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, StrPtr(sBuffer), LenB(sBuffer), 0, 0)
    #End If
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF16To8 = Left$(sBuffer, lLength - 1)
Else
    UTF16To8 = ""
End If
End Function


Public Function UTF8To16(ByVal UTF8 As String) As String
Dim sBuffer As String
Dim lLength As Long
If UTF8 <> "" Then
    #If VBA7 Then
        lLength = MultiByteToWideChar(CP_UTF8, 0, CLngPtr(StrPtr(UTF8)), -1, 0, 0)
    #Else
        lLength = MultiByteToWideChar(CP_UTF8, 0, StrPtr(UTF8), -1, 0, 0)
    #End If
    sBuffer = Space$(lLength * 2)
    #If VBA7 Then
        lLength = MultiByteToWideChar(CP_UTF8, 0, CLngPtr(StrPtr(UTF8)), -1, CLngPtr(StrPtr(sBuffer)), LenB(sBuffer) / 2)
    #Else
        lLength = MultiByteToWideChar(CP_UTF8, 0, StrPtr(UTF8), -1, StrPtr(sBuffer), LenB(sBuffer) / 2)
    #End If
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF8To16 = Left$(sBuffer, lLength - 1)
Else
    UTF8To16 = ""
End If
End Function


Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsSpace As Boolean = False, _
   Optional UTF8Encode As Boolean = True _
) As String

Dim StringValCopy As String: StringValCopy = IIf(UTF8Encode, UTF16To8(StringVal), StringVal)
Dim StringLen As Long: StringLen = Len(StringValCopy)

If StringLen > 0 Then
    ReDim Result(StringLen) As String
    Dim I As Long, CharCode As Integer
    Dim Char As String, Space As String

  If SpaceAsSpace Then Space = " " Else Space = "%20"

  For I = 1 To StringLen
    Char = Mid$(StringValCopy, I, 1)
    CharCode = Asc(Char)
    Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        Result(I) = Char
      Case 32
        Result(I) = Space
      Case 0 To 15
        Result(I) = "%0" & Hex(CharCode)
      Case Else
        Result(I) = "%" & Hex(CharCode)
    End Select
  Next I
  URLEncode = Join(Result, "")

End If
End Function



Public Function URLDecode( _
   StringVal As String, _
   Optional UTF8Decode As Boolean = False _
) As String
Dim StringLen As Long: StringLen = Len(StringVal)
If StringLen > 0 Then
    ReDim Result(StringLen) As String
    Dim I As Long, CharCode As Integer
    Dim Char As String
    For I = 1 To StringLen
        Char = Mid$(StringVal, I, 1)
        If Char = "%" Then
            On Error Resume Next
            Result(I) = Chr(CharCode)
            CharCode = CInt("&H" & Mid$(StringVal, I + 1, 2))
            If Err.Number <> 0 Then
                    Result(I) = Char
                    On Error GoTo 0
            Else
                On Error GoTo 0
            
                If CharCode > 0 Then
                    Result(I) = Chr(CharCode)
                    I = I + 2
                Else
                    Result(I) = Char
                End If
            End If
        Else
            Result(I) = Char
        End If
    Next I
    URLDecode = IIf(UTF8Decode, UTF8To16(Join(Result, "")), Join(Result, ""))
End If
End Function


