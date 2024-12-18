Class VbsJson
    'Author: Demon
    'Date: 2012/5/3
    'Website: https://demon.tw/my-work/vbs-json.html
    'License: CC BY-NC-SA 2.5 CN
    Private Whitespace, BackSpace, NumberRegex, StringChunk
    
    Private Sub Class_Initialize
        Whitespace = " " & vbTab & vbCr & vbLf
        BackSpace = ChrW(8)
        
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
        Dim buffer, position, character, item, firstItem
        Set buffer = CreateObject("Scripting.Dictionary")
        Select Case VarType(obj)
            Case vbNull
                buffer.Add buffer.Count, "null"
            Case vbBoolean
                If obj Then
                    buffer.Add buffer.Count, "true"
                Else
                    buffer.Add buffer.Count, "false"
                End If
            Case vbInteger, vbLong, vbSingle, vbDouble
                buffer.Add buffer.Count, obj
            Case vbString
                buffer.Add buffer.Count, """"
                For position = 1 To Len(obj)
                    character = Mid(obj, position, 1)
                    Select Case character
                        Case """"        buffer.Add buffer.Count, "\"""
                        Case "\"         buffer.Add buffer.Count, "\\"
                        Case "/"         buffer.Add buffer.Count, "/"
                        Case vbFormFeed  buffer.Add buffer.Count, "\f"
                        Case BackSpace   buffer.Add buffer.Count, "\b"
                        Case vbCr        buffer.Add buffer.Count, "\r"
                        Case vbLf        buffer.Add buffer.Count, "\n"
                        Case vbTab       buffer.Add buffer.Count, "\t"
                        Case Else
                            If AscW(character) >= 0 And AscW(character) <= 31 Then
                                character = Right("0" & Hex(AscW(character)), 2)
                                buffer.Add buffer.Count, "\u00" & character
                            Else
                                buffer.Add buffer.Count, character
                            End If
                    End Select
                Next
                buffer.Add buffer.Count, """"
            Case vbArray + vbVariant
                firstItem = True
                buffer.Add buffer.Count, "["
                For Each item In obj
                    If firstItem Then firstItem = False Else buffer.Add buffer.Count, ","
                    buffer.Add buffer.Count, Encode(item)
                Next
                buffer.Add buffer.Count, "]"
            Case vbObject
                If TypeName(obj) = "Dictionary" Then
                    firstItem = True
                    buffer.Add buffer.Count, "{"
                    For Each key In obj
                        If firstItem Then firstItem = False Else buffer.Add buffer.Count, ","
                        buffer.Add buffer.Count, """" & key & """" & ":" & Encode(obj(key))
                    Next
                    buffer.Add buffer.Count, "}"
                Else
                    Err.Raise 8732,,"None dictionary object"
                End If
            Case Else
                buffer.Add buffer.Count, """" & CStr(obj) & """"
        End Select
        Encode = Join(buffer.Items, "")
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
        Dim index
        index = SkipWhitespace(str, 1)

        If Mid(str, index, 1) = "{" Then
            Set Decode = ScanOnce(str, 1)
        Else
            Decode = ScanOnce(str, 1)
        End If
    End Function
    
    Private Function ScanOnce(ByRef str, ByRef index)
        Dim character, matchedString

        index = SkipWhitespace(str, index)
        character = Mid(str, index, 1)

        If character = "{" Then
            index = index + 1
            Set ScanOnce = ParseObject(str, index)
            Exit Function
        ElseIf character = "[" Then
            index = index + 1
            ScanOnce = ParseArray(str, index)
            Exit Function
        ElseIf character = """" Then
            index = index + 1
            ScanOnce = ParseString(str, index)
            Exit Function
        ElseIf character = "n" And StrComp("null", Mid(str, index, 4)) = 0 Then
            index = index + 4
            ScanOnce = Null
            Exit Function
        ElseIf character = "t" And StrComp("true", Mid(str, index, 4)) = 0 Then
            index = index + 4
            ScanOnce = True
            Exit Function
        ElseIf character = "f" And StrComp("false", Mid(str, index, 5)) = 0 Then
            index = index + 5
            ScanOnce = False
            Exit Function
        End If
        
        Set matchedString = NumberRegex.Execute(Mid(str, index))
        If matchedString.Count = 1 Then
            index = index + matchedString(0).Length
            ScanOnce = CDbl(matchedString(0))
            Exit Function
        End If
        
        Err.Raise 8732,,"No JSON object could be ScanOnced"
    End Function

    Private Function ParseObject(ByRef str, ByRef index)
        Dim character, key, value
        Set ParseObject = CreateObject("Scripting.Dictionary")
        index = SkipWhitespace(str, index)
        character = Mid(str, index, 1)
        
        If character = "}" Then
            index = index + 1
            Exit Function
        ElseIf character <> """" Then
            Err.Raise 8732,,"Expecting property name"
        End If

        index = index + 1
        
        Do
            key = ParseString(str, index)

            index = SkipWhitespace(str, index)
            If Mid(str, index, 1) <> ":" Then
                Err.Raise 8732,,"Expecting : delimiter"
            End If

            index = SkipWhitespace(str, index + 1)
            If Mid(str, index, 1) = "{" Then
                Set value = ScanOnce(str, index)
            Else
                value = ScanOnce(str, index)
            End If
            ParseObject.Add key, value

            index = SkipWhitespace(str, index)
            character = Mid(str, index, 1)
            If character = "}" Then
                Exit Do
            ElseIf character <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            index = SkipWhitespace(str, index + 1)
            character = Mid(str, index, 1)
            If character <> """" Then
                Err.Raise 8732,,"Expecting property name"
            End If

            index = index + 1
        Loop

        index = index + 1
    End Function
    
    Private Function ParseArray(ByRef str, ByRef index)
        Dim character, values, value
        Set values = CreateObject("Scripting.Dictionary")
        index = SkipWhitespace(str, index)
        character = Mid(str, index, 1)

        If character = "]" Then
            index = index + 1
            ParseArray = values.Items
            Exit Function
        End If

        Do
            index = SkipWhitespace(str, index)
            If Mid(str, index, 1) = "{" Then
                Set value = ScanOnce(str, index)
            Else
                value = ScanOnce(str, index)
            End If
            values.Add values.Count, value

            index = SkipWhitespace(str, index)
            character = Mid(str, index, 1)
            If character = "]" Then
                Exit Do
            ElseIf character <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            index = index + 1
        Loop

        index = index + 1
        ParseArray = values.Items
    End Function
    
    Private Function ParseString(ByRef str, ByRef index)
        Dim chunks, content, terminator, matchedString, escapedCharacter, char
        Set chunks = CreateObject("Scripting.Dictionary")

        Do
            Set matchedString = StringChunk.Execute(Mid(str, index))
            If matchedString.Count = 0 Then
                Err.Raise 8732,,"Unterminated string starting"
            End If
            
            content = matchedString(0).Submatches(0)
            terminator = matchedString(0).Submatches(1)
            If Len(content) > 0 Then
                chunks.Add chunks.Count, content
            End If
            
            index = index + matchedString(0).Length
            
            If terminator = """" Then
                Exit Do
            ElseIf terminator <> "\" Then
                Err.Raise 8732,,"Invalid control character"
            End If
            
            escapedCharacter = Mid(str, index, 1)

            If escapedCharacter <> "u" Then
                Select Case escapedCharacter
                    Case """" char = """"
                    Case "\"  char = "\"
                    Case "/"  char = "/"
                    Case "b"  char = BackSpace
                    Case "f"  char = vbFormFeed
                    Case "n"  char = vbNl
                    Case "r"  char = vbCr
                    Case "t"  char = vbTab
                    Case Else Err.Raise 8732,,"Invalid escape"
                End Select
                index = index + 1
            Else
                char = ChrW("&H" & Mid(str, index + 1, 4))
                index = index + 5
            End If

            chunks.Add chunks.Count, char
        Loop

        ParseString = Join(chunks.Items, "")
    End Function

    Private Function SkipWhitespace(ByRef str, ByVal index)
        Do While index <= Len(str) And _
            InStr(Whitespace, Mid(str, index, 1)) > 0
            index = index + 1
        Loop
        SkipWhitespace = index
    End Function

End Class
