<%
' =====================================================
' db.asp — Shared DB connection + JSON helpers
' auth.asp is included separately by each API page
' =====================================================

Const CONN_STRING = "Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=VMPortal;Integrated Security=SSPI;"

Function GetConnection()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STRING
    Set GetConnection = conn
End Function

Function JsonStr(s)
    If IsNull(s) Then
        JsonStr = "null"
    Else
        JsonStr = """" & Replace(Replace(Replace(Replace(CStr(s), "\", "\\"), """", "\"""), Chr(10), "\n"), Chr(13), "\r") & """"
    End If
End Function

Function JsonNum(n)
    If IsNull(n) Or n = "" Then JsonNum = "0" Else JsonNum = CStr(CLng(n))
End Function

Function FormatDateISO(d)
    If IsNull(d) Or d = "" Then
        FormatDateISO = "null"
    Else
        FormatDateISO = """" & Year(d) & "-" & Right("0"&Month(d),2) & "-" & Right("0"&Day(d),2) & _
                        "T" & Right("0"&Hour(d),2) & ":" & Right("0"&Minute(d),2) & ":" & Right("0"&Second(d),2) & "Z" & """"
    End If
End Function

Function RowToJSON(rs)
    Dim i, s
    s = "{"
    For i = 0 To rs.Fields.Count - 1
        If i > 0 Then s = s & ","
        Dim fname : fname = rs.Fields(i).Name
        Dim fval  : fval  = rs.Fields(i).Value
        Dim ftype : ftype = rs.Fields(i).Type
        s = s & JsonStr(fname) & ":"
        If IsNull(fval) Then
            s = s & "null"
        ElseIf ftype = 3 Or ftype = 2 Or ftype = 16 Or ftype = 17 Or ftype = 18 Or ftype = 19 Or ftype = 20 Or ftype = 21 Then
            s = s & CStr(CLng(fval))
        ElseIf ftype = 135 Then
            s = s & """" & Year(fval) & "-" & Right("0"&Month(fval),2) & "-" & Right("0"&Day(fval),2) & _
                    "T" & Right("0"&Hour(fval),2) & ":" & Right("0"&Minute(fval),2) & ":" & Right("0"&Second(fval),2) & "Z" & """"
        Else
            s = s & JsonStr(fval)
        End If
    Next
    s = s & "}"
    RowToJSON = s
End Function

Function RecordsetToJSON(rs)
    Dim s : s = "["
    Dim first : first = True
    Do While Not rs.EOF
        If Not first Then s = s & ","
        s = s & RowToJSON(rs)
        first = False
        rs.MoveNext
    Loop
    s = s & "]"
    RecordsetToJSON = s
End Function

Function ReadBody()
    Dim bytes : bytes = Request.TotalBytes
    If bytes > 0 Then
        Dim body : body = Request.BinaryRead(bytes)
        Dim stream
        Set stream = Server.CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write body
        stream.Position = 0
        stream.Type = 2
        stream.Charset = "utf-8"
        ReadBody = stream.ReadText
        stream.Close
        Set stream = Nothing
    Else
        ReadBody = ""
    End If
End Function

Function JsonGet(json, key)
    Dim pattern : pattern = """" & key & """" & ":"
    Dim pos : pos = InStr(json, pattern)
    If pos = 0 Then JsonGet = "" : Exit Function
    Dim valStart : valStart = pos + Len(pattern)
    Do While Mid(json, valStart, 1) = " " Or Mid(json, valStart, 1) = Chr(9)
        valStart = valStart + 1
    Loop
    Dim ch : ch = Mid(json, valStart, 1)
    If ch = """" Then
        Dim valEnd : valEnd = valStart + 1
        Do While valEnd <= Len(json)
            If Mid(json, valEnd, 1) = """" And Mid(json, valEnd-1, 1) <> "\" Then Exit Do
            valEnd = valEnd + 1
        Loop
        JsonGet = Mid(json, valStart+1, valEnd - valStart - 1)
    ElseIf ch = "n" Then
        JsonGet = ""
    Else
        Dim ve : ve = valStart
        Do While ve <= Len(json)
            Dim c : c = Mid(json, ve, 1)
            If c = "," Or c = "}" Or c = "]" Or c = " " Then Exit Do
            ve = ve + 1
        Loop
        JsonGet = Mid(json, valStart, ve - valStart)
    End If
End Function

Sub SendJSON(json)
    Response.ContentType = "application/json"
    Response.CharSet = "utf-8"
    Response.Write json
End Sub

Sub SendError(code, msg)
    Response.Status = code & " Error"
    Response.ContentType = "application/json"
    Response.Write "{""error"":" & JsonStr(msg) & "}"
End Sub
%>
