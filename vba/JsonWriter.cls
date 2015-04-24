VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "jsonwriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private json As New StringBuilder
Private needComma As New Stack

Private Sub Class_Initialize()
    needComma.Push False
End Sub

Function BeginArray()
    If needComma.Top Then json.Append ","
    json.Append "["
    needComma.Push False
End Function

Function EndArray()
    needComma.Pop
    json.Append "]"
    needComma.Top = True
End Function

Function BeginObject()
    If needComma.Top Then json.Append ","
    json.Append "{"
    needComma.Push False
End Function

Function EndObject()
    needComma.Pop
    json.Append "}"
    needComma.Top = True
End Function

Function WriteKey(k)
    WriteString CStr(k)
    json.Append ":"
    needComma.Top = False
End Function

Function WriteString(ByVal s As String)
    If needComma.Top Then json.Append ","
    json.Append """"
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\f")
    s = Replace(s, vbTab, "\t")
    json.Append s
    json.Append """"
    needComma.Top = True
End Function

Function WriteNumber(ByVal n As Double)
    If needComma.Top Then json.Append ","
    json.Append CStr(n)
    needComma.Top = True
End Function

Function WriteBool(ByVal b As Boolean)
    If needComma.Top Then json.Append ","
    If b Then
        json.Append "true"
    Else
        json.Append "false"
    End If
    needComma.Top = True
End Function

Function WriteKeyValue(k, v)
    WriteKey k
    WriteValue v
End Function

Function WriteValue(v)
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbDecimal, vbByte
            WriteNumber v
        Case vbString
            WriteString v
        Case vbBoolean
            WriteBool v
        Case Else
            WriteString CStr(v)
    End Select
End Function

Property Get Text() As String
    Text = json.Text
End Property

