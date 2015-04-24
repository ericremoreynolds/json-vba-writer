VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const initialLength As Long = 32

Private totalLength As Long  ' Length of the buffer
Private curLength As Long    ' Length of the string value within the buffer
Private buffer As String     ' The buffer

Private Sub Class_Initialize()
  ' We set the buffer up to it's initial size and the string value ""
  totalLength = initialLength
  buffer = Space(totalLength)
  curLength = 0
End Sub

Public Sub Append(Text As String)

  Dim incLen As Long ' The length that the value will be increased by
  Dim newLen As Long ' The length of the value after being appended
  incLen = Len(Text)
  newLen = curLength + incLen

  ' Will the new value fit in the remaining free space within the current buffer
  If newLen < totalLength Then
    ' Buffer has room so just insert the new value
    Mid(buffer, curLength + 1, incLen) = Text
  Else
    ' Buffer does not have enough room so
    ' first calculate the new buffer size by doubling until its big enough
    ' then build the new buffer
    While totalLength < newLen
      totalLength = totalLength + totalLength
    Wend
    buffer = Left(buffer, curLength) & Text & Space(totalLength - newLen)
  End If
  curLength = newLen
End Sub

Public Property Get Length() As Integer
  Length = curLength
End Property

Public Property Get Text() As String
  Text = Left(buffer, curLength)
End Property

Public Sub Clear()
  totalLength = initialLength
  buffer = Space(totalLength)
  curLength = 0
End Sub

