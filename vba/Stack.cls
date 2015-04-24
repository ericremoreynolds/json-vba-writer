VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const initialCapacity As Long = 32

Private totalCapacity As Long  ' Length of the buffer
Private curCount As Long    ' Length of the string value within the buffer

Private items

Private Sub Class_Initialize()
    totalCapacity = initialCapacity
    ReDim items(1 To totalCapacity)
    curCount = 0
End Sub

Public Sub Push(item)

    Dim newCount As Long
    newCount = curCount + 1
    
    If newCount > totalCapacity Then
        totalCapacity = totalCapacity * 2
        ReDim Preserve items(1 To totalCapacity)
    End If
    items(newCount) = item
    curCount = newCount
    
End Sub

Public Function Pop()
    
    If curCount = 0 Then
        Err.Raise 1234, Description:="No items on stack"
    End If
    
    Pop = items(curCount)
    curCount = curCount - 1
    
End Function

Public Property Get Top()
    If curCount = 0 Then
        Err.Raise 1234, Description:="No items on stack"
    End If
    Top = items(curCount)
End Property

Public Property Let Top(Value)
    If curCount = 0 Then
        Err.Raise 1234, Description:="No items on stack"
    End If
    items(curCount) = Value
End Property


