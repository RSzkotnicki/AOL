VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStack"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' A simple Stack class.


' Private Data
Private vData As New Collection


Public Sub Clear()
   Set vData = New Collection
End Sub


'
' Returns the number of items in the stack.
'
Property Get Count() As Long
   Count = vData.Count
End Property



'
'  Returns the value of the next item on the
'  stack without removing the item from the
'  stack.
'
Public Function Peek() As Variant
   If vData.Count = 0 Then
      Peek = Null
   Else
      If VarType(vData.Item(1)) = vbObject Then
         Set Peek = vData(1)
      Else
         Peek = vData(1)
      End If
   End If
End Function

'
'  Removes and returns the next item from
'  the stack.
'
Public Function Pop() As Variant
   If vData.Count = 0 Then
      Pop = Null
   Else
      If VarType(vData(1)) = vbObject Then
         Set Pop = vData(1)
      Else
         Pop = vData(1)
      End If
      vData.Remove 1
   End If
End Function




'
'  Adds an item to the Stack
'
Public Sub Push(v As Variant)
   If vData.Count = 0 Then
      vData.Add v
   Else
     vData.Add v, , 1
   End If
End Sub
