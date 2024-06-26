VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFifo"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'
' A FIFO (First In First Out) buffer.
'
Private vData As New Collection

'
'  Adds an item to the FIFO
'
Public Sub Push(v As Variant)
   If vData.Count = 0 Then
      vData.Add v
   Else
      vData.Add v, , , vData.Count
   End If
End Sub

'
'  Clears all items from the FIFO
'
Public Sub Clear()
   Set vData = New Collection
End Sub

'
' Returns the number of items in the FIFO
'
Public Property Get Count() As Integer
   Count = vData.Count
End Property



'
'  Returns an item from the FIFO
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
'  Returns the next item in the FIFO but does
'  not remove it from the FIFO
'
Public Function Peek() As Variant
   If vData.Count = 0 Then
      Peek = Null
   Else
      If VarType(vData(1)) = vbObject Then
         Peek = vData(1)
      Else
         Peek = vData(1)
      End If
   End If
End Function


