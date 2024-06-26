VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMapping"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'
'  This is very similar to a Collection object.
'  However, it is more resiliant, requires a
'  key, and a specific order is not kept.
'
'  It is actually very similar to a
'  mapping in lpc (mud)
'
'  .Item() sets a new key, changes an old
'  key, and returns a requested value by
'  key (or Null if none is found)
'  Keys can also be retrieved (by number)
'
Private vData As New Collection
Private vKeys As New Collection



Public Property Get Count() As Long
   Count = vData.Count
End Property
Public Function GetItem(Key As Variant, Default As Variant) As Variant
   On Error GoTo ErrorHandle
   
   If VarType(vData(Key)) = vbObject Then
      Set GetItem = vData(Key)
   Else
      GetItem = vData(Key)
   End If
   
   Exit Function
   
   
ErrorHandle:
   If VarType(Default) = vbObject Then
      Set GetItem = Default
   Else
      GetItem = Default
   End If
End Function

Public Property Get Item(Key As Variant) As Variant
   On Error GoTo GetError
   
   If VarType(vData(Key)) = vbObject Then
     Set Item = vData(Key)
   Else
      Item = vData(Key)
   End If
   
   Exit Property
   
GetError:
   Item = Null
End Property
Public Property Set Item(Key As Variant, v As Variant)
   On Error Resume Next
   vData.Remove Key
   If Err > 0 Then vKeys.Add Key, Key
   vData.Add v, Key
End Property

Public Property Let Item(Key As Variant, v As Variant)
   On Error Resume Next
   vData.Remove Key
   If Err > 0 Then vKeys.Add Key, Key
   vData.Add v, Key
End Property


Public Function Key(ByVal i As Long) As Variant
   On Error Resume Next
   Key = vKeys(i)
   If Err > 0 Then Key = Null
End Function

Public Function Remove(Key As Variant) As Boolean
   On Error GoTo ErrorHandle
   
   vData.Remove Key
   vKeys.Remove Key
   Remove = True
   Exit Function

ErrorHandle:
   Remove = False
End Function

'
' Clears the mapping
'
Public Sub Clear()
   Set vData = New Collection
   Set vKeys = New Collection
End Sub

