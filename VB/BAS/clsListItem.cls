VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListItem"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'  Used with the clsLinkedList class.

Private m_next As clsListItem
Private m_prev As clsListItem
Private m_data As Variant

Public Property Set ItemData(v As Variant)
   Set m_data = v
End Property

Public Property Get NextItem() As clsListItem
    Set NextItem = m_next
End Property

Public Property Set NextItem(Item As clsListItem)
    Set m_next = Item
End Property

Public Property Get PrevItem() As clsListItem
    Set PrevItem = m_prev
End Property

Public Property Set PrevItem(Item As clsListItem)
    Set m_prev = Item
End Property

Public Property Get ItemData() As Variant
    If (VarType(m_data) = vbObject) Then
        Set ItemData = m_data
    Else
        ItemData = m_data
    End If
End Property

Public Property Let ItemData(v As Variant)
    If (VarType(v) = vbObject) Then
        Set m_data = v
    Else
        m_data = v
    End If
End Property
