VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLinkedList"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'
'  A simple doublely linked list class
'

Private m_head As clsListItem
Private m_tail As clsListItem
Private m_cur  As clsListItem

Private m_count As Long
'
'  Adds ItemData to the head of the linked list
'
Public Sub AddFirst(ItemData As Variant)
    Dim Item As New clsListItem

    If VarType(ItemData) = vbObject Then
       Set Item.ItemData = ItemData
    Else
       Item.ItemData = ItemData
    End If

    If (m_head Is Nothing) Then
        Set m_head = Item
        Set m_tail = Item
        Set m_cur = Item
    Else
        Set Item.NextItem = m_head
        Set Item.PrevItem = Nothing
        Set Item.NextItem.PrevItem = Item
        Set m_head = Item
    End If
    
    m_count = m_count + 1
End Sub

'
'  Adds ItemData to the end of the linked list.
'
Public Sub AddLast(ItemData As Variant)
    Dim Item As New clsListItem

    If VarType(ItemData) = vbObject Then
       Set Item.ItemData = ItemData
    Else
       Item.ItemData = ItemData
    End If

    If (m_tail Is Nothing) Then
        Set m_head = Item
        Set m_tail = Item
        Set m_cur = Item
    Else
        Set Item.PrevItem = m_tail
        Set Item.NextItem = Nothing
        Set Item.PrevItem.NextItem = Item
        Set m_tail = Item
    End If
    
    m_count = m_count + 1
End Sub

'
'  Returns the number of items in the linked list.
'
Property Get Count() As Long
   Count = m_count
End Property

'
'  Returns the current item.
'
Property Get CurrentItem() As Variant
   If m_cur Is Nothing Then
      CurrentItem = Null
   Else
      If VarType(m_cur.ItemData) = vbObject Then
         Set CurrentItem = m_cur.ItemData
      Else
         CurrentItem = m_cur.ItemData
      End If
   End If
End Property

'
'  Sets the current item.
'
Property Let CurrentItem(ItemData As Variant)
   If Not m_cur Is Nothing Then
      m_cur.ItemData = ItemData
   End If
End Property
Property Set CurrentItem(ItemData As Variant)
   If Not m_cur Is Nothing Then
      Set m_cur.ItemData = ItemData
   End If
End Property
'
'  Inserts ItemData after the current item in the list.
'
Public Sub InsertAfter(ItemData As Variant)
    Dim Item As New clsListItem

    If VarType(ItemData) = vbObject Then
       Set Item.ItemData = ItemData
    Else
       Item.ItemData = ItemData
    End If

    If (m_cur Is Nothing) Then
        Set m_head = Item
        Set m_tail = Item
        Set m_cur = Item
    Else
        Set Item.NextItem = m_cur.NextItem
        Set Item.PrevItem = m_cur
        Set m_cur.NextItem = Item
        'Add the following line.
        Set m_cur.NextItem.PrevItem = Item
        
        If (m_cur.NextItem Is Nothing) Then
            Set m_tail = m_cur
        End If
    End If
    
    m_count = m_count + 1
End Sub

'
'  Delete's the
'
Public Sub DeleteAll()
   Dim m As clsListItem
   Dim m2 As clsListItem
   
   m = m_head
   
   Do While Not (m Is Nothing)
      Set m2 = m.NextItem
      Set m.NextItem = Nothing
      Set m.PrevItem = Nothing
      Set m = m2
   Loop
   
   m_head = Nothing
   m_tail = Nothing
   m_cur = Nothing
   m_count = 0
End Sub
   
Public Sub DeleteCurrent()
    Dim tmp As clsListItem

    If (m_cur Is Nothing) Then
        Exit Sub
    End If

    If (m_cur.PrevItem Is Nothing) Then
        '
        ' Delete head of list
        '
        Set m_head = m_cur.NextItem
        If (m_head Is Nothing) Then
            '
            ' Also deleting tail, list becomes empty
            '
            Set m_tail = Nothing
            Set m_cur = Nothing
        Else
            Set m_head.PrevItem = Nothing
            Set m_cur = m_head
        End If
    ElseIf (m_cur.NextItem Is Nothing) Then
        '
        ' Deleting end of list
        '
        Set m_tail = m_cur.PrevItem
        If (m_tail Is Nothing) Then
            '
            ' Also deleting head, list becomes empty
            '
            Set m_head = Nothing
            Set m_cur = Nothing
        Else
            Set m_cur = m_tail
            Set m_cur.NextItem = Nothing
        End If
    Else
        '
        ' Delete somewhere inside of list
        '
        Set tmp = m_cur.NextItem
        Set m_cur.PrevItem.NextItem = m_cur.NextItem
        Set m_cur.NextItem.PrevItem = m_cur.PrevItem
        Set m_cur = tmp
    End If
    
    m_count = m_count - 1
End Sub
'
'  Return's the first item in the list.
'
Public Function FirstItem() As Variant
    If (m_head Is Nothing) Then
        FirstItem = Null
    Else
        If (VarType(m_head.ItemData) = vbObject) Then
            Set FirstItem = m_head.ItemData
        Else
            FirstItem = m_head.ItemData
        End If
        Set m_cur = m_head
    End If
End Function


'
'  Returns the next item in the list.
'
Public Function NextItem() As Variant
    If (m_cur Is Nothing) Then
        NextItem = Null
        Debug.Print "First Null"
    Else
        If (m_cur Is Nothing) Then
            NextItem = Null
        Else
            Set m_cur = m_cur.NextItem
            If (VarType(m_cur.ItemData) = vbObject) Then
                Set NextItem = m_cur.ItemData
            Else
                NextItem = m_cur.ItemData
            End If
        End If
    End If
End Function

'
' Returns the last item in the list.
'
Public Function LastItem() As Variant
    If (m_tail Is Nothing) Then
        LastItem = Null
    Else
        Set m_cur = m_tail
        If (VarType(m_cur.ItemData) = vbObject) Then
            Set LastItem = m_cur.ItemData
        Else
            LastItem = m_cur.ItemData
        End If
    End If
End Function

'
'  Returns the previous item in the list.
'
Public Function PrevItem() As Variant
    If (m_cur Is Nothing) Then
        PrevItem = Null
    Else
        If (m_cur.PrevItem Is Nothing) Then
            PrevItem = Null
        Else
            Set m_cur = m_cur.PrevItem
            If (VarType(m_cur.ItemData) = vbObject) Then
                Set PrevItem = m_cur.ItemData
            Else
                PrevItem = m_cur.ItemData
            End If
        End If
    End If
End Function

