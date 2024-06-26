VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLastFiles"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'
'  An object to maintain the last files
'  opened by a program
'

Private LastFiles As New Collection
Private Num As Integer

Public Sub Add(file As String)
   If LastFiles.Count = 0 Then
      LastFiles.Add file, UCase(file)
   Else
      On Error Resume Next
      LastFiles.Remove UCase(file)
      LastFiles.Add file, UCase(file), 1
      If LastFiles.Count > Num Then
         LastFiles.Remove LastFiles.Count
      End If
   End If
End Sub

'
'  Clears all files from the list.
'
Public Sub Clear()
   Do While LastFiles.Count > 0
      LastFiles.Remove 1
   Loop
End Sub

'
'  Returns the number of files in the list.
'
Public Property Get Count() As Long
   Count = LastFiles.Count
End Property

'
'
' Returns the nth item from the list
'
Public Property Get Item(i As Integer) As String
   On Error GoTo ItemError
   Item = LastFiles(i)
   Exit Property
   
ItemError:
   Item = ""
End Property

Public Sub Load(Optional appname As Variant)
   Dim v As Variant
   Dim i As Integer
   Dim j As Integer
   Dim AppN As String
   
   If IsMissing(appname) Then
      AppN = App.ProductName
   Else
      AppN = CStr(appname)
   End If

   v = GetAllSettings(AppN, "LastFiles")
   
   If Not IsEmpty(v) Then
      i = UBound(v, 1)
      Me.Clear
      LastFiles.Add v(i, 1), UCase(v(i, 1))
      
      For j = i - 1 To LBound(v, 1) Step -1
         LastFiles.Add v(j, 1), UCase(v(j, 1)), 1
      Next j
   End If
End Sub
'
'  Gets the maximum size of the list.
'
Public Property Get Number() As Integer
   Number = Num
End Property

'
'  Sets the maximum size of the list.
'
Public Property Let Number(i As Integer)
   Num = i
End Property


Public Sub Remove(file As String)
   On Error Resume Next
   LastFiles.Remove UCase(file)
End Sub

Public Sub Save(Optional appname As Variant)
   Dim i As Integer
   Dim AppN As String
   
   On Error Resume Next
   
   If IsMissing(appname) Then
      AppN = App.ProductName
   Else
      AppN = CStr(appname)
   End If
      
   DeleteSetting AppN, "LastFiles"
   
   For i = 1 To LastFiles.Count
      SaveSetting AppN, "LastFiles", i, LastFiles(i)
   Next i
End Sub
'
'  Note:  The form must contain a menu control array
'         named mnuLastFiles that is at least as big
'         as Number.
'
'
Public Sub Update(f As Form)
   Dim i As Long
   
  
On Error GoTo NextStep
   For i = 1 To Num
      f.mnuLastFiles(i).Visible = False
   Next i

NextStep:

On Error GoTo MenuEnd
   If LastFiles.Count > 0 Then
      f.mnuLastFiles(0).Visible = True
      
      For i = 1 To LastFiles.Count
         f.mnuLastFiles(i).Caption = LastFiles(i)
         f.mnuLastFiles(i).Visible = True
      Next i
      
      Do
        f.mnuLastFiles(i).Visible = False
        i = i + 1
      Loop
   Else
      i = 0
      Do
        f.mnuLastFiles(i).Visible = False
        i = i + 1
      Loop
   End If
   
MenuEnd:
End Sub


Private Sub Class_Initialize()
   Num = 5
End Sub


