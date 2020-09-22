VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Queue data structure
' You get back the objects you've added in the order you put them in.

' Testing using mCol as a VB collection
' On 466Mhz Celeron, added 20,000 objects (command buttons) to queue in 78.44 seconds in VB debug mode
' On 466Mhz Celeron, retrieved 10,000 objects (command buttons) from queue in 8.74 seconds in VB debug mode

' I assume from this that VB collections do not maintain a tail pointer (pointer to end of list).
' Given N elements, this means that Inserts will work very slowly as N grows large,
' because it must traverse the list each time (taking time O(N)).
' So I've changed the Q to use QueueDBL instead.  This should speed up a bit...

' On 466Mhz Celeron, added 20,000 (command buttons) to queue in 21.88 seconds in VB debug mode
' On 466Mhz Celeron, retrieved 20,000 (command buttons) to queue in 22.0 seconds in VB debug mode

' When not in debug mode:
' added 20,000 (command buttons) to Q in 483 ms
' retrieved 20,000 (command buttons) from Q in 1788ms


'   Elements:           10      100     1000       10000        100000
'Time (Add/Delete):    1/1      3/1     21/18     219/608      2311/43723

Private mCol As QueueDBL
'Public Count As Long

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Private Sub Class_Initialize()
    Set mCol = New QueueDBL
'    Count = 0
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub Add(pObject As Object)
    mCol.Add pObject, 0
End Sub

Public Function Retrieve() As Object
    If mCol.Count > 0 Then
        Set Retrieve = mCol.FirstElement
        mCol.Remove
    Else
        Set Retrieve = Nothing
    End If
End Function
