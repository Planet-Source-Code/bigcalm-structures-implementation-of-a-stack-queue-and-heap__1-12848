VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Stack data structure
' you get back the objects in the most recently added order
' Testing:
' On Celeron 466Mhz, pushed 20,000 objects (command buttons) to stack in 18.32 seconds whilst running in debug mode
' On Celeron 466Mhz, popped 10,000 objects (command buttons) off the stack in 8.22 seconds whilst running in debug mode

' Testing (non-debug mode):
' pushed 20,000 command buttons on to stack in 92ms
' popped 20,000 command buttons from stack in 85ms

'   Elements:           10      100     1000    10000   100000
'Time (Add/Delete):     0/0     1/1     4/4     40/44   403/447

' Testing for built in Collection class:
'   Elements:           10      100     1000    10000   100000
'Time (Add/Delete)      0/0     1/1     2/2     27/69   281/2099

Private mCol As Collection
Public Count As Long

Private Sub Class_Initialize()
    Set mCol = New Collection
    Count = 0
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub Push(pObject As Object)
    If Count = 0 Then
        mCol.Add pObject
    Else
        mCol.Add pObject, , 1
    End If
    Count = Count + 1
End Sub

Public Function Pop() As Object
    If Count > 0 Then
        Set Pop = mCol.Item(1)
        mCol.Remove (1)
    Else
        Set Pop = Nothing
    End If
    Count = Count - 1
End Function
