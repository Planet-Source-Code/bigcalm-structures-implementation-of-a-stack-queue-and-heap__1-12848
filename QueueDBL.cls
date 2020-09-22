VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueueDBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' QueueDBL = Queue Double Linked List
Private HeadPtr As QueueDBLItem
Private TailPtr As QueueDBLItem
Public Count As Long

Public Property Get FirstElement() As Object
    If Count > 0 Then
        Set FirstElement = HeadPtr.Data
    Else
        Set FirstElement = Nothing
    End If
End Property
Public Property Get LastElement() As Object
    If Count > 0 Then
        Set LastElement = TailPtr.Data
    Else
        Set LastElement = Nothing
    End If
End Property
Private Sub Class_Initialize()
    Set HeadPtr = Nothing
    Set TailPtr = Nothing
    Count = 0
End Sub


Private Sub Class_Terminate()
Dim CurrentPtr As QueueDBLItem
Dim NextPtr As QueueDBLItem
    Set CurrentPtr = HeadPtr
    Do While Not (CurrentPtr Is Nothing)
        Set NextPtr = CurrentPtr.NextElement
        Set CurrentPtr.NextElement = Nothing
        Set CurrentPtr = Nothing
        Set CurrentPtr = NextPtr
        If Not (CurrentPtr Is Nothing) Then
            Set CurrentPtr.PrevElement = Nothing
        End If
    Loop
End Sub

Public Sub Add(Item As Object, Optional pStartOrEnd As Long = 1)
Dim NewElement As QueueDBLItem
    Set NewElement = New QueueDBLItem
    With NewElement
        Set .Data = Item
    
        If Count = 0 Then
            Set .NextElement = Nothing
            Set .PrevElement = Nothing
            Set HeadPtr = NewElement
            Set TailPtr = NewElement
        Else
            If pStartOrEnd And 1 Then
                ' Insert at head
                Set .PrevElement = Nothing
                Set .NextElement = HeadPtr
                Set HeadPtr.PrevElement = NewElement
                Set HeadPtr = NewElement
            Else
                ' Insert at tail
                Set .PrevElement = TailPtr
                Set .NextElement = Nothing
                Set TailPtr.NextElement = NewElement
                Set TailPtr = NewElement
            End If
        End If
    End With
    Count = Count + 1
End Sub

Public Sub Remove(Optional pHeadOrTail As Long = 1)
Dim OldPtr As QueueDBLItem
    If Count = 0 Then
        Exit Sub
    End If
    If pHeadOrTail And 1 Then
        ' Remove head
        Set OldPtr = HeadPtr
        Set HeadPtr = HeadPtr.NextElement
        If Not (HeadPtr Is Nothing) Then
            Set HeadPtr.PrevElement = Nothing
        End If
        Set OldPtr = Nothing
    Else
        ' Remove tail
        Set OldPtr = TailPtr
        Set TailPtr = TailPtr.PrevElement
        If Not (TailPtr Is Nothing) Then
            Set TailPtr.NextElement = Nothing
        End If
        Set OldPtr = Nothing
    End If
    Count = Count - 1
End Sub
