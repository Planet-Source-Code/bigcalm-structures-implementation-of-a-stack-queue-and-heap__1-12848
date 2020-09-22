VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Heap (v2)"
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Collection"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Text            =   "10"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Queue"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Heap"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stack"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Command1_Click()
Dim X As New Stack
Dim Y As CommandButton
Dim i As Long
Dim StartTime As Long
Dim EndTime As Long

'    X.Add Command1
'    X.Add Command2
'    X.Add Command3
'    For i = 1 To 3
'        Set Y = X.Retrieve
'        MsgBox Y.Caption
'    Next

    StartTime = timeGetTime
    For i = 1 To Text1.Text
        X.Push Command1
    Next
    EndTime = timeGetTime
    MsgBox "Time taken to put " & Text1.Text & " elemens in to stack = " & EndTime - StartTime & " ms"
    StartTime = timeGetTime
    For i = 1 To Text1.Text
        Set Y = X.Pop
    Next
    EndTime = timeGetTime
    MsgBox "Time taken to retrieve all " & Text1.Text & " elements out of stack = " & EndTime - StartTime & " ms"
End Sub

Private Sub Command2_Click()
Dim X As New Heap
Dim i As Long
Dim StartTime As Long
Dim EndTime As Long
    StartTime = timeGetTime
    For i = 1 To Text1.Text
        X.Add Command1, Rnd, "Yo"
    Next
    EndTime = timeGetTime
    MsgBox "Time taken to put " & Text1.Text & " elements in to heap = " & EndTime - StartTime & " ms"
    StartTime = timeGetTime
    Set X = Nothing
    EndTime = timeGetTime
    MsgBox "Time taken to remove " & Text1.Text & " elements in heap = " & EndTime - StartTime & " ms"
End Sub

Private Sub Command3_Click()
Dim X As New Queue
Dim Y As CommandButton
Dim i As Long
Dim StartTime As Long
Dim EndTime As Long

    StartTime = timeGetTime
    For i = 1 To Text1.Text
        X.Add Command1
    Next
    EndTime = timeGetTime
    MsgBox "Time taken to put " & Text1.Text & " elements in to queue = " & EndTime - StartTime & " ms"
    StartTime = timeGetTime
    For i = 1 To Text1.Text
        Set Y = X.Retrieve
    Next
    EndTime = timeGetTime
    MsgBox "Time taken to retrieve all " & Text1.Text & " elements out of queue = " & EndTime - StartTime & " ms"
End Sub

Private Sub Command4_Click()
Dim X As New Collection
Dim StartTime As Long
Dim EndTime As Long
Dim i As Long
    StartTime = timeGetTime
    For i = 1 To Text1.Text
        X.Add Command1
    Next
    EndTime = timeGetTime
    MsgBox "Time taken to put " & Text1.Text & " elements in to collection: " & EndTime - StartTime & " ms"
    StartTime = timeGetTime
    Set X = Nothing
    EndTime = timeGetTime
    MsgBox "Time taken to remove " & Text1.Text & " elements from collection: " & EndTime - StartTime & " ms"
End Sub

Private Sub Command5_Click()
Dim X As New Heap2
Dim i As Long
Dim StartTime As Long
Dim EndTime As Long
    StartTime = timeGetTime
    For i = 1 To Text1.Text
        X.Add Command1, Rnd, "Yo"
    Next
    EndTime = timeGetTime
    MsgBox "Time taken to put " & Text1.Text & " elements in to heap = " & EndTime - StartTime & " ms"
    StartTime = timeGetTime
    Set X = Nothing
    EndTime = timeGetTime
    MsgBox "Time taken to remove " & Text1.Text & " elements in heap = " & EndTime - StartTime & " ms"
End Sub
