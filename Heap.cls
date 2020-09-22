VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Heap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Heap data structure.
' Sorted Binary Tree

' Testing:
' Time taken to insert 20,000 command buttons into heap: 991ms
' Time taken to destroy 20,000 command buttons from heap: 8201ms

'   Elements:           10      100     1000       10000       100000
'Time (Add/Delete):     1/0     4/2     38/15     469/1262     8211/???A VERY long time

Public TopNode As HeapNode
Public Count As Long
Private LeftOrRightDeletion As Long ' Variable is used to vary whether we pick the left or right item to "bring up" the tree during a deletion.
                                                            ' If we were to constantly pick "left" we might end up with a skewed heap.
' Class functions
Private Sub Class_Initialize()
    Set TopNode = Nothing
    Count = 0
    LeftOrRightDeletion = 0
End Sub
Private Sub Class_Terminate()
    DestroyHeap
End Sub

Public Function Delete(ByVal ItemKey As String, ByVal Value As Single) As Boolean
Dim NodeToDelete As HeapNode
Dim ParentNode As HeapNode
Dim FoundAtDepth As Long
Dim ChildNode As HeapNode
Dim SeperateTree As HeapNode
Dim ReinsertNode As HeapNode
Dim LeftOrRightChild As Long

    Set NodeToDelete = FindInHeap(Value, ItemKey)
    If NodeToDelete Is Nothing Then
        Delete = False
        Exit Function
    End If
    Delete = True
    If NodeToDelete.LeftChild Is Nothing Then
        If NodeToDelete.RightChild Is Nothing Then
            If NodeToDelete.ParentNode Is Nothing Then
                ' Deleting top level element
                Set NodeToDelete = Nothing
                Set TopNode = Nothing
            Else
                ' Easy peasy, we're deleting a bottom level element.
                Set ParentNode = NodeToDelete.ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then ' Left child
                    Set ParentNode.LeftChild = Nothing
                Else
                    Set ParentNode.RightChild = Nothing
                End If
                Set NodeToDelete = Nothing
            End If
        Else
            ' We're deleting a node with no left children, but with right children.
            '   5                                                   8
            '  /  \                                                /   \
            ' 4    6                   OR                    6     10
            '         \                                            \
            '           7                                           7
            ' Deleting Node 6.  So, we need to attach 5 to 7, or deleting node 6, so we need to attach 7 to 8
            If NodeToDelete.ParentNode Is Nothing Then
                ' I've got to rejig the top node.
                Set TopNode = NodeToDelete.RightChild
                Set NodeToDelete.RightChild.ParentNode = Nothing
                Set NodeToDelete = Nothing
            Else
                Set ParentNode = NodeToDelete.ParentNode
                Set NodeToDelete.RightChild.ParentNode = ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then ' Left child of parent node
                    Set ParentNode.LeftChild = NodeToDelete.RightChild
                Else    ' Right child of parent node
                    Set ParentNode.RightChild = NodeToDelete.RightChild
                End If
                Set NodeToDelete = Nothing ' Destroys node.
                ' We now have: from the previous examples:
                '    5                                              8
                '  /   \                  OR                    /    \
                ' 4    7                                       7      10
            End If
        End If
    Else
        If NodeToDelete.RightChild Is Nothing Then
                ' We're deleting a node with no right children, but with left children.
            If NodeToDelete.ParentNode Is Nothing Then
                ' I've got to rejig the top node.
                Set TopNode = NodeToDelete.LeftChild
                Set NodeToDelete.LeftChild.ParentNode = Nothing
                Set NodeToDelete = Nothing
            Else
                Set ParentNode = NodeToDelete.ParentNode
                Set NodeToDelete.LeftChild.ParentNode = ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then ' Left child of parent node
                    Set ParentNode.LeftChild = NodeToDelete.LeftChild
                Else
                    Set ParentNode.RightChild = NodeToDelete.LeftChild
                End If
                Set NodeToDelete.LeftChild.ParentNode = ParentNode
                Set NodeToDelete = Nothing ' Destroys node
            End If
        Else
            ' We are deleting an element with both Left and Right Children nodes.
            ' If we delete the Left hand node:
            ' Step 1: Store all right hand children off the left child in a seperate tree
            ' Step 2: Set the Left hand children's Right Child node to be the deleting item's
            '             right hand node.
            ' Step 3: Delete the now unattached top node.
            ' Step 4: Find place for the seperate tree
            ' Step 5: Re-insert
            ' e.g For deleting item 5:
            '       Initial                  Step 1                                   Steps 2&3                       Steps 4&5
            '          5                           5                                         2                                      2
            '        /     \                      /    \                                     /   \                                   /   \
            '      2        8                 2       8       Tree 2                 1     8   Tree 2                   1     8
            '    /   \     /   \              /  \     /  \         \                            /  \       \                            /  \
            '  1     4  7     9          1    4  7    9        4                         7    9      4                         7   9
            '                                                                                                                          /
            '                                                                                                                        4
            ' The reverse is used for right-hand child.
            If LeftOrRightDeletion = 0 Then ' Left node becomes this node
                ' Step 1 - Store
                Set ChildNode = NodeToDelete.LeftChild
                Set SeperateTree = ChildNode.RightChild
                If Not (SeperateTree Is Nothing) Then
                    Set SeperateTree.ParentNode = Nothing
                End If
                ' Step 2 - Move
                Set ChildNode.RightChild = NodeToDelete.RightChild
                Set ChildNode.RightChild.ParentNode = ChildNode
                If NodeToDelete.ParentNode Is Nothing Then
                    Set ChildNode.ParentNode = Nothing
                    Set TopNode = ChildNode
                Else
                    Set ChildNode.ParentNode = NodeToDelete.ParentNode
                    If NodeToDelete.Value <= NodeToDelete.ParentNode.Value Then
                        Set NodeToDelete.ParentNode.LeftChild = ChildNode
                    Else
                        Set NodeToDelete.ParentNode.RightChild = ChildNode
                    End If
                End If
                ' Step 3 - Delete
                Set NodeToDelete = Nothing
                ' Step 4 - Reinsert
                If Not (SeperateTree Is Nothing) Then
                    Set ReinsertNode = FindFreeSpace(SeperateTree.Value, LeftOrRightChild)
                    Set SeperateTree.ParentNode = ReinsertNode
                    If LeftOrRightChild = 0 Then
                        Set ReinsertNode.LeftChild = SeperateTree
                    Else
                        Set ReinsertNode.RightChild = SeperateTree
                    End If
                End If
            Else ' Right node becomes this node
                ' Step 1 - Store
                Set ChildNode = NodeToDelete.RightChild
                Set SeperateTree = ChildNode.LeftChild
                If Not (SeperateTree Is Nothing) Then
                    Set SeperateTree.ParentNode = Nothing
                End If
                ' Step 2 - Move
                Set ChildNode.LeftChild = NodeToDelete.LeftChild
                Set ChildNode.LeftChild.ParentNode = ChildNode
                If NodeToDelete.ParentNode Is Nothing Then
                    Set ChildNode.ParentNode = Nothing
                    Set TopNode = ChildNode
                Else
                    Set ChildNode.ParentNode = NodeToDelete.ParentNode
                    If NodeToDelete.Value <= NodeToDelete.ParentNode.Value Then
                        Set NodeToDelete.ParentNode.LeftChild = ChildNode
                    Else
                        Set NodeToDelete.ParentNode.RightChild = ChildNode
                    End If
                End If
                ' Step 3 - Delete
                Set NodeToDelete = Nothing
                ' Step 4 - Reinsert
                If Not (SeperateTree Is Nothing) Then
                    Set ReinsertNode = FindFreeSpace(SeperateTree.Value, LeftOrRightChild)
                    Set SeperateTree.ParentNode = ReinsertNode
                    If LeftOrRightChild = 0 Then
                        Set ReinsertNode.LeftChild = SeperateTree
                    Else
                        Set ReinsertNode.RightChild = SeperateTree
                    End If
                End If
            End If ' Left or Right move
            LeftOrRightDeletion = LeftOrRightDeletion Xor 1
        End If ' Both left and right nodes filled
    End If ' Left child is nothing
    Count = Count - 1
End Function

Public Function Add(Obj As Object, Value As Single, ItemKey As String) As HeapNode
Dim objHeapNode As New HeapNode
Dim objParentNode As HeapNode
Dim LeftOrRight As Long
    
    ' Initialize the heap node that we're going to insert
    With objHeapNode
        .ItemKey = ItemKey
        Set .Item = Obj
        .Value = Value
        Set .LeftChild = Nothing
        Set .RightChild = Nothing
        Set .ParentNode = Nothing
    End With

    ' If this is the first element being added....
    If TopNode Is Nothing Then
        Set TopNode = objHeapNode
    Else
        Set objParentNode = FindFreeSpace(Value, LeftOrRight)
        With objHeapNode
            Set .ParentNode = objParentNode
        End With
        With objParentNode
            If LeftOrRight = 0 Then
                Set objParentNode.LeftChild = objHeapNode
            Else
                Set objParentNode.RightChild = objHeapNode
            End If
        End With
    End If
    Count = Count + 1
End Function

Public Function FindInHeap(Value As Single, ItemKey As String) As HeapNode
Dim CurrentNode As HeapNode
Dim FoundIt As Boolean

    Set CurrentNode = TopNode
    FoundIt = False
    Do While FoundIt = False
        If Value <= CurrentNode.Value Then
            If Value = CurrentNode.Value And ItemKey = CurrentNode.ItemKey Then
                ' gotcha
                Set FindInHeap = CurrentNode
                FoundIt = True
            Else
                If CurrentNode.LeftChild Is Nothing Then ' Cant find it.
                    Set FindInHeap = Nothing
                    FoundIt = True
                Else
                    Set CurrentNode = CurrentNode.LeftChild
                End If
            End If
        Else
            If CurrentNode.RightChild Is Nothing Then ' cant find it
                Set FindInHeap = Nothing
                FoundIt = True
            Else
                Set CurrentNode = CurrentNode.RightChild
            End If
        End If
    Loop
End Function
Public Function GetLeftMostElement() As HeapNode
Dim CurrentNode As HeapNode
    Set CurrentNode = TopNode
    If TopNode Is Nothing Then
        Set GetLeftMostElement = Nothing
        Exit Function
    End If
    Do While Not (CurrentNode.LeftChild Is Nothing)
        Set CurrentNode = CurrentNode.LeftChild
    Loop
    Set GetLeftMostElement = CurrentNode
End Function

Public Function FindFreeSpace(ByRef Value As Single, ByRef LeftOrRight As Long) As HeapNode
Dim CurrentNode As HeapNode
Dim FoundFree As Boolean
    ' Note that Identical values will always go down the left hand side of the tree
    
    Set CurrentNode = TopNode
    FoundFree = False
    Do While FoundFree = False
        If Value <= CurrentNode.Value Then
            If CurrentNode.LeftChild Is Nothing Then
                FoundFree = True
                LeftOrRight = 0
            Else
                Set CurrentNode = CurrentNode.LeftChild
            End If
        Else
            If CurrentNode.RightChild Is Nothing Then
                FoundFree = True
                LeftOrRight = 1
            Else
                Set CurrentNode = CurrentNode.RightChild
            End If
        End If
    Loop
    Set FindFreeSpace = CurrentNode
End Function

Public Function DestroyHeap() As Boolean
Dim CurrentNode As HeapNode
Dim ParentNode As HeapNode
    Set CurrentNode = TopNode
    Do While Count > 0
        If CurrentNode.LeftChild Is Nothing Then
            If CurrentNode.RightChild Is Nothing Then
                If CurrentNode.ParentNode Is Nothing Then ' Top level element
                    Set CurrentNode = Nothing
                    Set TopNode = Nothing
                Else
                    Set ParentNode = CurrentNode.ParentNode
                    If CurrentNode.Value <= ParentNode.Value Then ' Left hand child
                        Set ParentNode.LeftChild = Nothing
                    Else ' Right hand child
                        Set ParentNode.RightChild = Nothing
                    End If
                    Set CurrentNode = Nothing ' just make sure it kills the memory associated with the object
                    Set CurrentNode = ParentNode
                End If
                Count = Count - 1
            Else
                ' Right child is not nothing, therefore progress to that element and check again
                Set CurrentNode = CurrentNode.RightChild
            End If
        Else
            ' Left child is not nothing, so progress to that node and check again
            Set CurrentNode = CurrentNode.LeftChild
        End If
    Loop
End Function

