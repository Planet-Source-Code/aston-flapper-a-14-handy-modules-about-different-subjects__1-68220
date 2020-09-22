Attribute VB_Name = "modTreeView"
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long 'window api


Public Function tvnPut(Tree As TreeView, NodesText As String, PlusString As String, Optional Speedup As Boolean = False) As Long
'
'This function put nodes from text in an treeview
'Return = time it took in milliseconds
'Tree = Treeview to place the nodes in
'
'NodesText = the nodes in text form in an format like this:
'
'                ----------------------
'                |  Node1             |
'                |  >Subnode1         |
'                |  >Subnode2         |
'                |  >>SubSubnode1     |
'                |  >>SubSubnode2     |
'                |  >Subnode3         |
'                ----------------------
'
'PlusString is the seperator string, if the PlusString is "#" then
'The above example must be changed to:
'
'                ----------------------
'                |  Node1             |
'                |  #Subnode1         |
'                |  #Subnode2         |
'                |  ##SubSubnode1     |
'                |  ##SubSubnode2     |
'                |  #Subnode3         |
'                ----------------------
'
'Multicharacter PlusStrings are allowed! so "->" can be used like this:
'
'                ----------------------
'                |  Node1             |
'                |  ->Subnode1        |
'                |  ->Subnode2        |
'                |  ->->SubSubnode1   |
'                |  ->->SubSubnode2   |
'                |  ->Subnode3        |
'                ----------------------
'
'------------------------------
'Dim some variables
Dim PlusUses(999) As String      'Used for wich key the sub, subsub, subsubsubnodes etc uses
Dim Plussus As Long              'Used for storing wich node you are now in, sub, subsub, subsubsub, subsubsubsub
Dim A As Long                    'Used in the first loop
Dim B As Long                    'Used in the second loop
Dim G() As String                'Used to store the lines in an array
Dim NowNode As String            'The now where you are now in
'------------------------------
tmr = Timer                      'To calculate the time it costed
If Speedup Then StartEditTree Tree              'Puts tree in an fast state for editing
Tree.Nodes.Clear                 'Clear nodes
G = Split(NodesText, vbCrLf)     'Put the text lines in array G
NowNode = ""                     'Not neccacary
'------------------------------
For A = 0 To UBound(G)           'Begin loop from first line to last line
    DoEvents                     'To let the program not freeze with big files
    For B = 1 To Len(G(A)) Step Len(PlusString) 'Start loop to search for PlusString
        If Mid$(G(A), B, Len(PlusString)) <> PlusString Then Exit For
    Next
    '------------------------------
    Plussus = (B - 1) / Len(PlusString)         'Plussus = node depth
    '------------------------------
    If Plussus = 0 Then
        NowNode = ""                            'Set it is not an subnode
    Else
        NowNode = PlusUses(Plussus - 1)         'It is an subnode
    End If
    '------------------------------
    If NowNode <> "" Then
        'Add the subnode
        Tree.Nodes.Add NowNode, tvwChild, "x" & A, Replace$(Mid$(G(A), (Plussus * Len(PlusString)) + 1), "%A", A)
    Else
        'Add the node
        Tree.Nodes.Add , tvwChild, "x" & A, Replace$(Mid$(G(A), Plussus + 1), "%A", A)
    End If
    '------------------------------
    PlusUses(Plussus) = "x" & A 'Sets the key of this depth
Next
'------------------------------
If Speedup Then StopEditTree Tree                          'Removes the state of fast editing
Tree.Refresh                                'Refreshes it
tvnPut = (Timer * 1000) - (tmr * 1000)    'Return the time it costed in milliseconds
End Function

Public Sub StartEditTree(Tree As TreeView)
Tree.visible = False             'Speeds up: tree isn't visible
Tree.Enabled = False             'Speeds up: tree isn't enabled
LockWindowUpdate Tree.hwnd       'Speeds up: tree can't update
End Sub


Public Sub StopEditTree(Tree As TreeView)
Tree.visible = True              'makes it visible again
Tree.Enabled = True              'Makes it enable again
LockWindowUpdate 0               'Unlock it, only 1 window at a time locked
End Sub








