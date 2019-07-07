Attribute VB_Name = "QuickSort_GeoSets"
'Target:        CATIA V6 - 3D Experience Platform
'Title:         QuickSort_GeoSets Macro
'Description:   This Macro will reorder GeoSets alphabetically
'Author:        John Boyle
'Version:       1.0
'Modified Date: 3/November/2017 (1.0)

Sub CATMain()

    'Identify Selection
    Dim mySel As Selection
    Set mySel = CATIA.ActiveEditor.Selection
    'MsgBox "Selection = " + mySelect.Item(1).Name

    'Identify Part
    Dim oPart As Part
    Set oPart = CATIA.ActiveEditor.ActiveObject
    Dim MyPart As Variant
    Set MyPart = oPart

    'Identify Geosets
    Dim oHB As HybridBody
    Dim oHBs As HybridBodies
    Set oHBs = oPart.HybridBodies

    'Initialize Sorting Array
    Dim MyList() As Variant
    Dim iGSCount As Integer
    'Find number of GSs
    iGSCount = oHBs.Count
    ReDim MyList(1 To iGSCount)

    'Index the array
    For i = 1 To iGSCount
        MyList(i) = i
    Next i

    'Show unsorted list of GSs
    Dim str As String
    str = ""
    For i = 1 To iGSCount
        str = str + " " + oHBs.Item(MyList(i)).Name
    Next
    MsgBox str

    'Create sorted list
    sQuickSort MyList, oHBs, 1, iGSCount

    'Show sorted list of GSs
    str = ""
    For i = 1 To iGSCount
        str = str + " " + oHBs.Item(MyList(i)).Name
    Next
    MsgBox str

    'Move the Geosets to match the sorted list
    'Copy GeoSet and Paste at end, then cut original
    Dim oMyGS As Variant
    Dim Target As Integer

    For i = 1 To iGSCount
        Target = MyList(i)
        Set oMyGS = oHBs.Item(MyList(i))
        'Move the GS
        mySel.Clear
        mySel.Add (oMyGS)
        mySel.Copy
        mySel.Clear
        mySel.Add (MyPart)
        mySel.Paste
        mySel.Clear
        mySel.Add (oMyGS)
        mySel.Cut

        'Bump all MyList pointers > than Target down by 1
        sShift MyList, MyList(i), iGSCount

    Next

End Sub

Private Sub sQuickSort(ByRef A() As Variant, ByRef B As HybridBodies, ByVal lo As Long, ByVal hi As Long)
    If lo < hi Then
        p = sPartition(A, B, lo, hi)
        sQuickSort A, B, lo, p - 1
        sQuickSort A, B, p + 1, hi
    End If

End Sub

Private Function sPartition(C() As Variant, D As HybridBodies, ByVal lo As Long, ByVal hi As Long)
    Dim iPivot As String
    Dim i, j As Integer

    iPivot = ""
    iPivot = D.Item(C(hi)).Name
    i = lo - 1

    For j = lo To hi - 1
        If D.Item(C(j)).Name < iPivot Then
            i = i + 1
            sSwap C, i, j
        End If
    Next
    If D.Item(C(hi)).Name < D.Item(C(i + 1)).Name Then
        sSwap C, i + 1, hi
    End If
    sPartition = i + 1

End Function

Private Sub sSwap(ByRef E As Variant, ByVal loc As Long, ByVal tar As Long)
    Dim t As Integer
    t = E(loc)
    E(loc) = E(tar)
    E(tar) = t

End Sub

Private Sub sShift(ByRef F As Variant, ByVal piv As Long, ByVal cnt As Long)

    For i = 1 To cnt
        If F(i) > piv Then
            F(i) = F(i) - 1
        End If
    Next

End Sub
