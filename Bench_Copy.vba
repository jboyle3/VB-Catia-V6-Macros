Attribute VB_Name = "Bench_Copy"
'Target:        CATIA V6 - 3D Experience Platform
'Title:         Bench_Copy Macro
'Description:   This Macro copies the 1st GeoSet in all open parts in the session into selected part
'Author:        John Boyle
'Version:       1.0
'Modified Date: 1/February/2018 (1.0)

Sub CATMain()
    '--Macro assumes target part is the active part & donor parts are open in session

    Dim cWindows As Windows
    Dim oWindow As Window
    Set cWindows = CATIA.Windows

    'Disables file alert so user doesn't have to click on Save or Not - used to close donor parts w/o saving
    CATIA.DisplayFileAlerts = False

    'CATIA.RefreshDisplay = True
    'CATIA.Interactive = True

    '--Select and identify the target part
    Dim myPartSel
    Set myPartSel = CATIA.ActiveEditor.Selection

    'We ask the user to select the Part (3DShape)
    Dim Status, InputObjectType(0)
    InputObjectType(0) = "AnyObject"
    Status = myPartSel.SelectElement(InputObjectType, "Select the target 3DShape", False)
        If (Status = "Cancel") Then Exit Sub

    'Retrieve the selected Part
    Dim MyPart As Variant
    Set MyPart = myPartSel.Item(1).Value
    'Ensure that the 3DShape was selected, not the Root Occurance Part
    If TypeName(MyPart) <> "Part" Then
        MsgBox "Please make the 3DShape active, then select it"
        GoTo EndSub
    End If

    'Display the Part name and type

    'MsgBox "Current selection = " + myPartSel.Name + vbCrLf _
    + "Selected feature name = " & myPart.Name & "; type = " & TypeName(myPart)

    Dim oHBs As HybridBodies
    Set oHBs = MyPart.HybridBodies

    '--Identify and retrieve all open parts

    'Retrieves session service related to Product data
    Dim oProdSessServ As ProductSessionService
    Set oProdSessServ = CATIA.GetSessionService("ProductSessionService")

    Dim oShape3Ds As Shape3Ds
    Set oShape3Ds = oProdSessServ.Shape3Ds

    'Set the loop counter
    Dim iShape3DsCount As Integer
    iShape3DsCount = 0
    iShape3DsCount = oShape3Ds.Count
    MsgBox iShape3DsCount

    'CATIA.RefreshDisplay = False
    'CATIA.Interactive = False

    '--Loop through open parts
    Dim iCount As Integer

    StatusBar = "Starting..."

    For i = 1 To iShape3DsCount
        Dim oShape3D As Shape3D
        Dim oShapePart

        Set oShape3D = oShape3Ds.Item(i)
        Set oShapePart = oShape3D.Part

        Set oWindow = cWindows.Item(i)
        oWindow.Activate

        'MsgBox oShapePart.Name + vbCrLf _
        + myPart.Name

        '--Check to see that next donor part is NOT the target part
        If oShapePart.Name <> MyPart.Name Then

            '--Change GeoSet name to the 3Dshape name
            Dim myFeat As Variant
            Dim myHBs, my3DShape
            Dim cHBCol As HybridBodies
            Set cHBCol = oShapePart.HybridBodies

            'Retrieves the Hybrid Body from the list
            Dim oHB As HybridBody
            Set oHB = cHBCol.Item(1)

            'Walks up the tree to the 3DShape
            Set myFeat = oHB
            Set myHBs = myFeat.Parent
            Set my3DShape = myHBs.Parent

            'Splits resulting string at the first two space delimeters and combines them
            Dim sGSName As String
            sGSName = Split(my3DShape.Name, " ")(0) & " " & Split(my3DShape.Name, " ")(1)

            'Rename the GeoSet
            myFeat.Name = sGSName
            iCount = iCount + 1

            '--Copy the GeoSet from donor part to the target part
            'Select the 1st GeoSet
            Dim myGSSel As Selection
            Dim myTar3DShSel As Selection
            Set myGSSel = CATIA.ActiveEditor.Selection

            myGSSel.Clear
            myGSSel.Add (myFeat)
            StatusBar = "Copying..."
            CATIA.Application.StartCommand ("Copy")

            'Switch to target part
            Set oWindow = cWindows.Item(1)
            oWindow.Activate

            'Select the 3DShape
            Set myTar3DShSel = CATIA.ActiveEditor.Selection
            myTar3DShSel.Clear
            myTar3DShSel.Add (MyPart)

            'Paste the GeoSet from the donor part
            StatusBar = "Pasting..."
            CATIA.Application.StartCommand ("Paste")

        End If

    '--Next part
    Next i
    MsgBox "Parts copied = " & iCount

    'Close donor parts
    StatusBar "Closing open parts"
    For i = 0 To iShape3DsCount - 2
        Set oWindow = cWindows.Item(iShape3DsCount - i)
        oWindow.Close
    Next i

    'CATIA.RefreshDisplay = True
    'CATIA.Interactive = True

EndSub:
End Sub
