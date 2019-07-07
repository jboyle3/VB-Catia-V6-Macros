Attribute VB_Name = "Close_All_Other_Open_parts"
'Target:        CATIA V6 - 3D Experience Platform
'Title:         Close_All_Other_Open_Parts Macro
'Description:   This Macro closes all open parts WITHOUT saving them except the target part
'Author:        John Boyle
'Version:       1.0
'Modified Date: 3/November/2017 (xxx)

Sub CATMain()

    '--Select and identify the part NOT to close
    Dim myPartSelect As AnyObject
    Set myPartSelect = CATIA.ActiveEditor.Selection

    'We ask the user to select the Part (3DShape)
    Dim Status, InputObjectType(0)
    InputObjectType(0) = "AnyObject"
    Status = myPartSelect.SelectElement(InputObjectType, "Select the target 3DShape", False)
        If (Status = "Cancel") Then Exit Sub

    'Retrieve the selected Part
    Dim MyPart As AnyObject
    Set MyPart = myPartSelect.Item(1).Value
    'Ensure that the 3DShape was selected, not the Root Occurance Part
    If TypeName(MyPart) <> "Part" Then
        MsgBox "Please make the 3DShape active, then select it"
        GoTo EndSub
    End If

    'Identifies the Part (3DShape) Name
    Dim oPart As Part
    Set oPart = CATIA.ActiveEditor.ActiveObject

    'Identifies the Prod Rep Ref Name & PLM_ExternalID
    Dim oVPMRepRef   As VPMRepReference
    Set oVPMRepRef = oPart.Parent

    'Display the Part name and type

    MsgBox "Part to leave open = " + MyPart.Parent.Father.Name

    '--Identify and retrieve all open parts
    'Retrieves session service related to Product data
    Dim oProductSessionService As ProductSessionService
    Set oProductSessionService = CATIA.GetSessionService("ProductSessionService")

    Dim oShape3Ds As Shape3Ds
    Set oShape3Ds = oProductSessionService.Shape3Ds

    'Disables file alert so user doesn't have to click on Save or Not - used to close donor parts w/o saving
    CATIA.DisplayFileAlerts = False

    'Set the loop counter
    Dim iShape3DsCount As Integer
    iShape3DsCount = 0
    iShape3DsCount = oShape3Ds.Count

    'Close donor parts
    Dim cWindows As Windows
    Dim oWindow As Window
    Set cWindows = CATIA.Windows

    For i = 0 To iShape3DsCount - 2
        Set oWindow = cWindows.Item(iShape3DsCount - i)
        oWindow.Close
    Next i

EndSub:
End Sub
