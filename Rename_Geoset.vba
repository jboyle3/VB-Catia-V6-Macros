Attribute VB_Name = "Rename_GeoSet"
'Target:        CATIA V6 - 3D Experience Platform
'Title:         Rename_GeoSet Macro
'Description:   This Macro copies the 3D Shape Title to the selected GeoSet Feature name
'               attribute field so that the GeoSet Feature name is displayed in the tree
'Author:        John Boyle
'Version:       1.0
'Modified Date: 3/November/2017 (1.0 Working)

Sub CATMain()

    Dim mySelection
    Set mySelection = CATIA.ActiveEditor.Selection

    'We ask the user to select a feature
    Dim Status, InputObjectType(0)

    InputObjectType(0) = "AnyObject"
    Status = mySelection.SelectElement(InputObjectType, "Select a feature", False)
    If (Status = "Cancel") Then Exit Sub

    Dim myFeature, myHBodies, my3DShape
    Set myFeature = mySelection.Item(1).Value
    Set myHBodies = myFeature.Parent
    Set my3DShape = myHBodies.Parent

    'Debugging check
    'MsgBox "Selected feature name = " & myFeature.Name & "; type = " & TypeName(myFeature)

    Dim sGeoSetName As String
    sGeoSetName = Split(my3DShape.Name, " ")(0) & " " & Split(my3DShape.Name, " ")(1)

    'Debugging check
    'MsgBox "Result GeoSet name = " & sGeoSetName

    'Rename
    myFeature.Name = sGeoSetName

End Sub
