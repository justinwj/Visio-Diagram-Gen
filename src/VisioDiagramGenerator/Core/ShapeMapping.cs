' Module: ShapeMapping
' Purpose: map domain types to Visio master names or stencil ids.
Module ShapeMapping
    Function GetShapeFor(typeName As String) As String
        ' Read mapping from shapeMapping.json or return default shape
        Return "Rectangle"
    End Function
End Module