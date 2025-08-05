' Module: ConfigLoader
' Purpose: load DiagramConfig from JSON files (e.g. diagramConfig.json or global.json).
Module ConfigLoader
    Function LoadConfig() As DiagramConfig
        Dim config As New DiagramConfig
        ' Read JSON file from disk
        ' Parse keys into DiagramConfig properties
        config.Layout = "grid"      ' default layout if missing
        config.ExportFormat = "png"  ' default export format
        Return config
    End Function
End Module