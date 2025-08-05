// Guide-post comment for RibbonCommands.cs
// This file is part of the Visio Diagram Generator.
// TODO: Replace this guide-post with your implementation.
// Use clear, meaningful names for classes, methods, and variables.
// Encapsulate functionality and follow SOLID principles.
// Add XML documentation comments and unit tests to improve maintainability.
// Handle exceptions gracefully and log meaningful messages.
// Keep methods short and focused; avoid deep nesting and duplicated code.
// Use asynchronous programming for I/O-bound tasks when appropriate.
// See the provided standard and class modules for inspiration and reuse common patterns.
' Module: RibbonCommands
' Purpose: centralize command handlers invoked by the ribbon.
Module RibbonCommands
    Sub Generate()
        Dim cfg As DiagramConfig = ConfigLoader.LoadConfig()
        Dim pipeline As New Pipeline(cfg)
        pipeline.Run("defaultInput.fs")
    End Sub
    Sub Export()
        Dim diagram As Object = Nothing
        Dim exporter As New DiagramExporter
        exporter.Export(diagram, ExportFormat.Png, "diagram.png")
    End Sub
End Module