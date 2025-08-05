// Guide-post comment for Program.fs
// This F# file is part of the Visio Diagram Generator.
// TODO: Replace this guide-post with your implementation using functional paradigms.
// Embrace immutability, pure functions, and pattern matching.
// Use clear, meaningful names for classes, methods, and variables.
// Encapsulate functionality and follow SOLID principles.
// Add XML documentation comments and unit tests to improve maintainability.
// Handle exceptions gracefully and log meaningful messages.
// Keep methods short and focused; avoid deep nesting and duplicated code.
// Use asynchronous programming for I/O-bound tasks when appropriate.
// See the provided standard and class modules for inspiration and reuse common patterns.
' Module: Program
' Purpose: application entry point for CLI; dispatch commands.
Module Program
    Sub Main()
        Dim args As String() = Environment.GetCommandLineArgs()
        Dim cmd As Command = ArgParser.Parse(args)
        Select Case cmd.Name
            Case "generate"
                GenerateCommand.Execute(cmd.Options)
            Case "export"
                ExportCommand.Execute(cmd.Options)
            Case Else
                ' Print help text and available commands
                LoggingExtensions.LogInfo("Usage: vdg generate <source> | vdg export <diagram> <format> <output>")
        End Select
    End Sub
End Module