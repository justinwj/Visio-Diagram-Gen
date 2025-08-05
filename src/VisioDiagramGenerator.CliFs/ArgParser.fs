// Guide-post comment for ArgParser.fs
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
' Module: ArgParser
' Purpose: parse raw command line arguments into structured commands.
Module ArgParser
    Function Parse(args As String()) As Command
        Dim cmd As New Command
        ' Inspect args array to determine which command to run and set options
        If args.Length = 0 Then
            cmd.Name = "help"
        ElseIf args(0) = "generate" Then
            cmd.Name = "generate"
            If args.Length > 1 Then cmd.Options("InputPath") = args(1)
        ElseIf args(0) = "export" Then
            cmd.Name = "export"
            If args.Length > 1 Then cmd.Options("DiagramFile") = args(1)
            If args.Length > 2 Then cmd.Options("Format") = args(2)
            If args.Length > 3 Then cmd.Options("Output") = args(3)
        Else
            cmd.Name = "help"
        End If
        Return cmd
    End Function
End Module