module LayoutTests

open Xunit
open DiagramCore   // namespace from your F# core project

type LayoutTests() =

    [<Fact>]
    member _.VerticalLayout_returns_expected_shape_count() =
        let items  = Fixtures.smallGraph()          // helper you’ll port
        let result = LayoutAlgorithms.vertical items
        Assert.Equal(items.Length, result.Nodes.Length)