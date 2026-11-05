let
    Source = Web.BrowserContents("wikipedia.com"),
    #"Extracted Table From Html" = Html.Table(Source, {{"Column1", ".other-project-title"}, {"Column2", ".other-project-tagline"}}, [RowSelector=".other-project"]),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Table From Html",{{"Column1", type text}, {"Column2", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Column2", Order.Ascending}})
in
    #"Sorted Rows"