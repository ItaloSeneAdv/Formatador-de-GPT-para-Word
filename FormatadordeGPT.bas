Attribute VB_Name = "Módulo1"
Sub FormatadordeGPT()

    Dim doc As Document
    Dim rng As Range
    Dim i As Integer
    Dim shp As InlineShape

    Set doc = ActiveDocument

    Set rng = doc.Content
    With rng.Find
        .Text = "\*\*(*)\*\*"            
        .Replacement.Text = "\1"          
        .Replacement.Font.Bold = True     
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True            
        .Execute Replace:=wdReplaceAll
    End With

    
    Set rng = doc.Content
    With rng.Find
        .Text = "---"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With


    Set rng = doc.Content
    With rng.Find
        .Text = "#"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With

  
    For i = doc.InlineShapes.Count To 1 Step -1
        Set shp = doc.InlineShapes(i)
        If shp.Type = wdInlineShapeHorizontalLine Then
            shp.Delete
        End If
    Next i

 
    Set shp = Nothing
    Set rng = Nothing
    Set doc = Nothing

End Sub

