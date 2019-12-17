Attribute VB_Name = "prepWordForHTML"
Function RemoveContentControls()
    Dim cc As ContentControl
    Dim docCCs As ContentControls
    Dim ccTag As String
    Dim ccCount As Integer
    ccCount = 0
    
    
    ' Iterate through all the content controls in the document
    ' and select those having the specified tag value.
    If ActiveDocument.ContentControls.Count <> 0 Then
        For Each cc In ActiveDocument.ContentControls
            cc.Delete
        Next
    End If
     
End Function

Function ConvertNumberedListToManual()
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
End Function

Function ConvertHeadings()
    ActiveDocument.Range.Select
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Box Heading")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Heading 3")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Heading 2")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Heading 3")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Heading 1")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Heading 2")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Caption")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "*"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("TOC Heading")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "*"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.WholeStory
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
End Function

Function RemoveHeaderFooter()
    Dim oHF As HeaderFooter
    Dim oSection As Section
    For Each oSection In ActiveDocument.Sections
        For Each oFF In oSection.Headers
            oFF.Range.Delete
        Next
        For Each oFF In oSection.Footers
            oFF.Range.Delete
        Next
    Next
End Function

Function RemoveBookmarks()
    Dim bkm As Bookmark
    ActiveDocument.Bookmarks.ShowHidden = True
    For Each bkm In ActiveDocument.Bookmarks
    bkm.Delete
    Next bkm
    ActiveDocument.TablesOfContents(1).Delete
End Function

Sub PrepWordForHtml()

    Dim Ret_type As Integer
    Dim strMsg As String
    Dim strTitle As String
    ' Dialog Message
    strMsg = "This macro removes and changes formatting and content. These actions cannot be undone. Click OK to continue or Cancel to stop."
    ' Dialog's Title
    strTitle = "WARNING: Action cannot be undone."
    'Display MessageBox
        Ret_type = MsgBox(strMsg, vbOKCancel + vbExclamation, strTitle)
    ' Check pressed button
    Select Case Ret_type
    ' If OK, execute Macros
    Case 1
        Call RemoveContentControls
        Call ConvertNumberedListToManual
        Call ConvertHeadings
        Call RemoveHeaderFooter
        Call RemoveBookmarks
        
    Case 2
        
    End Select
End Sub


