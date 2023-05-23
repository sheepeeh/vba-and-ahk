' Replace super and subscript formatted text with unicode equivalents

Sub ReplaceSubscript()
    Dim vFindText As Variant
    Dim vReplText As Variant
    Dim sFormat As Boolean
    Dim sQuotes As String
    Dim oRng As Range
    Dim i As Long

    sFormat = Options.AutoFormatAsYouTypeReplaceQuotes

    vFindText = Array(ChrW(48), ChrW(49), ChrW(50), ChrW(51), ChrW(52), ChrW(53), ChrW(54), ChrW(55), ChrW(56), ChrW(57), ChrW(43), ChrW(45), ChrW(8722), ChrW(61), ChrW(40), ChrW(41), ChrW(97), ChrW(101), ChrW(104), ChrW(105), ChrW(106), ChrW(107), ChrW(108), ChrW(109), ChrW(110), ChrW(111), ChrW(112), ChrW(114), ChrW(115), ChrW(116), ChrW(117), ChrW(118), ChrW(120), ChrW(601), ChrW(946), ChrW(947), ChrW(961), ChrW(966), ChrW(967))
    vReplText = Array(ChrW(8320), ChrW(8321), ChrW(8322), ChrW(8323), ChrW(8324), ChrW(8325), ChrW(8326), ChrW(8327), ChrW(8328), ChrW(8329), ChrW(8330), ChrW(8331), ChrW(8331), ChrW(8332), ChrW(8333), ChrW(8334), ChrW(8336), ChrW(8337), ChrW(8341), ChrW(7522), ChrW(11388), ChrW(8342), ChrW(8343), ChrW(8344), ChrW(8345), ChrW(8338), ChrW(8346), ChrW(7523), ChrW(8347), ChrW(8348), ChrW(7524), ChrW(7525), ChrW(8339), ChrW(8340), ChrW(7526), ChrW(7527), ChrW(7528), ChrW(7529), ChrW(7530))
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    For i = LBound(vFindText) To UBound(vFindText)
        Set oRng = ActiveDocument.Range
        With oRng.Find
        .Font.Subscript = True
        Do While .Execute(vFindText(i))

        oRng.Text = vReplText(i)
        oRng.Font.Subscript = False
        oRng.Collapse wdCollapseEnd

        Loop
        End With
        Next i

        ActiveDocument.Range.AutoFormat
    Options.AutoFormatAsYouTypeReplaceQuotes = sFormat
lbl_Exit:
    Exit Sub
End Sub



Sub ReplaceSuperscript()
    Dim vFindText As Variant
    Dim vReplText As Variant
    Dim sFormat As Boolean
    Dim sQuotes As String
    Dim oRng As Range
    Dim i As Long

    sFormat = Options.AutoFormatAsYouTypeReplaceQuotes

    vFindText = Array(ChrW(48), ChrW(49), ChrW(50), ChrW(51), ChrW(52), ChrW(53), ChrW(54), ChrW(55), ChrW(56), ChrW(57), ChrW(43), ChrW(45), ChrW(8722), ChrW(61), ChrW(40), ChrW(41), ChrW(97), ChrW(98), ChrW(99), ChrW(100), ChrW(101), ChrW(102), ChrW(103), ChrW(104), ChrW(105), ChrW(106), ChrW(107), ChrW(108), ChrW(109), ChrW(110), ChrW(111), ChrW(112), ChrW(114), ChrW(115), ChrW(116), ChrW(117), ChrW(118), ChrW(119), ChrW(120), ChrW(121), ChrW(122), ChrW(65), ChrW(66), ChrW(68), ChrW(69), ChrW(71), ChrW(72), ChrW(73), ChrW(74), ChrW(75), ChrW(76), ChrW(77), ChrW(78), ChrW(79), ChrW(80), ChrW(82), ChrW(84), ChrW(85), ChrW(86), ChrW(87), ChrW(945), ChrW(946), ChrW(947), ChrW(948), ChrW(949), ChrW(952), ChrW(617), ChrW(966), ChrW(967))
    vReplText = Array(ChrW(8304), ChrW(185), ChrW(178), ChrW(179), ChrW(8308), ChrW(8309), ChrW(8310), ChrW(8311), ChrW(8312), ChrW(8313), ChrW(8314), ChrW(8315), ChrW(8315), ChrW(8316), ChrW(8317), ChrW(8318), ChrW(7491), ChrW(7495), ChrW(7580), ChrW(7496), ChrW(7497), ChrW(7584), ChrW(7501), ChrW(688), ChrW(8305), ChrW(690), ChrW(7503), ChrW(737), ChrW(7504), ChrW(8319), ChrW(7506), ChrW(7510), ChrW(691), ChrW(738), ChrW(7511), ChrW(7512), ChrW(7515), ChrW(695), ChrW(739), ChrW(696), ChrW(7611), ChrW(7468), ChrW(7470), ChrW(7472), ChrW(7473), ChrW(7475), ChrW(7476), ChrW(7477), ChrW(7478), ChrW(7479), ChrW(7480), ChrW(7481), ChrW(7482), ChrW(7484), ChrW(7486), ChrW(7487), ChrW(7488), ChrW(7489), ChrW(11389), ChrW(7490), ChrW(7493), ChrW(7517), ChrW(7518), ChrW(7519), ChrW(7499), ChrW(7615), ChrW(7589), ChrW(7520), ChrW(7521))
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    For i = LBound(vFindText) To UBound(vFindText)
        Set oRng = ActiveDocument.Range
        With oRng.Find
        .Font.Superscript = True
        Do While .Execute(vFindText(i))

        oRng.Text = vReplText(i)
        oRng.Font.Superscript = False
        oRng.Collapse wdCollapseEnd

        Loop
        End With
        Next i

    Options.AutoFormatAsYouTypeReplaceQuotes = sFormat
lbl_Exit:
    Exit Sub
End Sub
