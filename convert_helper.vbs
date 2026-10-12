Option Explicit
Dim oWord, oDoc, oSel, oShape
Dim sMode, sIn, sOut
Dim dPageW, dPageH, dRatioW, dRatioH, dRatio

If WScript.Arguments.Count < 3 Then WScript.Quit 1

sMode = LCase(WScript.Arguments(0))
sIn   = WScript.Arguments(1)
sOut  = WScript.Arguments(2)

Set oWord = CreateObject("Word.Application")
oWord.Visible       = False
oWord.DisplayAlerts = 0

If sMode = "docx" Then
    Set oDoc = oWord.Documents.Open(sIn, False, True)
    oDoc.SaveAs2 sOut, 17
    oDoc.Close False

ElseIf sMode = "image" Then
    Set oDoc = oWord.Documents.Add()
    oDoc.PageSetup.TopMargin    = 28.35
    oDoc.PageSetup.BottomMargin = 28.35
    oDoc.PageSetup.LeftMargin   = 28.35
    oDoc.PageSetup.RightMargin  = 28.35
    Set oSel   = oWord.Selection
    Set oShape = oSel.InlineShapes.AddPicture(sIn)
    dPageW = oDoc.PageSetup.PageWidth  - 56.7
    dPageH = oDoc.PageSetup.PageHeight - 56.7
    dRatioW = dPageW / oShape.Width
    dRatioH = dPageH / oShape.Height
    If dRatioW < dRatioH Then dRatio = dRatioW Else dRatio = dRatioH
    If dRatio < 1 Then
        oShape.Width  = oShape.Width  * dRatio
        oShape.Height = oShape.Height * dRatio
    End If
    oDoc.SaveAs2 sOut, 17
    oDoc.Close False
    Set oShape = Nothing
    Set oSel   = Nothing

ElseIf sMode = "merge" Then
    ' sIn = cesty k PDF oddelene znakem | (pipe)
    Dim sParts, i
    Dim oRng
    sParts = Split(sIn, "|")
    ' Novy prazdny dokument jako zakladna (vyhneme se PDF->Word konfliktum)
    Set oDoc = oWord.Documents.Add()
    For i = 0 To UBound(sParts)
        Set oRng = oDoc.Content
        oRng.Collapse 0      ' wdCollapseEnd
        If i > 0 Then
            oRng.InsertBreak 7   ' wdPageBreak
            oRng.Collapse 0
        End If
        oRng.InsertFile sParts(i)
    Next
    ' ExportAsFixedFormat je spolehlivejsi nez SaveAs2 pro PDF vystup
    oDoc.ExportAsFixedFormat sOut, 17, False, 0
    oDoc.Close False

ElseIf sMode = "extract" Then
    ' Arguments(3) = stranka od, Arguments(4) = stranka do
    If WScript.Arguments.Count < 5 Then WScript.Quit 1
    Dim nFrom, nTo
    nFrom = CInt(WScript.Arguments(3))
    nTo   = CInt(WScript.Arguments(4))
    Set oDoc = oWord.Documents.Open(sIn, False, False)
    ' ExportAsFixedFormat: OutputFileName, Format, OpenAfter, OptimizeFor, Range(3=FromTo), From, To
    oDoc.ExportAsFixedFormat sOut, 17, False, 0, 3, nFrom, nTo
    oDoc.Close False

Else
    oWord.Quit
    Set oWord = Nothing
    WScript.Quit 2
End If

oWord.Quit
Set oDoc  = Nothing
Set oWord = Nothing
WScript.Quit 0
