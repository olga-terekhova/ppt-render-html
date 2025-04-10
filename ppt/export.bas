Attribute VB_Name = "Module1"
Sub ExportSlideWithLinkData()

    Dim slide As slide
    Dim shape As shape
    Dim outputPath As String
    Dim slideIndex As Integer
    Dim slideWidth As Single, slideHeight As Single
    Dim imgPath As String
    Dim jsonJSPath As String
    Dim jsFile As Integer
    Dim linkData As String
    Dim linkArray As String
    Dim dpiFactor As Double

    ' Set up
    Set slide = ActivePresentation.Slides(ActiveWindow.View.slide.slideIndex)
    slideIndex = slide.slideIndex
    outputPath = ActivePresentation.Path & "\"
    
    ' Define desired export DPI multiplier (1.5x or 2x for HiDPI)
    dpiFactor = 2

    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight

    ' Export full slide as high-res PNG
    imgPath = outputPath & "slide_" & slideIndex & ".png"
    slide.Export imgPath, "PNG", slideWidth * dpiFactor, slideHeight * dpiFactor

    ' Start JS metadata structure
    linkData = "const slideMetadata = {" & vbCrLf
    linkData = linkData & "  width: " & slideWidth & "," & vbCrLf
    linkData = linkData & "  height: " & slideHeight & "," & vbCrLf
    linkData = linkData & "  links: [" & vbCrLf

    ' Collect link areas
    Dim hasLink As Boolean
    hasLink = False

    Dim i As Integer
    For i = 1 To slide.Shapes.Count
        Set shape = slide.Shapes(i)
        If shape.Type = msoPicture Then
            If Not shape.ActionSettings(ppMouseClick).Hyperlink.Address = "" Then
                hasLink = True
                linkData = linkData & "    { x: " & shape.Left & ", y: " & shape.Top & _
                    ", w: " & shape.Width & ", h: " & shape.Height & ", url: '" & _
                    shape.ActionSettings(ppMouseClick).Hyperlink.Address & "' }, " & vbCrLf
            End If
        End If
    Next i

    If hasLink Then
        linkData = Left(linkData, Len(linkData) - 3) & vbCrLf ' Remove last comma
    End If

    linkData = linkData & "  ]" & vbCrLf
    linkData = linkData & "};"

    ' Save to slideData.js
    jsonJSPath = outputPath & "slideData.js"
    jsFile = FreeFile
    Open jsonJSPath For Output As #jsFile
    Print #jsFile, linkData
    Close #jsFile

    MsgBox "Slide exported as PNG and metadata saved to slideData.js", vbInformation

End Sub

