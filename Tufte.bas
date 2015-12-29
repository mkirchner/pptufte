Attribute VB_Name = "Tufte"
'
' tufte_macros.bas
'
' Copyright (C) 2015 Marc Kirchner
'
' Permission is hereby granted, free of charge, to any person obtaining a
' copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation
' the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the
' Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'

Option Explicit
Public Type TufteConfig
    SlideWidth As Single
    SlideHeight As Single
    
    BackgroundColor As Long
    TextColor As Long
    DefaultFont As String
    MinKerningSize As Integer
        
    SlideMarginTop As Single
    SlideMarginBottom As Single
    SlideMarginLeft As Single
    SlideMarginRight As Single
    
    PageWidth As Single
    PageHeight As Single
    
    TitleTop As Single
    TitleLeft As Single
    TitleHeight As Single
    TitleWidth As Single
    TitleDefaultTextSize As Integer
    
    SubtitleTop As Single
    SubtitleLeft As Single
    SubtitleHeight As Single
    SubtitleWidth As Single
    SubtitleDefaultTextSize As Integer
        
    CanvasTop As Single
    CanvasHeight As Single
    CanvasLeft As Single
    CanvasWidth As Single
    CanvasMarginLeft As Single
    CanvasMarginRight As Single
    CanvasDefaultTextSize As Integer
    
    MarginTop As Single
    MarginLeft As Single
    MarginWidth As Single
    MarginHeight As Single
    MarginDefaultTextSize As Integer
End Type

Public Type FigureComponents
    caption As Shape
    obj As Shape
    group As Shape
    isGrouped As Boolean
End Type

Public Enum CaptionType
    CaptionTypeTitle
    CaptionTypeSubtitle
End Enum

Function GetConfig() As TufteConfig
    Dim c As TufteConfig
    c.SlideHeight = ActivePresentation.SlideMaster.Height
    c.SlideWidth = ActivePresentation.SlideMaster.Width
    c.SlideMarginLeft = c.SlideWidth / 4 / 2
    c.SlideMarginRight = c.SlideMarginLeft
    c.PageWidth = c.SlideWidth - c.SlideMarginLeft - c.SlideMarginRight
    
    Const marginRatio As Single = 0.3333
    Dim midMarginWidth As Single
    midMarginWidth = c.PageWidth * 0.07
    
    ' lefts and widths
    c.CanvasLeft = c.SlideMarginLeft
    c.CanvasWidth = c.PageWidth * (1 - 0.07) * (1 - marginRatio)
    c.MarginLeft = c.SlideMarginLeft + c.CanvasWidth + midMarginWidth
    c.MarginWidth = c.CanvasWidth / 2
    
    ' tops and heights
    c.SlideMarginTop = c.SlideHeight * 0.125 / 2.618
    c.SlideMarginBottom = c.SlideMarginTop * 1.618 ' golden ratio
    c.PageHeight = c.SlideHeight * (1 - 0.125)
    c.MarginTop = c.SlideMarginTop
    c.MarginHeight = c.PageHeight
    c.CanvasTop = c.MarginTop
    c.CanvasHeight = c.PageHeight
    
    'title
    c.TitleDefaultTextSize = 36
    c.TitleTop = c.SlideMarginTop
    c.TitleHeight = 0.12 * c.SlideHeight
    c.TitleLeft = c.CanvasLeft
    c.TitleWidth = c.CanvasWidth
    
    'subtitle
    c.SubtitleDefaultTextSize = 14
    c.SubtitleHeight = 0.045 * c.SlideHeight
    c.SubtitleTop = 80
    c.SubtitleLeft = c.CanvasLeft
    c.SubtitleWidth = c.CanvasWidth
    
    ' other
    c.BackgroundColor = RGB(255, 255, 252)
    c.TextColor = RGB(17, 17, 17)
    c.CanvasDefaultTextSize = 12
    c.MarginDefaultTextSize = 10
    c.MinKerningSize = 1
    
    GetConfig = c
    Exit Function
End Function

Function GetCursorYPosition() As Single
    Dim last As Integer
    Dim tp, hght As Single
    With ActiveWindow.selection
        If .Type = ppSelectionText Then
            If Len(.TextRange) = 0 Then
                tp = .ShapeRange(1).top
                hght = .ShapeRange(1).Height
                last = .TextRange.Start
            End If
        End If
    End With
    Dim tr As TextRange
    Dim currentLine, nLines As Integer
    Set tr = ActiveWindow.selection.ShapeRange(1).TextFrame.TextRange
    nLines = tr.Lines.Count
    currentLine = tr.Characters(1, last).Lines.Count
    GetCursorYPosition = tp + (hght / nLines) * (currentLine - 1)
    Exit Function
End Function

Sub Auto_Open()
    Dim oToolbar As CommandBar
    Dim Button1, Button2, Button3, Button4, Button5, Button6, Button7, Button8, Button9 As CommandBarButton
    Dim toolbarName As String

    Debug.Print "Attempting to load Tufte Addin"
    toolbarName = "Tufte Tools"

    On Error Resume Next

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(name:=toolbarName, _
        Position:=msoBarTop, Temporary:=True)
    
    If Err.Number <> 0 Then
          ' The toolbar's already there, so we have nothing to do
          MsgBox "Tufte Toolbar has already been loaded. Exiting."
          Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' Now add a button to the new toolbar
    Set Button1 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button1
         .DescriptionText = "Description"
         .caption = "New slide"
         .OnAction = "TufteCreateSlide"
         .Style = msoButtonIcon
         .FaceId = 583
    End With
    Set Button7 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button7
         .DescriptionText = "Description"
         .caption = "Make title"
         .OnAction = "TufteMakeTitle"
         .Style = msoButtonIcon
         .FaceId = 598
    End With
    Set Button9 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button9
         .DescriptionText = "Description"
         .caption = "Make subtitle"
         .OnAction = "TufteMakeSubtitle"
         .Style = msoButtonIcon
         .FaceId = 599
    End With
    Set Button6 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button6
         .DescriptionText = "Description"
         .caption = "Make canvas"
         .OnAction = "TufteMakeCanvas"
         .Style = msoButtonIcon
         .FaceId = 7
    End With
    Set Button5 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button5
         .DescriptionText = "Description"
         .caption = "Make canvas figure"
         .OnAction = "TufteMakeCanvasFigure"
         .Style = msoButtonIcon
         .FaceId = 931
    End With
    Set Button3 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button3
         .DescriptionText = "Description"
         .caption = "Make margin figure"
         .OnAction = "TufteMakeMarginFigure"
         .Style = msoButtonIcon
         .FaceId = 218
    End With
    Set Button2 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button2
         .DescriptionText = "Description"
         .caption = "Make margin note"
         .OnAction = "TufteMakeMarginNote"
         .Style = msoButtonIcon
         .FaceId = 244
    End With
    Set Button4 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button4
         .DescriptionText = "Description"
         .caption = "Make referenced margin note"
         .OnAction = "TufteMakeReferencedMarginNote"
         .Style = msoButtonIcon
         .FaceId = 246
    End With
    Set Button8 = oToolbar.Controls.Add(Type:=msoControlButton)
    With Button8
         .DescriptionText = "Description"
         .caption = "Auto-layout selected canvas objects"
         .OnAction = "TufteAutolayoutCanvas"
         .Style = msoButtonIcon
         .FaceId = 144
    End With

    oToolbar.Visible = True

NormalExit:
    Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCrLf & Err.Description
     Resume NormalExit:
End Sub

Sub TufteCreateSlide()
    '''
    ' Creates a new slide using #fffff8 as a background color.
    '''
    Dim c As TufteConfig
    c = GetConfig
    Dim currentSlide As Slide
    Set currentSlide = Nothing
    Dim currentView As View
    Dim currentIndex As Long
    With ActivePresentation.Slides
        Set currentView = ActiveWindow.View
        currentIndex = 0
        If .Count > 0 Then
            currentIndex = currentView.Slide.SlideIndex
        End If
        Set currentSlide = .Add(currentIndex + 1, ppLayoutBlank)
    End With
    With currentSlide
        .FollowMasterBackground = False
        .Background.Fill.Solid
        .Background.Fill.ForeColor.RGB = c.BackgroundColor
    End With
    currentView.GotoSlide (currentSlide.SlideIndex)
End Sub

Function TufteGetSelectedText(keepText As Boolean) As String
    '''
    ' Get selected text or return an empty string.
    '
    ' :param keepText: If false, the function deletes the
    '                  selected text
    ' :type keepText: Boolean
    '''
    Dim txt As String: txt = ""
    With ActiveWindow.selection
        If .Type = ppSelectionText Then
            If Len(.TextRange) <> 0 Then
                txt = .TextRange.text
                If Not keepText Then
                    .Delete
                End If
            End If
        End If
    End With
    TufteGetSelectedText = txt
    Exit Function
End Function



Sub TufteMakeTitle()
    '''
    ' Creates a title box (from selected text, if available)
    '''
    Dim currentSlide As Slide
    Set currentSlide = Application.ActiveWindow.View.Slide
    Dim c As TufteConfig
    c = GetConfig
    Dim shp As Shape
    Set shp = TufteCreateCaption(currentSlide, CaptionTypeTitle, "Title", c)
End Sub

Sub TufteMakeSubTitle()
    '''
    ' Creates a subtitle box (from selected text, if available)
    '''
    Dim currentSlide As Slide
    Set currentSlide = Application.ActiveWindow.View.Slide
    Dim c As TufteConfig
    c = GetConfig
    Dim shp As Shape
    Set shp = TufteCreateCaption(currentSlide, CaptionTypeSubtitle, "Subtitle", c)
End Sub


Function TufteCreateCaption(currentSlide As Slide, ctype As CaptionType, _
                            defaultText As String, c As TufteConfig) As Shape
    Dim shp As Shape
    Dim shapeMode As Boolean: shapeMode = False
    Dim text As String
    ' check if there is a selected shape and no text is selected
    With ActiveWindow.selection
        If .Type = ppSelectionShapes Then
            If .ShapeRange(1).HasTextFrame And .ShapeRange(1).TextFrame.HasText Then
                Set shp = .ShapeRange(1)
                shapeMode = True
            End If
        Else
            ' check if there is selected text
            text = TufteGetSelectedText(True)
            If text = "" Then
                text = defaultText
            End If
        End If
    End With
    Dim cLeft, cTop, cWidth, cHeight, textSize As Single
    Dim name As String
    If ctype = CaptionTypeTitle Then
        cLeft = c.TitleLeft
        cTop = c.TitleTop
        cWidth = c.TitleWidth
        cHeight = c.TitleHeight
        textSize = c.TitleDefaultTextSize
        name = "tufte:title"
    Else
        cLeft = c.SubtitleLeft
        cTop = c.SubtitleTop
        cWidth = c.SubtitleWidth
        cHeight = c.SubtitleHeight
        textSize = c.SubtitleDefaultTextSize
        name = "tufte:subtitle"
    End If

    With currentSlide.Shapes
        If Not shapeMode Then
            Set shp = .AddTextbox(msoTextOrientationHorizontal, _
                                  cLeft, cTop, cWidth, cHeight)
        Else
            shp.Left = cLeft
            shp.top = cTop
            shp.Width = cWidth
            shp.Height = cHeight
        End If
            shp.TextFrame.AutoSize = ppAutoSizeNone
        With shp.TextFrame.TextRange
            If Not shapeMode Then
                .text = text
            End If
            .ParagraphFormat.Alignment = ppAlignLeft
            .Font.Italic = msoTrue
            .Font.Color.RGB = c.TextColor
            .Font.Size = textSize
        End With
        With shp.TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .AutoSize = ppAutoSizeShapeToFitText
        End With
        shp.Height = cHeight
        shp.Tags.Add "tufte:type", name
    End With
    Set TufteCreateCaption = shp
    Exit Function
End Function

Function TufteFindNextTop() As Integer
    ' adds up all canvas shape heights and returns the result
    Dim shp As Shape
    Dim currentSlide As Slide
    Dim pos, bottom As Integer
    Dim tufteType As String
    pos = 1
    Set currentSlide = Application.ActiveWindow.View.Slide
    For Each shp In currentSlide.Shapes
        tufteType = shp.Tags("tufte:type")
        If tufteType = "tufte:canvas" Or _
           tufteType = "tufte:title" Or _
           tufteType = "tufte:subtitle" Or _
           tufteType = "tufte:figuregroup" Then
            bottom = shp.top + shp.Height
            If bottom > pos Then
                pos = bottom
            End If
        End If
    Next shp
    TufteFindNextTop = pos + 1
End Function

Sub TufteMakeCanvas()
    '''
    ' Converts a selected textbox into a Tufte canvas or creates a new
    ' canvas if nothing appropriate is selected.
    '''
    Dim c As TufteConfig
    c = GetConfig
        
    Dim canvas As Shape
    Dim shapeMode As Boolean: shapeMode = False
    ' Check if there is a selected shape that should be converted into a
    ' Tufte canvas
    With ActiveWindow.selection
        If .Type = ppSelectionShapes Then
            If .ShapeRange(1).HasTextFrame And .ShapeRange(1).TextFrame.HasText Then
                Set canvas = .ShapeRange(1)
                shapeMode = True
                canvas.Tags.Add "tufte:type", "tufte:none" 'get current obj out of the way for calculations
            End If
        End If
        
        If .Type <> ppSelectionShapes Then
            ' create shape
            Dim currentSlide As Slide
            Set currentSlide = Application.ActiveWindow.View.Slide
            With currentSlide.Shapes
                Set canvas = .AddTextbox(msoTextOrientationHorizontal, _
                           c.CanvasLeft, c.CanvasTop, c.CanvasWidth, c.CanvasHeight)
            End With
        End If
    End With
        
    ' find proper top and height
    Dim top As Integer
    top = TufteFindNextTop
    If top < c.CanvasTop Then
        top = c.CanvasTop
    End If
    
    With canvas
        .TextFrame.AutoSize = ppAutoSizeNone
        .TextFrame.WordWrap = msoTrue
        .top = top
        .Left = c.CanvasLeft
        .Width = c.CanvasWidth
        .Height = c.CanvasHeight - top
        With .TextFrame.TextRange
            If Not shapeMode Then
                .text = "Canvas text"
            End If
            .ParagraphFormat.Alignment = ppAlignLeft
            .Font.Italic = msoFalse
            .Font.Color.RGB = c.TextColor
            .Font.Size = c.CanvasDefaultTextSize
        End With
        With canvas.TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .AutoSize = ppAutoSizeShapeToFitText
        End With
        .Tags.Add "tufte:type", "tufte:canvas"
    End With
    If Not shapeMode Then
        canvas.TextFrame.TextRange.Select
    End If
End Sub

Sub TufteMakeMarginNote()
    Dim mn As Shape
    Dim text As String
    Dim selectText As Boolean: selectText = False
    text = TufteGetSelectedText(False)
    If text = "" Then
        text = "Margin note."
        selectText = True
    End If
    Set mn = TufteCreateMarginNote(text)
    If selectText Then
        mn.TextFrame.TextRange.Select
    End If
End Sub

Function TufteCreateMarginNote(Optional txt As String) As Shape
    Dim c As TufteConfig
    c = GetConfig
    
    If IsMissing(txt) Then
        txt = "Margin Note."
    End If
    
    Dim currentSlide As Slide
    Dim marginNote As Shape
    Set currentSlide = Application.ActiveWindow.View.Slide
    
    Dim y As Single: y = c.MarginHeight / 2
    If ActiveWindow.selection.Type = ppSelectionShapes Then
        Set marginNote = ActiveWindow.selection.ShapeRange(1)
    Else
        If ActiveWindow.selection.Type = ppSelectionText Then
            y = GetCursorYPosition
        End If
        Set marginNote = currentSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            c.MarginLeft, y, c.MarginWidth, 100) 'Height changes dynamically
        marginNote.TextFrame.TextRange.text = txt
    End If
    marginNote.TextFrame.AutoSize = ppAutoSizeNone
    With marginNote
        .top = y
        .Left = c.MarginLeft
        .Width = c.MarginWidth
    End With
    With marginNote.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .TextRange.Font.Size = c.MarginDefaultTextSize
        .WordWrap = msoTrue
        .AutoSize = ppAutoSizeShapeToFitText
    End With
    marginNote.Tags.Add "tufte:type", "tufte:marginnote"
    Set TufteCreateMarginNote = marginNote
    Exit Function
End Function

Sub TufteMakeMarginFigure()
    Dim c As TufteConfig
    c = GetConfig
    Dim currentSlide As Slide
    Set currentSlide = Application.ActiveWindow.View.Slide

    Dim figureParts As FigureComponents
    figureParts = TufteGetFigureComponents("Figure caption.")
    With figureParts
        ' format caption and object
        With .caption.TextFrame.TextRange
            .ParagraphFormat.Alignment = ppAlignLeft
            .Font.Italic = msoFalse
            .Font.Color.RGB = c.TextColor
            .Font.Size = c.MarginDefaultTextSize
        End With
        With .caption.TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .WordWrap = msoTrue
            .AutoSize = ppAutoSizeShapeToFitText
        End With
        With .caption
            .Left = c.MarginLeft
            .Width = c.MarginWidth
        End With
        With .obj
            .LockAspectRatio = msoTrue
            .Width = c.MarginWidth
            .Left = c.MarginLeft
        End With
        ' align caption w/ object
        .caption.top = .obj.top + .obj.Height
        ' group if necessary
        Dim shp As Shape
        If Not .isGrouped Then
            Set shp = currentSlide.Shapes.Range(Array(.caption.name, .obj.name)).group
            .caption.TextFrame.TextRange.Select
        Else
            Set shp = .group
        End If
        shp.Tags.Add "tufte:type", "tufte:marginfiguregroup"
    End With
ErrorHandler:
End Sub


Function TufteGetFigureComponents(defaultCaption As String) As FigureComponents
    Dim c As TufteConfig
    c = GetConfig
    Dim currentSlide As Slide
    Set currentSlide = Application.ActiveWindow.View.Slide
    Dim shp, caption, obj As Shape
    Dim isGrouped As Boolean
    If ActiveWindow.selection.ShapeRange.Count = 1 Then
        Set shp = ActiveWindow.selection.ShapeRange(1)
        Select Case shp.Type
            Case msoGroup
                ' assume we have a group w/ a text box and an object
                With shp
                    ' we only support textbox + object
                    If .GroupItems.Count > 2 Then
                        MsgBox "Too many elements in Group; please pre-group to have text box and object.", vbOKOnly
                        GoTo ErrorHandler
                    End If
                    isGrouped = True
                    Dim item As Integer
                    For item = 1 To .GroupItems.Count
                        Select Case .GroupItems(item).Type
                            Case msoTextBox
                                ' this is the figure caption
                                Set caption = .GroupItems(item)
                            Case Else
                                ' this is the object
                                Set obj = .GroupItems(item)
                        End Select
                    Next item
                End With
            Case msoChart, msoDiagram, msoEmbeddedOLEObject, msoFreeform, _
                 msoGroup, msoLine, msoLinkedOLEObject, msoLinkedPicture, _
                 msoPicture, msoPlaceholder, msoTable, msoTextEffect
                Set obj = shp
                'Create caption
                Set caption = currentSlide.Shapes.AddTextbox( _
                                msoTextOrientationHorizontal, c.MarginLeft, _
                                c.CanvasHeight / 2, c.MarginWidth, 100)
                With caption.TextFrame.TextRange
                    .text = defaultCaption
                End With
            Case Else
                MsgBox "Cannot convert selection into Tufte-style figure."
                GoTo ErrorHandler
        End Select
    ElseIf ActiveWindow.selection.ShapeRange.Count = 2 Then
        ' again, textbox and object, but not grouped
        For item = 1 To 2
            With ActiveWindow.selection
            Select Case .ShapeRange(item).Type
                Case msoTextBox
                    ' this is the figure caption
                    Set caption = .ShapeRange(item)
                Case Else
                    ' this is the object
                    Set obj = .ShapeRange(item)
            End Select
            End With
        Next item
    End If
    Dim retVal As FigureComponents
    Set retVal.caption = caption
    Set retVal.obj = obj
    If isGrouped Then
        Set retVal.group = shp
    End If
    retVal.isGrouped = isGrouped
    TufteGetFigureComponents = retVal
    Exit Function
ErrorHandler:
End Function

Sub TufteMakeCanvasFigure()
    Dim c As TufteConfig
    c = GetConfig
    Dim currentSlide As Slide
    Set currentSlide = Application.ActiveWindow.View.Slide

    Dim figureParts As FigureComponents
    figureParts = TufteGetFigureComponents("Figure caption.")
    With figureParts
        ' format caption and object
        With .caption.TextFrame.TextRange
            .ParagraphFormat.Alignment = ppAlignLeft
            .Font.Italic = msoFalse
            .Font.Color.RGB = c.TextColor
            .Font.Size = c.MarginDefaultTextSize
        End With
        With .caption.TextFrame
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .WordWrap = msoTrue
            .AutoSize = ppAutoSizeShapeToFitText
        End With
        With .caption
            .Left = c.MarginLeft
            .Width = c.MarginWidth
        End With
        With .obj
            .LockAspectRatio = msoTrue
            .Width = c.CanvasWidth
            .Left = c.CanvasLeft
        End With
        ' align caption w/ object
        .caption.top = .obj.top
        ' group if necessary
        If Not .isGrouped Then
            currentSlide.Shapes.Range(Array(.caption.name, .obj.name)).group
        End If
    End With
ErrorHandler:
End Sub


Sub TufteAutolayoutCanvas()
    Dim c As TufteConfig
    c = GetConfig
    Dim currentSlide As Slide
    Dim shp As Shape
    Set currentSlide = Application.ActiveWindow.View.Slide
    'limit number of shapes because VBA sucks
    If ActiveWindow.selection.ShapeRange.Count > 100 Then
        MsgBox "Limited to 100 selected shapes per slide."
    End If
        
    'iterate over all selected shapes and adjust sizes
    Dim i As Integer
    Dim indexes(100) As Integer
    Dim tops(100) As Single
    For i = ActiveWindow.selection.ShapeRange.Count To 1 Step -1
        Set shp = ActiveWindow.selection.ShapeRange(i)
        If shp.Type = msoTextBox Then
            With shp.TextFrame.TextRange
                .ParagraphFormat.Alignment = ppAlignLeft
                .Font.Italic = msoFalse
                .Font.Color.RGB = c.TextColor
                .Font.Size = c.CanvasDefaultTextSize
            End With
            With shp.TextFrame
                .MarginLeft = 0
                .MarginRight = 0
                .WordWrap = msoTrue
                .AutoSize = ppAutoSizeShapeToFitText
            End With
            shp.Left = c.CanvasLeft
            shp.Width = c.CanvasWidth
        End If
        
        If shp.Type = msoGroup Then 'TODO: pull this out
            With shp
                ' we only support textbox + object
                If .GroupItems.Count > 2 Then
                    MsgBox "Too many elements in Group; please pre-group to have text box and object.", vbOKOnly
                End If
                Dim item As Integer
                Dim caption As Shape
                Dim obj As Shape
                For item = 1 To .GroupItems.Count
                    Select Case .GroupItems(item).Type
                        Case msoTextBox
                            ' this is the figure caption
                            With .GroupItems(item).TextFrame.TextRange
                                .ParagraphFormat.Alignment = ppAlignLeft
                                .Font.Italic = msoFalse
                                .Font.Color.RGB = c.TextColor
                                .Font.Size = c.MarginDefaultTextSize
                            End With
                            With .GroupItems(item).TextFrame
                                .MarginLeft = 0
                                .MarginRight = 0
                                .MarginTop = 0
                                .WordWrap = msoTrue
                                .AutoSize = ppAutoSizeShapeToFitText
                            End With
                            With .GroupItems(item)
                                .Left = c.MarginLeft
                                .Width = c.MarginWidth
                            End With
                            Set caption = .GroupItems(item)
                        Case Else
                            ' this is the object
                            With .GroupItems(item)
                                .LockAspectRatio = msoTrue
                                .Width = c.CanvasWidth
                                .Left = c.CanvasLeft
                            End With
                            Set obj = .GroupItems(item)
                    End Select
                Next item
                ' align caption w/ object
                caption.top = obj.top
            End With
        End If
        tops(i) = shp.top
        indexes(i) = i
    Next i
    ' bubble sort by top
    Dim r, l As Integer
    Dim tmpTop As Single
    Dim tmpIndex As Integer
    For r = ActiveWindow.selection.ShapeRange.Count To 1 Step -1
        For l = 1 To r - 1
            If tops(l) > tops(l + 1) Then
                tmpTop = tops(l + 1)
                tops(l + 1) = tops(l)
                tops(l) = tmpTop
                tmpIndex = indexes(l + 1)
                indexes(l + 1) = indexes(l)
                indexes(l) = tmpIndex
            End If
        Next l
    Next r
    'put shapes one after another
    Dim k As Integer
    Dim currentTop As Single
    currentTop = c.CanvasTop
    ' check if there is a title and/or subtitle
    For Each shp In currentSlide.Shapes
        If shp.Tags("tufte:type") = "tufte:title" Then
            currentTop = currentTop + shp.Height
        ElseIf shp.Tags("tufte:type") = "tufte:subtitle" Then
            currentTop = currentTop + shp.Height
        End If
    Next shp
    For k = 1 To ActiveWindow.selection.ShapeRange.Count
        With ActiveWindow.selection.ShapeRange(indexes(k))
            .top = currentTop
            currentTop = currentTop + .Height + 0.01 * c.CanvasHeight
            Debug.Print currentTop
        End With
    Next k
    
ErrorExit:
End Sub

Sub TufteMakeReferencedMarginNote()
    Dim c As TufteConfig
    c = GetConfig
    Dim currentSlide As Slide
    Dim shp As Shape
    Set currentSlide = Application.ActiveWindow.View.Slide
    
    'count margin notes on this slide
    Dim refnr As Integer: refnr = 1
    For Each shp In currentSlide.Shapes
        If shp.Tags("tufte:type") = "tufte:referencedmarginnote" Then
            refnr = refnr + 1
        End If
    Next shp
    
    ' go the easy way (FIXME!): disallow more than 9 margin notes (that actually makes sense)
    If refnr > 9 Then
        MsgBox "More than 9 referenced margin notes on a single slide are currently not supported.", vbExclamation
        GoTo ErrorExit
    End If
    
    Dim txRng As TextRange
    On Error Resume Next
    If ActiveWindow.selection.Type = ppSelectionText Then
        If Len(ActiveWindow.selection.TextRange) = 0 Then
            Set txRng = ActiveWindow.selection.TextRange.InsertSymbol( _
                        FontName:=ActiveWindow.selection.TextRange.Font.name, _
                        CharNumber:=48 + refnr, Unicode:=msoTrue)
            txRng.Font.Superscript = True
            ' add margin note
            Dim mn As Shape
            Set mn = TufteCreateMarginNote("Margin note.")
            mn.Tags.Add "tufte:type", "tufte:referencedmarginnote"
            With mn.TextFrame.TextRange
                .InsertBefore (Chr$(48 + refnr))
                .Characters(1).Font.Superscript = msoTrue
                .Characters(2, .Length).Select
            End With
        End If
    End If
ErrorExit:
End Sub
