'
'  https://stackoverflow.com/a/29750646/180275
'
Option Explicit

Sub CreatePowerPointQuestions()

 'Add a reference to the Microsoft PowerPoint Library by:
'1. Go to Tools in the VBA menu
'2. Click on Reference
'3. Scroll down to Microsoft PowerPoint X.0 Object Library, check the box, and press Okay

'First we declare the variables we will be using
    Dim newPowerPoint As PowerPoint.Application
    Dim activeSlide As PowerPoint.Slide
    Dim Question As String
    Dim Options As String 'comma separated list of options
    Dim Choices() As String 'split up options for printing
    Dim printOptions As String 'string to print in contents of slide
    Dim Answer As String
    Dim limit As Integer
'set question amount:
    limit = 5
    
 'Look for existing instance
    On Error Resume Next
    Set newPowerPoint = GetObject(, "PowerPoint.Application")
    On Error GoTo 0

'Let's create a new PowerPoint
    If newPowerPoint Is Nothing Then
        Set newPowerPoint = New PowerPoint.Application
    End If
'Make a presentation in PowerPoint
    If newPowerPoint.Presentations.Count = 0 Then
        newPowerPoint.Presentations.Add
    End If

'Show the PowerPoint
    newPowerPoint.Visible = True
    
    
'Select worksheet and cells activate
    ' Worksheets("Sheet1").Activate
    Worksheets(1).Activate

'Loop through each question
Dim i As Integer
    For i = 1 To limit

    'Add a new slide where we will paste the Question and Options:
        newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
        newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
        Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)

    'Set the variables to the cells
        Question = ActiveSheet.Cells(i, 1).Value
        Options = ActiveSheet.Cells(i, 2).Value
        Answer = ActiveSheet.Cells(i, 3).Value

    'Split options into choices a,b,c,d based on comma separation
        Choices() = Split(Options, ",")
    'Formate printOptions to paste into content
        printOptions = "A) " & Choices(0) & vbNewLine & "B) " & Choices(1) & vbNewLine & "C) " & Choices(2) & vbNewLine & "D) " & Choices(3)
        activeSlide.Shapes(2).TextFrame.TextRange.Text = printOptions

    'Set the title of the slide the same as the question for the options
        activeSlide.Shapes(1).TextFrame.TextRange.Text = Question

    'Add answer slide and select it
        newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
        newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
        Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)
    'Set title:
        activeSlide.Shapes(1).TextFrame.TextRange.Text = "Answer:"
    'Set contents to answer:
        activeSlide.Shapes(2).TextFrame.TextRange.Text = Answer
    'Finished with a row (question)
    Next

AppActivate ("Microsoft PowerPoint")
Set activeSlide = Nothing
Set newPowerPoint = Nothing

End Sub

