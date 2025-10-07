' VBScript to Create a PowerPoint Presentation for the Speech-to-ISL Project

' Declare variables
Dim pptApp, pptPres, pptSld, sldLayout, shpTitle, shpContent, contentText

' Create a PowerPoint application object
Set pptApp = CreateObject("PowerPoint.Application")
pptApp.Visible = True

' Add a new presentation
Set pptPres = pptApp.Presentations.Add

' --- Slide 1: Title Slide ---
Set pptSld = pptPres.Slides.Add(1, 1) ' ppLayoutTitle
Set shpTitle = pptSld.Shapes.Title
shpTitle.TextFrame.TextRange.Text = "Speech-to-ISL Conversion using Transformer-based NLP and MediaPipe Avatar Animation"

Set shpContent = pptSld.Shapes.Placeholders(2)
shpContent.TextFrame.TextRange.Text = "Student Name(s) with ROLL NO." & vbCrLf & "Supervisor Name with Designation" & vbCrLf & "DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING"

' --- Slide 2: Introduction ---
Set pptSld = pptPres.Slides.Add(2, 2) ' ppLayoutText
Set shpTitle = pptSld.Shapes.Title
shpTitle.TextFrame.TextRange.Text = "INTRODUCTION"

Set shpContent = pptSld.Shapes.Placeholders(2)
contentText = "Indian Sign Language (ISL) is a critical communication tool for millions in India's deaf and hard-of-hearing community. " & _
"However, a significant communication gap persists due to a shortage of interpreters and tailored resources, leading to social and educational exclusion. " & _
"While digital technology offers hope, existing solutions often fail to capture the unique grammar and nuances of ISL, focusing instead on Western sign languages like ASL or BSL. " & _
"This project addresses the need for a real-time, automated translation platform specifically designed for the Indian context, converting spoken Hindi and English into accurate ISL gestures."
shpContent.TextFrame.TextRange.Text = contentText
shpContent.TextFrame.TextRange.Font.Size = 28

' --- Slide 3: Project Overview ---
Set pptSld = pptPres.Slides.Add(3, 2) ' ppLayoutText
Set shpTitle = pptSld.Shapes.Title
shpTitle.TextFrame.TextRange.Text = "PROJECT OVERVIEW"

Set shpContent = pptSld.Shapes.Placeholders(2)
contentText = "This project aims to design and develop a real-time Speech to Indian Sign Language (ISL) converter that translates spoken Hindi and English into ISL gestures rendered by a 3D avatar. " & _
"The overall approach is a three-stage pipeline:" & vbCrLf & _
"1. Automatic Speech Recognition (ASR) to convert audio to text, tailored for Indian accents. " & vbCrLf & _
"2. An NLP module to transform the text from standard SVO (Subject-Verb-Object) grammar to ISL's SOV (Subject-Object-Verb) structure. " & vbCrLf & _
"3. A Computer Vision and Animation component to synthesize and render accurate ISL gestures using a 3D avatar."
shpContent.TextFrame.TextRange.Text = contentText
shpContent.TextFrame.TextRange.Font.Size = 24

' --- Slide 4: Existing System ---
Set pptSld = pptPres.Slides.Add(4, 2) ' ppLayoutText
Set shpTitle = pptSld.Shapes.Title
shpTitle.TextFrame.TextRange.Text = "EXISTING SYSTEM"

Set shpContent = pptSld.Shapes.Placeholders(2)
contentText = "Current Systems/Solutions:" & vbCrLf & _
"Many existing speech-to-sign systems focus on American (ASL) or British (BSL) Sign Languages, which differ significantly from ISL in vocabulary and syntax. " & _
"Some ISL systems exist, but are often underdeveloped, not available in real-time, or focus on sign-to-text translation rather than speech-to-sign." & vbCrLf & vbCrLf & _
"Limitations & Challenges:" & vbCrLf & _
"- Lack of accommodation for Indian accents and regional languages, leading to poor ASR accuracy." & vbCrLf & _
"- Inadequate handling of ISL's unique Subject-Object-Verb (SOV) grammar structure." & vbCrLf & _
"- Absence of natural, expressive 3D avatars that include crucial non-manual features like facial expressions." & vbCrLf & _
"- Shortage of standardized, comprehensive ISL gesture datasets for training."
shpContent.TextFrame.TextRange.Text = contentText
shpContent.TextFrame.TextRange.Font.Size = 22

' --- Slide 5: Problem Statement ---
Set pptSld = pptPres.Slides.Add(5, 2) ' ppLayoutText
Set shpTitle = pptSld.Shapes.Title
shpTitle.TextFrame.TextRange.Text = "PROBLEM STATEMENT"

Set shpContent = pptSld.Shapes.Placeholders(2)
contentText = "To design and develop an end-to-end system that bridges the communication gap for the deaf community in India by providing real-time, accurate translation from spoken Hindi and English into Indian Sign Language (ISL) gestures, rendered by an expressive 3D avatar." & vbCrLf & vbCrLf & _
"Importance:" & vbCrLf & _
"Solving this is crucial to combat the social exclusion, educational marginalization, and lack of access to public services faced by millions of deaf individuals in India. An effective tool empowers users, fosters independence, and promotes inclusivity in a digitally connected society."
shpContent.TextFrame.TextRange.Text = contentText
shpContent.TextFrame.TextRange.Font.Size = 24

' --- Slide 6: Objectives and Scope ---
Set pptSld = pptPres.Slides.Add(6, 2) ' ppLayoutText
Set shpTitle = pptSld.Shapes.Title
shpTitle.TextFrame.TextRange.Text = "OBJECTIVES AND SCOPE"

Set shpContent = pptSld.Shapes.Placeholders(2)
contentText = "Key Objectives:" & vbCrLf & _
"1. To develop a robust Automatic Speech Recognition (ASR) module optimized for Indian-accented Hindi and English. " & vbCrLf & _
"2. To implement an NLP pipeline for transforming SVO sentence structure to ISL-compliant SOV grammar. " & vbCrLf & _
"3. To create a comprehensive ISL gesture database from high-quality motion capture data. " & vbCrLf & _
"4. To design and animate a 3D avatar capable of rendering fluid ISL gestures with basic facial expressions." & vbCrLf & vbCrLf & _
"Scope:" & vbCrLf & _
"Included: Real-time translation of spoken Hindi and English, conversational phrases, deployment via a user-friendly interface." & vbCrLf & _
"Excluded: Support for other regional Indian languages, highly technical jargon, complex emotional nuances in facial animation, and sign language to speech conversion."
shpContent.TextFrame.TextRange.Text = contentText
shpContent.TextFrame.TextRange.Font.Size = 22

' --- Slide 7: Literature Survey - Table Header ---
Set pptSld = pptPres.Slides.Add(7, 12) ' ppLayoutTable
Set shpTitle = pptSld.Shapes.Title
shpTitle.TextFrame.TextRange.Text = "LITERATURE SURVEY"

' --- Add a table to the slide ---
Dim tbl
Set tbl = pptSld.Shapes.AddTable(5, 4, 50, 150, 620, 300).Table
' Set headers
tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Author(s) / Year"
tbl.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Title of the Paper"
tbl.Cell(1, 3).Shape.TextFrame.TextRange.Text = "Methodology / Approach"
tbl.Cell(1, 4).Shape.TextFrame.TextRange.Text = "Inference"

' --- Populate Literature Survey Data ---

' Paper 1
tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = "Sonawane et al. [2021]"
tbl.Cell(2, 2).Shape.TextFrame.TextRange.Text = "Speech To Indian Sign Language (ISL) Translation System"
tbl.Cell(2, 3).Shape.TextFrame.TextRange.Text = "Three-stage pipeline: speech recognition, motion data capture using Xbox Kinect, and 3D animation in Unity3D."
tbl.Cell(2, 4).Shape.TextFrame.TextRange.Text = "Demonstrated the feasibility of integrating motion capture with mobile platforms for a real-time ISL translation application."

' Paper 2
tbl.Cell(3, 1).Shape.TextFrame.TextRange.Text = "Kulkarni et al. [2021]"
tbl.Cell(3, 2).Shape.TextFrame.TextRange.Text = "Speech to Indian Sign Language Translator"
tbl.Cell(3, 3).Shape.TextFrame.TextRange.Text = "Utilized computational linguistics and Python-based processing with lemmatization for real-time translation of spoken language to ISL gestures."
tbl.Cell(3, 4).Shape.TextFrame.TextRange.Text = "Proposed a system that can serve as both a practical communication aid and an educational resource, without requiring the user to learn ISL."

' Paper 3
tbl.Cell(4, 1).Shape.TextFrame.TextRange.Text = "Rawat et al. [2025]"
tbl.Cell(4, 2).Shape.TextFrame.TextRange.Text = "A Comprehensive Approach to Indian Sign Language Recognition"
tbl.Cell(4, 3).Shape.TextFrame.TextRange.Text = "Integrated MediaPipe Holistic for landmark extraction and a sequential LSTM model for real-time ISL gesture recognition (sign-to-text)."
tbl.Cell(4, 4).Shape.TextFrame.TextRange.Text = "Achieved high accuracy (96.97%) on a custom dataset, showing the effectiveness of deep learning for sign recognition, though focused on gesture-to-text synthesis."

' Paper 4
tbl.Cell(5, 1).Shape.TextFrame.TextRange.Text = "Chaudhary et al. [2022]"
tbl.Cell(5, 2).Shape.TextFrame.TextRange.Text = "SignNet II: A Transformer-Based Two-Way Sign Language Translation Model"
tbl.Cell(5