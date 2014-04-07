Attribute VB_Name = "questions_and_answers"


Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As Any, ByVal lpParameters As Any, ByVal lpDirectory As Any, ByVal nShowCmd As Long) As Long
Public TmpFRD As Date
Public TmpLRD As Date
Public expiry_date As String



'this function is called when 'ordinary' or 'higher' is clicked
'it takes the text in text_question, adds it onto '.txt', and that
'year, ordinary or higher, is opened in text_questionbox

Public Function open_questions()

Dim stext As String
Dim iFile As Integer
Dim afile As Integer
                                       
iFile = FreeFile
On Error Resume Next
Open App.Path & "\Aural\Questions & Answers\" & Frm_LCF_Test_Yourself.txtauralyear & Frm_LCF_Test_Yourself.Text_Question.Text & ".aye" For Input As iFile
DoEvents

stext = Input(LOF(iFile), iFile)

DoEvents
Close iFile


DoEvents
Frm_LCF_Test_Yourself.Text_JCF_QuestionBox.Text = stext
Frm_LCF_Test_Yourself.Text_JCF_QuestionBox2.Text = stext



errH:
If Err.Number > 0 Then
    MsgBox "Error occurred" & Err.Description, vbCritical + vbSystemModal, "Your Application Title"
    Exit Function
End If

Frm_LCF_Test_Yourself.Text_JCF_QuestionBox.Text = DecryptText(Mid(stext, 13))
Frm_LCF_Test_Yourself.Text_JCF_QuestionBox2.Text = DecryptText(Mid(stext, 13))


End Function





'this function is called when 'ordinary' or 'higher' is clicked
'it takes the text in text_answer, adds it onto '.txt', and that
'year, ordinary or higher, is opened in text_answerbox

Public Function open_answers()

Dim stext As String
Dim iFile As Integer
Dim afile As Integer
                                       
iFile = FreeFile
On Error Resume Next
Open App.Path & "\Aural\Questions & Answers\" & Frm_LCF_Test_Yourself.txtauralyear & Frm_LCF_Test_Yourself.Text_Answer.Text & ".aye" For Input As iFile


On Error Resume Next
stext = Input(LOF(iFile), iFile)
On Error Resume Next
Close iFile

On Error Resume Next
Frm_LCF_Test_Yourself.Text_JCF_AnswerBox.Text = DecryptText(Mid(stext, 13))

On Error Resume Next
End Function





'this function is called to open the transcript.
'it is encrypted in text1, and the webster control takes
'its text from there.

Public Function open_transcript(Year As String)

Dim stext As String
Dim iFile As Integer
Dim afile As Integer
                                       
iFile = FreeFile
On Error Resume Next
Open App.Path & "\Aural\Transcripts\" & Year & "text.aye" For Input As iFile
DoEvents

stext = Input(LOF(iFile), iFile)

DoEvents
Close iFile


DoEvents
Frm_LCF_Learn_and_Revise.Text1.Text = stext




errH:
If Err.Number > 0 Then
    MsgBox "Error occurred" & Err.Description, vbCritical + vbSystemModal, "Your Application Title"
    Exit Function
End If

Frm_LCF_Learn_and_Revise.Text1.Text = DecryptText(Mid(stext, 13))


End Function





'Public Function Show_Answersrt()
'
'Select Case Frm_LCF_Test_Yourself.Show_Answers_Button.Caption
'Case "Show Answers"
'    Frm_LCF_Test_Yourself.Show_Answers_Button.Caption = "Hide Answers"
'    'Form1.Text1.Width = 3455
'    'Form1.Text3.Visible = False
'    Frm_LCF_Test_Yourself.Text_jcf_QuestionBox2.Visible = False
''    Frm_LCF_Test_Yourself.Text_jcf_QuestionBox2.Move 3360, 2160  '13360, 13200
'
'Case "Hide Answers"
'    Frm_LCF_Test_Yourself.Show_Answers_Button.Caption = "Show Answers"
'   ' Form1.Text1.Width = 6655
'   ' Form1.Text3.Visible = True
'    Frm_LCF_Test_Yourself.Text_jcf_QuestionBox2.Visible = True
'
'End Select
'End Function













'**************************************
'**************************************

Public Function DateGood(NumDays As Integer) As Boolean
    'The purpose of this module is to allow you to place a time
    'limit on the unregistered use of your shareware application.
    'This module can not be defeated by rolling back the system clock.
    'Simply call the DateGood function when your application is first
    'loading, passing it the number of days it can be used without
    'registering.
    '
'Ex:     If DateGood(30) = False Then
'     CrippleApplication
'     End If
'    Register LCFrencheters:
'     CRD: Current Run Date
'     LRD: Last Run Date
'     FRD: First Run Date

    Dim TmpCRD As Date
'    Dim TmpLRD As Date
'    Public TmpFRD As Date
    
    Dim Install_date As String
    
    expiry_date = DateAdd("yyyy", 1, TmpFRD)

    Install_date = TmpFRD

    TmpCRD = Date
    TmpLRD = GetSetting("myapp", "LCFrench", "LRD", "1/1/2000")
    TmpFRD = GetSetting("myapp", "LCFrench", "FRD", "1/1/2000")
    
    DateGood = False

    'If this is the applications first load, write initial settings
    'to the register
    If TmpLRD = "1/1/2000" Then
        SaveSetting "myapp", "LCFrench", "LRD", TmpCRD
        SaveSetting "myapp", "LCFrench", "FRD", TmpCRD
'        SaveSetting "myapp", "LCFrench", "timesrun", no_of_times + 5
    End If
    
    'Read LRD and FRD from register
    TmpLRD = GetSetting("myapp", "LCFrench", "LRD", "1/1/2000")
    TmpFRD = GetSetting("myapp", "LCFrench", "FRD", "1/1/2000")
    
'    timesrun = GetSetting("myapp", "LCFrench", "timesrun", no_of_times + 5)

    If TmpLRD > TmpCRD Then 'System clock rolled back
        DateGood = False
    ElseIf Now > DateAdd("d", NumDays, TmpFRD) Then 'Expiration expired
        DateGood = False
    ElseIf TmpCRD > TmpLRD Then 'Everything OK write New LRD date
        SaveSetting "myapp", "LCFrench", "LRD", TmpCRD
'        SaveSetting "myapp", "LCFrench", "timesrun", no_of_times + 5
        DateGood = True
    ElseIf TmpCRD = Date Then
        DateGood = True
    Else
        DateGood = False
    End If
End Function





'this function is called to open the dictionary, with the ultimate objective of
'printing it.
'it is encrypted in text1, and the webster control takes
'its text from there.

Public Function print_dictionary()
'Public Function open_dictionary(Year As String)

Dim stext As String
Dim iFile As Integer
Dim afile As Integer
                                       
iFile = FreeFile
On Error Resume Next
Open App.Path & "\Aural\Transcripts\dictionary" & Frm_LCF_Learn_and_Revise.Text2.Text & "_print.aye" For Input As iFile
DoEvents
'& Year & "_print.aye"
stext = Input(LOF(iFile), iFile)

DoEvents
Close iFile


DoEvents
Frm_LCF_Dictionary.Text1.Text = stext




errH:
If Err.Number > 0 Then
    MsgBox "Error occurred" & Err.Description, vbCritical + vbSystemModal, "Your Application Title"
    Exit Function
End If

Frm_LCF_Dictionary.Text1.Text = DecryptText(Mid(stext, 13))


End Function








