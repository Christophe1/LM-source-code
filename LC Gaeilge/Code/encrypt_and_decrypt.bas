Attribute VB_Name = "encrypt_and_decrypt"




'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''








'This is code for the 'Encryption Panel', allows us to encrypt
'and protect files. They can only be viewed within Lingua-Master

'It uses the Decrypt function, above, and the following
'Encrypt function



'encrypt function
Function EncryptText(ByVal stext As String) As String
Dim sKey As String
Dim iKey As Integer
Dim lPos As Long
Dim iChar As Integer
Randomize Timer
sKey = Right("000" & Hex(Int(Rnd(1) * 4095)), 3)
For lPos = 1 To Len(stext)
iChar = Asc(Mid$(stext, lPos, 1))
Mid$(stext, lPos, 1) = Chr(iChar + Val(Mid(iKey + 1, 1)))
iKey = (iKey + 1) Mod 3
Next
EncryptText = "Lingua-Master" & stext & Right(Space(20) & (Val("&H" & sKey) + Len(stext)), 20)
                                       
End Function


'This is the Decrypt function for the questions and answers
Function DecryptText(ByVal stext As String) As String
Dim sKey As String
Dim iKey As Integer
Dim lPos As Long
Dim iChar As Integer
sKey = Right("000" & Hex(Val(Right(stext, 20)) - (Len(stext) - 20)), 3)
For lPos = 1 To Len(stext) - 20
iChar = Asc(Mid$(stext, lPos, 1))
Mid$(stext, lPos, 1) = Chr(iChar - Val(Mid(iKey + 1, 1)))
iKey = (iKey + 1) Mod 3
Next
DecryptText = Left(stext, Len(stext) - 20)
End Function








'opens a regular text file. This can then be saved
'as a text file or as an encrypted file


Function Open_Text_file()
' Set CancelError to True
Frm_LCI_Test_Yourself.CommonDialog1.CancelError = True
Frm_LCI_Test_Yourself.CommonDialog1.Flags = cdlOFNFileMustExist
On Error GoTo ErrHandler
Frm_LCI_Test_Yourself.CommonDialog1.Flags = cdlOFNFileMustExist
' Specify default filter
'Set the properties of the text control
Frm_LCI_Test_Yourself.CommonDialog1.Filter = "Text Files(*.txt)|*.txt*|All files(*.*)|*.*"
'.Filter = "AY Encrypted|*.aye"
Frm_LCI_Test_Yourself.CommonDialog1.DefaultExt = "txt"
Frm_LCI_Test_Yourself.CommonDialog1.FilterIndex = 1
Frm_LCI_Test_Yourself.CommonDialog1.ShowOpen

Open Frm_LCI_Test_Yourself.CommonDialog1.FileName For Input As #1
Frm_LCI_Test_Yourself.Text_JCF_QuestionBox.Text = Input(LOF(1), 1)
Close #1

ErrHandler:
'User pressed the Cancel button
Exit Function
End Function


'opens an encrypted file. This can then be saved
'as a text file or as an encrypted file

Function Open_Encrypted_file()
'Open Encrypted File
Dim stext As String
Dim iFile As Integer
                                       
On Error GoTo Cancelled
With Frm_LCI_Test_Yourself.CommonDialog1
.CancelError = True
.Filter = "AY Encrypted|*.aye*|All files(*.*)|*.*"
'CommonDialog1.Filter = "Text Files(*.txt)|*.txt*|All files(*.*)|*.*"
.ShowOpen
iFile = FreeFile
Open .FileName For Input As iFile
stext = Input(LOF(iFile), iFile)
Close iFile
End With
'If Left(sText, 12) <> "AY Encrypted" Then
'MsgBox "Not an AY Encrypted File Format"
'Else
Frm_LCI_Test_Yourself.Text_JCF_QuestionBox.Text = DecryptText(Mid(stext, 13))
'End If
                                       
Cancelled:
End Function



'saves a text file or encrypted file as an encrypted file

Function save_encrypted()

On Error GoTo Cancelled
With Frm_LCI_Test_Yourself.CommonDialog1
.CancelError = True
.Filter = "AY Encrypted|*.aye"
.ShowSave
iFile = FreeFile
Open .FileName For Output As iFile
Print #iFile, EncryptText(Frm_LCI_Test_Yourself.Text_JCF_QuestionBox);
Close iFile
End With
MsgBox "Saved"
Cancelled:

End Function










''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''






'opens a regular text file. This can then be saved
'as a text file or as an encrypted file


Function Open_Text_file2()
' Set CancelError to True
Frm_LCI_Test_Yourself.CommonDialog1.CancelError = True
Frm_LCI_Test_Yourself.CommonDialog1.Flags = cdlOFNFileMustExist
On Error GoTo ErrHandler
Frm_LCI_Test_Yourself.CommonDialog1.Flags = cdlOFNFileMustExist
' Specify default filter
'Set the properties of the text control
Frm_LCI_Test_Yourself.CommonDialog1.Filter = "Text Files(*.txt)|*.txt*|All files(*.*)|*.*"
'.Filter = "AY Encrypted|*.aye"
Frm_LCI_Test_Yourself.CommonDialog1.DefaultExt = "txt"
Frm_LCI_Test_Yourself.CommonDialog1.FilterIndex = 1
Frm_LCI_Test_Yourself.CommonDialog1.ShowOpen

Open Frm_LCI_Test_Yourself.CommonDialog1.FileName For Input As #1
Frm_LCI_Test_Yourself.Text_JCF_AnswerBox.Text = Input(LOF(1), 1)
Close #1

ErrHandler:
'User pressed the Cancel button
Exit Function
End Function


'opens an encrypted file. This can then be saved
'as a text file or as an encrypted file

Function Open_Encrypted_file2()
'Open Encrypted File
Dim stext As String
Dim iFile As Integer
                                       
On Error GoTo Cancelled
With Frm_LCI_Test_Yourself.CommonDialog1
.CancelError = True
.Filter = "AY Encrypted|*.aye*|All files(*.*)|*.*"
'CommonDialog1.Filter = "Text Files(*.txt)|*.txt*|All files(*.*)|*.*"
.ShowOpen
iFile = FreeFile
Open .FileName For Input As iFile
stext = Input(LOF(iFile), iFile)
Close iFile
End With
'If Left(sText, 12) <> "AY Encrypted" Then
'MsgBox "Not an AY Encrypted File Format"
'Else
Frm_LCI_Test_Yourself.Text_JCF_AnswerBox.Text = DecryptText(Mid(stext, 13))
'End If
                                       
Cancelled:
End Function



'saves a text file or encrypted file as an encrypted file

Function save_encrypted2()

On Error GoTo Cancelled
With Frm_LCI_Test_Yourself.CommonDialog1
.CancelError = True
.Filter = "AY Encrypted|*.aye"
.ShowSave
iFile = FreeFile
Open .FileName For Output As iFile
Print #iFile, EncryptText(Frm_LCI_Test_Yourself.Text_JCF_AnswerBox);
Close iFile
End With
MsgBox "Saved"
Cancelled:

End Function


















