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



