Attribute VB_Name = "Unlock_Code_prices"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As Any, ByVal lpParameters As Any, ByVal lpDirectory As Any, ByVal nShowCmd As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Const MAX_PATH = 260


Public price1 As Single
Public price2 As Single
Public price3 As Single
Public price4 As Single
Public price As Single

'this is the software code. It's value is given
'as a public string because it unloads in
'frm_unlock_code1, and no other forms know
'what it's talking about
Public the_code As String

Public registered_user As String
Public fr_registered_user As String
Public strKey As String

Function set_values()
If Form_More.Check1.Value = Checked Then
price1 = 7.5
Else: price1 = 0
End If
'this line below ensures we get cents, not just stuff like
'.9, but .90
'If ((price1 + price2 + price3 + price4) * 100) Mod 10 = 0 Then
'form_more.Label9.Caption = "€" & price1 + price2 + price3 + price4 & "0"
'Else
Form_More.Label9.Caption = "€" & price1 + price2 + price3 + price4
'End If
'form_more.Command1.Enabled = True
'If price1 + price2 + price3 + price4 = 0 Then
'form_more.Command1.Enabled = False
'End If

If Form_More.Check2.Value = Checked Then
price2 = 7.5
Else: price2 = 0
End If
'this line below ensures we get cents, not just stuff like
'.9, but .90
'If ((price1 + price2 + price3 + price4) * 100) Mod 10 = 0 Then
'form_more.Label9.Caption = "€" & price1 + price2 + price3 + price4 & "0"
'Else
Form_More.Label9.Caption = "€" & price1 + price2 + price3 + price4
'End If
'form_more.Command1.Enabled = True
'If price1 + price2 + price3 + price4 = 0 Then
'form_more.Command1.Enabled = False
'End If

If Form_More.Check3.Value = Checked Then
price3 = 7.5
Else: price3 = 0
End If
'this line below ensures we get cents, not just stuff like
'.9, but .90
'If ((price1 + price2 + price3 + price4) * 100) Mod 10 = 0 Then
'form_more.Label9.Caption = "€" & price1 + price2 + price3 + price4 & "0"
'Else
Form_More.Label9.Caption = "€" & price1 + price2 + price3 + price4
'End If
'form_more.Command1.Enabled = True
'If price1 + price2 + price3 + price4 = 0 Then
'form_more.Command1.Enabled = False
'End If

If Form_More.Check4.Value = Checked Then
price4 = 7.5
Else: price4 = 0
End If
'this line below ensures we get cents, not just stuff like
'.9, but .90
'If ((price1 + price2 + price3 + price4) * 100) Mod 10 = 0 Then
'form_more.Label9.Caption = "€" & price1 + price2 + price3 + price4 & "0"
'Else
Form_More.Label9.Caption = "€" & price1 + price2 + price3 + price4
'End If
'form_more.Command1.Enabled = True
'If price1 + price2 + price3 + price4 = 0 Then
'form_more.Command1.Enabled = False
'End If

price = price1 + price2 + price3 + price4

If price1 + price2 + price3 + price4 = 7.5 Then
Form_More.Label9.Caption = "€7.50"
End If

If price1 + price2 + price3 + price4 = 22.5 Then
Form_More.Label9.Caption = "€22.50"
End If
End Function

