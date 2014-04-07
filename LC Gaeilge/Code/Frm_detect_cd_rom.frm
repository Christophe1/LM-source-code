VERSION 5.00
Begin VB.Form Frm_detect_cd_rom 
   Caption         =   "Frm_detect_cd_rom"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Frm_detect_cd_rom.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Frm_detect_cd_rom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this declaration material is to make sure
'the cd-rom is in the computer

Private Declare Function GetLogicalDriveStrings _
Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal _
nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetDriveType Lib "kernel32" _
Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const DRIVE_CDROM = 5




Private Sub Form_Load()

'we're going to make sure the cd-rom is in the
'drive, first. It looks for a file called
'Licence.txt in all cd drives attached to the
'computer. If and when found, it goes
'to check if the program in ouse is unlocked

    Dim sArr() As String
    Dim i As Integer
    Dim bAllOk As Boolean
    Dim sFileToLookFor As String
    
    sFileToLookFor = "setup.exe"
    sArr = GetCDDrives
    For i = 0 To UBound(sArr) - 1 '-1 because of that last space
        On Error Resume Next
        If Dir(sArr(i) & ":\" & sFileToLookFor) <> "" Then
            If Err Then
                'if an error occur, probably no cd in drive
                bAllOk = False
                Err.Clear
            Else
                'found the file in question
                bAllOk = True
            End If
        Else
            'not found
            bAllOk = False
        End If
        If bAllOk = True Then GoTo line2
    Next i
    
line2:
    If bAllOk Then
    'if cd is in drive, go to line4. sees if Lingua-
    'Master is unlocked or not.
    Unload Me
        Frm_LCI_Learn_and_Revise.Show
    Else
        MsgBox "Please insert the LC LinguaMaster CD-ROM."
        Unload Me
        Exit Sub
    End If
    End Sub

'this function checks to see
'if the cd is in drive
Private Function GetCDDrives() As String()
    Dim tmp As Integer
    Dim tmpStr As String
    Dim Drives As String
    Dim CDsCount As Integer
    Dim CDsLetters As String
    Dim ret As Long
    Dim results As String
    
    'init Drives to 255 spaces
    Drives = Space(255)
    'get drives, Drives var will look like
    'A:\, C:\, D:\, E:\, ..:\
    'ret& is the new length of Drives
    
    ret = GetLogicalDriveStrings(Len(Drives), Drives)
    
    For tmp = 1 To ret& Step 4
        'get a drive root directory (like "C:\")
        tmpStr = Mid(Drives, tmp, 3)
        'if drive is a CD
        
        If GetDriveType(tmpStr) = DRIVE_CDROM Then
            ' CDsCount = CDsCount + 1
            CDsLetters = CDsLetters & Left(tmpStr, 1) & " "
        End If
    Next tmp
    GetCDDrives = Split(CDsLetters, " ")
End Function



