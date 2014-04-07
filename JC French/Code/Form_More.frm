VERSION 5.00
Object = "{147D0F6E-DDA3-44B2-A616-1A85E16C08DA}#1.0#0"; "Lingua.ocx"
Begin VB.Form Form_More 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products list"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "Form_More.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   240
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton Command4 
         Caption         =   "Trigonometry, Higher Level"
         Height          =   255
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Algebra, Higher Level"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Paper 2, Ordinary Level"
         Height          =   255
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Paper 1, Ordinary Level"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1st 
         Caption         =   $"Form_More.frx":08CA
         Height          =   615
         Left            =   360
         TabIndex        =   41
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label2nd 
         Caption         =   "2. Click a button to install the product of your choice."
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label3rd 
         Caption         =   "3. When you start the program for the first time, enter the             Unlock code and this will activate the program."
         Height          =   495
         Left            =   360
         TabIndex        =   39
         Top             =   2640
         Width           =   4335
      End
   End
   Begin VB.CommandButton BACK2 
      Caption         =   "<BACK"
      Height          =   375
      Left            =   720
      TabIndex        =   33
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton NEXT2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   32
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Back1 
      Caption         =   "<BACK"
      Height          =   375
      Left            =   720
      TabIndex        =   31
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Next1 
      Caption         =   "NEXT>"
      Height          =   375
      Left            =   2880
      TabIndex        =   30
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CheckBox Check5 
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   3240
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   600
         ScaleHeight     =   1215
         ScaleWidth      =   2415
         TabIndex        =   10
         Top             =   960
         Width           =   2415
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            MouseIcon       =   "Form_More.frx":0966
            MousePointer    =   99  'Custom
            Picture         =   "Form_More.frx":0C70
            ScaleHeight     =   255
            ScaleWidth      =   2415
            TabIndex        =   14
            Top             =   960
            Width           =   2415
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            MouseIcon       =   "Form_More.frx":1A9A
            MousePointer    =   99  'Custom
            Picture         =   "Form_More.frx":1DA4
            ScaleHeight     =   255
            ScaleWidth      =   2295
            TabIndex        =   13
            Top             =   640
            Width           =   2295
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            MouseIcon       =   "Form_More.frx":298E
            MousePointer    =   99  'Custom
            Picture         =   "Form_More.frx":2C98
            ScaleHeight     =   255
            ScaleWidth      =   2295
            TabIndex        =   12
            Top             =   325
            Width           =   2295
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            MouseIcon       =   "Form_More.frx":392E
            MousePointer    =   99  'Custom
            Picture         =   "Form_More.frx":3C38
            ScaleHeight     =   255
            ScaleWidth      =   2295
            TabIndex        =   11
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   1935
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   1635
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   1335
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "Terms and Conditions."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         MouseIcon       =   "Form_More.frx":497B
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   3285
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "I agree with the "
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   3285
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "€7.50"
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   1965
         Width           =   495
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   2640
         Width           =   495
      End
      Begin VB.Line Line6 
         X1              =   600
         X2              =   3960
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label8 
         Caption         =   "Total Cost:"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "€7.50"
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   1665
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "€7.50"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   1365
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "€7.50"
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "What Junior Cert. product(s) do you want to buy?"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton Buy_Maths 
      Caption         =   "Buy Maths-Master"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton More_info 
      Caption         =   "More about Maths-Master"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form_More.frx":4C85
      Top             =   2160
      Width           =   5295
   End
   Begin Lingua.ActiveLock AL_JC_H_ALGEBRA 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   820
      SoftwareName    =   "JC_H_ALGEBRA"
      SoftwarePassword=   "JC_H_ALGEBRA"
      LiberationKeyLength=   6
      SoftwareCodeLength=   6
      LockToHardDrive =   -1  'True
      LockToWindowsSerial=   -1  'True
      LockToRandomNumber=   0   'False
      LockToComputerName=   0   'False
      LockToMACAddress=   0   'False
      UseDataLock     =   0   'False
      RegistryPath    =   "Purchases"
      RegistryKey     =   "Rustic Services"
      RegistryName    =   "Transactions"
      RegistryHive    =   "HKLM"
      LockToCustomString=   ""
      HashAlgorithm   =   0
      RegCounterKey   =   "Counter"
      RegLiberationKey=   "Credits"
      RegLastRunDateKey=   "LastRunDate"
      RegInitialRunDateKey=   "InitialRunDate"
      RegRandomKey    =   "RandomKey"
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      Picture         =   "Form_More.frx":4F3C
      ScaleHeight     =   1215
      ScaleWidth      =   2295
      TabIndex        =   42
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label11 
      Caption         =   "User Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label_about 
      Caption         =   "About Maths-Master"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label_Copy 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Maths-Master (tm) Copyright © 2003 Rosk Education Systems. The term Maths-Master is a trademark of Rosk Education Systems Ltd."
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      DrawMode        =   4  'Mask Not Pen
      X1              =   120
      X2              =   5400
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "Form_More"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

'this declaration material is to make sure
'the cd-rom is in the computer

Private Declare Function GetLogicalDriveStrings _
Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal _
nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetDriveType Lib "kernel32" _
Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const DRIVE_CDROM = 5




Private Sub Buy_Maths_Click()

Label_about.Visible = False
Label_Copy.Visible = False
Picture3.Visible = False
Picture3.Visible = False
Text1.Visible = False

Buy_Maths.Visible = False
More_info.Visible = False


Frame1.Visible = True
Label11.Visible = True
Label12.Visible = True

Label12.Caption = AL_JC_H_ALGEBRA.SoftwareCode

Next1.Visible = True
Back1.Visible = True

End Sub





Private Sub More_info_Click()
Call ShellExecute(hWnd, "Open", "http://rosk.ie/maths-master/maths-master.htm", "", App.Path, 1)
End Sub



Private Sub Back1_Click()

Label_about.Visible = True
Label_Copy.Visible = True
Picture3.Visible = True
'Label18.Visible = True
Text1.Visible = True

Buy_Maths.Visible = True
More_info.Visible = True


Frame1.Visible = False
Label11.Visible = False
Label12.Visible = False

Next1.Visible = False
Back1.Visible = False

End Sub


Private Sub NEXT1_Click()

'If Check5.Value = Unchecked Then
'MsgBox "Please confirm you agree with Terms and Conditions by clicking the check box."
'Else
'Me.Hide
'Frm_Unlock_Code3.Show
'End If

If Check1.Value = Unchecked And Check2.Value = Unchecked _
And Check3.Value = Unchecked And Check4.Value = Unchecked Then
MsgBox "Please tick which product(s) you want to buy."
End If

If Check5.Value = Unchecked Then
MsgBox "Please confirm you agree with Terms and Conditions by clicking the check box."
End If

If Check1.Value = Checked And Check5.Value = Checked _
Or Check2.Value = Checked And Check5.Value = Checked _
Or Check3.Value = Checked And Check5.Value = Checked _
Or Check4.Value = Checked And Check5.Value = Checked _
Then

Next1.Visible = False
Back1.Visible = False
Frame1.Visible = False

Frame2.Visible = True
NEXT2.Visible = True
BACK2.Visible = True

End If

End Sub


Private Sub NEXT2_Click()
Unload Form_More
End Sub

Private Sub BACK2_Click()

Next1.Visible = True
Back1.Visible = True
Frame1.Visible = True

Frame2.Visible = False
NEXT2.Visible = False
BACK2.Visible = False
End Sub


''''''''''''''''''''''
''''''''''''''''''''''
''''''''''''''''''''''
''''''''''''''''''''''






Private Sub Check1_Click()
Call set_values

End Sub


Private Sub Check2_Click()
Call set_values

End Sub


Private Sub Check3_Click()
Call set_values

End Sub

Private Sub Check4_Click()
Call set_values

End Sub



'Private Sub Command2_Click()
'Me.Hide
'Frm_Unlock_Code1.Show
'End Sub

'Private Sub Form_Load()
'Label12.Caption = AL_JC_H_ALGEBRA.SoftwareCode
'If registered_user = "True" Then
'Me.Show vbModeless, Form1
'Else
'Me.Show
'End If
'
'
'End Sub



Private Sub Label14_Click()
Frm_Licence.Show
End Sub



Private Sub Picture1_Click()
Select Case Check1.Value
Case Checked
Check1.Value = Unchecked
Case Unchecked
Check1.Value = Checked
End Select
Call set_values
End Sub

Private Sub Picture4_Click()
Select Case Check2.Value
Case Checked
Check2.Value = Unchecked
Case Unchecked
Check2.Value = Checked
End Select
Call set_values
End Sub

Private Sub Picture5_Click()
Select Case Check3.Value
Case Checked
Check3.Value = Unchecked
Case Unchecked
Check3.Value = Checked
End Select
Call set_values
End Sub

Private Sub Picture6_Click()
Select Case Check4.Value
Case Checked
Check4.Value = Unchecked
Case Unchecked
Check4.Value = Checked
End Select
Call set_values
End Sub




'''''''''''''''''
'''''''''''''''''
'''''''''''''''''
'''''''''''''''''
'Install product of choice buttons
'''''''''''''''''
'''''''''''''''''
'''''''''''''''''
'''''''''''''''''


Private Sub Command1_Click()
'The use has chosen to install lc irish

'we're going to make sure the cd-rom is in the
'drive, first. It looks for a file called
'Lingua-Master lc irish.exe in all cd drives attached to the
'computer. If and when found, it starts the install process.

'if not found, show a message box asking to insert the
'cd-rom

    Dim sArr() As String
    Dim i As Integer
    Dim bAllOk As Boolean
    Dim sFileToLookFor As String
'    Dim cd_drive As String
    
    sFileToLookFor = "Maths-Master\Paper 1\M-M JC Ordinary P1.exe"
    sArr = GetCDDrives
    For i = 0 To UBound(sArr) - 1 '-1 because of that last space
        On Error Resume Next
        If Dir(sArr(i) & ":\" & sFileToLookFor) <> "" Then
'        cd_drive = Dir(sArr(i))
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
    'if cd is in drive we start to install Lc Irish

        Shell (sArr(i) & ":\" & sFileToLookFor)
            Dim Form As Form
                    For Each Form In Forms
                       Unload Form
                       Set Form = Nothing
                    Next Form
    Else
        MsgBox "Please insert the correct CD-ROM.", vbOKOnly
    End If
    End Sub


Private Sub Command2_Click()
'The user has chosen to install lc french

'we're going to make sure the cd-rom is in the
'drive, first. It looks for a file called
'Lingua-Master lc french.exe in all cd drives attached to the
'computer. If and when found, it starts the install process.

'if not found, show a message box asking to insert the
'cd-rom

    Dim sArr() As String
    Dim i As Integer
    Dim bAllOk As Boolean
    Dim sFileToLookFor As String
'    Dim cd_drive As String
    
    sFileToLookFor = "Maths-Master\Paper 2\M-M JC Ordinary P2.exe"
    sArr = GetCDDrives
    For i = 0 To UBound(sArr) - 1 '-1 because of that last space
        On Error Resume Next
        If Dir(sArr(i) & ":\" & sFileToLookFor) <> "" Then
'        cd_drive = Dir(sArr(i))
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
    'if cd is in drive we start to install Lc Irish

        Shell (sArr(i) & ":\" & sFileToLookFor)
            Dim Form As Form
                    For Each Form In Forms
                       Unload Form
                       Set Form = Nothing
                    Next Form
    Else
        MsgBox "Please insert the correct CD-ROM.", vbOKOnly
    End If
End Sub



Private Sub Command3_Click()
'The user has chosen to install lc german

'we're going to make sure the cd-rom is in the
'drive, first. It looks for a file called
'Lingua-Master lc german.exe in all cd drives attached to the
'computer. If and when found, it starts the install process.

'if not found, show a message box asking to insert the
'cd-rom

    Dim sArr() As String
    Dim i As Integer
    Dim bAllOk As Boolean
    Dim sFileToLookFor As String
'    Dim cd_drive As String
    
    sFileToLookFor = "Maths-Master\JC Algebra\M-M JC Higher Algebra.exe"
    sArr = GetCDDrives
    For i = 0 To UBound(sArr) - 1 '-1 because of that last space
        On Error Resume Next
        If Dir(sArr(i) & ":\" & sFileToLookFor) <> "" Then
'        cd_drive = Dir(sArr(i))
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
    'if cd is in drive we start to install Lc Irish

        Shell (sArr(i) & ":\" & sFileToLookFor)
            Dim Form As Form
                    For Each Form In Forms
                       Unload Form
                       Set Form = Nothing
                    Next Form
    Else
        MsgBox "Please insert the correct CD-ROM.", vbOKOnly
    End If
End Sub



Private Sub Command4_Click()
'The user has chosen to install lc german

'we're going to make sure the cd-rom is in the
'drive, first. It looks for a file called
'Lingua-Master lc german.exe in all cd drives attached to the
'computer. If and when found, it starts the install process.

'if not found, show a message box asking to insert the
'cd-rom

    Dim sArr() As String
    Dim i As Integer
    Dim bAllOk As Boolean
    Dim sFileToLookFor As String
'    Dim cd_drive As String
    
    sFileToLookFor = "Maths-Master\JC Trig\M-M JC Higher Trigonometry.exe"
    sArr = GetCDDrives
    For i = 0 To UBound(sArr) - 1 '-1 because of that last space
        On Error Resume Next
        If Dir(sArr(i) & ":\" & sFileToLookFor) <> "" Then
'        cd_drive = Dir(sArr(i))
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
    'if cd is in drive we start to install Lc Irish

        Shell (sArr(i) & ":\" & sFileToLookFor)
            Dim Form As Form
                    For Each Form In Forms
                       Unload Form
                       Set Form = Nothing
                    Next Form
    Else
        MsgBox "Please insert the correct CD-ROM.", vbOKOnly
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






