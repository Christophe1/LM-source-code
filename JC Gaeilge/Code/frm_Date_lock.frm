VERSION 5.00
Begin VB.Form frm_Date_lock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lingua-Master"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5475
   Icon            =   "frm_Date_lock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "This fully functional version of Lingua-Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   " (Folens Edition) will expire on"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "The Listening Comprehension recordings will continue to work on an Audio CD Player."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
   End
End
Attribute VB_Name = "frm_Date_lock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strState As String

Private Sub Form_Load()

strState = GetSetting("myapp", "JCIrish", "FRD", "1/1/2000")
Label4.Caption = DateAdd("yyyy", 1, Date)

If strState = "1/1/2000" Then
Call DateGood(365)
'Unload Me
End If

If strState <> "1/1/2000" Then
Frm_JCI_Learn_and_Revise.Show
Call DateGood(365)
Unload Me
End If

If Not DateGood(365) Then
MsgBox "The Trial Period has Expired!", vbExclamation, "Lingua-Master"
    Dim Form As Form
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next Form
End If

End Sub

Private Sub Command1_Click()
Frm_JCI_Learn_and_Revise.Show
Unload Me
End Sub















'Private Sub createBlankFile(ByVal strFileName As String)
'    Open strFileName For Output As #1: Close #1
'End Sub
'
'Private Sub Form_Load()
'date1 = DateAdd("yyyy", 1, Date)
'Label3.Caption = date1 & "."
'End Sub
'
'Private Sub Command1_Click()
'createBlankFile "date.txt"
'Unload Me
'End Sub
