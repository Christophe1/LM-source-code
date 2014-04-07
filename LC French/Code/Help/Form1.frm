VERSION 5.00
Object = "{22BBD0E3-74FD-11D1-B2C7-00A0C98B5B55}#2.8#0"; "WebPro32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Using Lingua-Master"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":030A
   ScaleHeight     =   6105
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8400
      MouseIcon       =   "Form1.frx":045C
      MousePointer    =   99  'Custom
      ScaleHeight     =   255
      ScaleWidth      =   975
      TabIndex        =   3
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3960
      MouseIcon       =   "Form1.frx":05AE
      MousePointer    =   99  'Custom
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   5640
      Width           =   855
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   10610
      _Version        =   393217
      Style           =   4
      HotTracking     =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WebsterProLib.WebsterPro WebsterPro2 
      Height          =   6015
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   5730
      _Version        =   131080
      _ExtentX        =   10107
      _ExtentY        =   10610
      _StockProps     =   109
      BackColor       =   -2147483647
      PageURL         =   "Webster://Internal/315"
      TitleWindowStyle=   0
      UrlWindowStyle  =   0
      PagesToCache    =   16
      BackColor       =   -2147483647
      BeginProperty FontHeading1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeading2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeading3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeading4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeading5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeading6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontMenu {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDir {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontAddress {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontBlockQuote {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontExample {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontPreformatted {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontListing {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontNormal {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AnchorColor     =   12582912
      ButtonMask      =   2147418112
      VisitedColor    =   12582912
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is the code for bringing up the dictionary box,
'and showing the translation, when a word is clicked


'this is for the back and forward buttons
Private Sub Form_Activate()
'Picture1.ForeColor = vbGreen
Picture1.Print "BACK"

'Picture2.ForeColor = vbGreen
Picture2.Print "FORWARD"
End Sub

Private Sub Form_Load()



Dim ndsnode As Node
' Here your making the first Main Nodes
Set ndsnode = TreeView1.Nodes.Add(, , "", "")
Set ndsnode = TreeView1.Nodes.Add(, , "Heading1", "1.What is LinguaMaster?")
'Set ndsnode = TreeView1.Nodes.Add(, , "Heading2", "2.Introduction to the Aural exam.")
Set ndsnode = TreeView1.Nodes.Add(, , "Heading3", "3.What will LinguaMaster do for me?")
Set ndsnode = TreeView1.Nodes.Add(, , "Heading4", "4.What's on the CD?")
Set ndsnode = TreeView1.Nodes.Add(, , "Heading5", "5.How do I use LinguaMaster?")
'Set ndsnode = TreeView1.Nodes.Add(, , "Heading6", "6.The 'Test Yourself' section.")
Set ndsnode = TreeView1.Nodes.Add(, , "Heading7", "7.Features of LinguaMaster.")


' Here you make child nodes under Main1 Node

'Set ndsnode = TreeView1.Nodes.Add("Heading2", tvwChild, "CHILD1" & A, "2.1")
'Set ndsnode = TreeView1.Nodes.Add("Heading2", tvwChild, "CHILD2" & A, "2.2")
'Set ndsnode = TreeView1.Nodes.Add("Heading2", tvwChild, "CHILD3" & A, "2.3")


'Set ndsnode = TreeView1.Nodes.Add("Heading3", tvwChild, "CHILD4" & A, "3.1")

Set ndsnode = TreeView1.Nodes.Add("Heading5", tvwChild, "CHILD5" & A, "5.1")
Set ndsnode = TreeView1.Nodes.Add("Heading5", tvwChild, "CHILD6" & A, "5.2")

'Set ndsnode = TreeView1.Nodes.Add("Heading6", tvwChild, "CHILD7" & A, "6.1")
'Set ndsnode = TreeView1.Nodes.Add("Heading6", tvwChild, "CHILD8" & A, "6.2")


Set ndsnode = TreeView1.Nodes.Add("Heading7", tvwChild, "CHILD9" & A, "7.1 Play Bar")
Set ndsnode = TreeView1.Nodes.Add("Heading7", tvwChild, "CHILD10" & A, "7.2 Main Menu")
Set ndsnode = TreeView1.Nodes.Add("Heading7", tvwChild, "CHILD11" & A, "7.3 Print Button")
'Set ndsnode = TreeView1.Nodes.Add("Heading7", tvwChild, "CHILD12" & A, "7.4 Show Answers Button")
'Set ndsnode = TreeView1.Nodes.Add("Heading7", tvwChild, "CHILD13" & A, "7.5 Font Size Button")
'Set ndsnode = TreeView1.Nodes.Add("Heading7", tvwChild, "CHILD14" & A, "7.6 Previous,Next,Re-play Buttons")

'Set ndsnode = TreeView1.Nodes.Add("Heading1", tvwChild, "CHILD2" & A, "Part 2")
'Set ndsnode = TreeView1.Nodes.Add("Heading1", tvwChild, "CHILD3" & A, "Part 3")


Text1.Text = Command$

Select Case Text1.Text

Case ""

WebsterPro2.Navigate "file:///" & App.Path & "\First_Page.htm", NavGet, 0, "", "", ""

Case 1
TreeView1.Nodes("Heading5").Expanded = True

WebsterPro2.Navigate "file:///" & App.Path & "\How do I use.htm", NavGet, 0, "", "", ""

Case 2
TreeView1.Nodes("Heading6").Expanded = True

WebsterPro2.Navigate "file:///" & App.Path & "\test_yourself.htm", NavGet, 0, "", "", ""

Case 3
TreeView1.Nodes("Heading2").Expanded = True

WebsterPro2.Navigate "file:///" & App.Path & "\introduction1.htm", NavGet, 0, "", "", ""


End Select

End Sub



'go back
Private Sub Picture1_Click()
WebsterPro2.PageBack
'If WebsterPro2.PageTitle = "First_Page" Then
'WebsterPro2.Navigate "file:///" & App.Path & "\First_Page.htm", NavGet, 0, "", "", ""
'End If

End Sub


'go forwards
Private Sub Picture2_Click()
If Text1.Text = "1" Then
WebsterPro2.Navigate "file:///" & App.Path & "\How do I use2.htm", NavGet, 0, "", "", ""
End If


'WebsterPro2.PageForth
End Sub


Private Sub TreeView1_nodeclick(ByVal Node As MSComctlLib.Node)




If Node.Expanded = True Then
Node.Expanded = False
Else
Node.Expanded = True
End If

    Select Case Node


        Case "1.What is LinguaMaster?"

            WebsterPro2.Navigate "file:///" & App.Path & "\What is.htm", NavGet, 0, "", "", ""


       Case "3.What will LinguaMaster do for me?"

            WebsterPro2.Navigate "file:///" & App.Path & "\What will.htm", NavGet, 0, "", "", ""

'        Case "3.1"
'
'            WebsterPro2.Navigate "file:///" & App.Path & "\What will2.htm", NavGet, 0, "", "", ""


        Case "4.What's on the CD?"

            WebsterPro2.Navigate "file:///" & App.Path & "\What's on.htm", NavGet, 0, "", "", ""



        Case "5.How do I use Lingua-Master?"

            WebsterPro2.Navigate "file:///" & App.Path & "\How do I use.htm", NavGet, 0, "", "", ""

        Case "5.1"

            WebsterPro2.Navigate "file:///" & App.Path & "\How do I use.htm", NavGet, 0, "", "", ""
            Text1.Text = "1"
            
        Case "5.2"

            WebsterPro2.Navigate "file:///" & App.Path & "\How do I use2.htm", NavGet, 0, "", "", ""
            Text1.Text = "2"


        Case "7.Features of Lingua-Master."

            WebsterPro2.Navigate "file:///" & App.Path & "\Features.htm", NavGet, 0, "", "", ""

        Case "7.1 Play Bar"

            WebsterPro2.Navigate "file:///" & App.Path & "\Play Bar.htm", NavGet, 0, "", "", ""

        Case "7.2 Main Menu"

            WebsterPro2.Navigate "file:///" & App.Path & "\Main Menu.htm", NavGet, 0, "", "", ""

        Case "7.3 Print Button"

            WebsterPro2.Navigate "file:///" & App.Path & "\Print Button.htm", NavGet, 0, "", "", ""


        Case "7.4 Previous,Next,Re-play Buttons"

            WebsterPro2.Navigate "file:///" & App.Path & "\Previous,Next,Re-play.htm", NavGet, 0, "", "", ""


    End Select

End Sub



Private Sub Websterpro1_BeforeNavigate(URL As String, ByVal mFlags As Long, ByVal nHandle As WebsterProLib.ObjectHandle, TargetName As String, TextToPost As String, ExtraHeaders As String, Cancel As Boolean)

   ' If flip-flop mode and this is a not a container action (i.e. this is a user action)
   If (mFlags And NavContainerAction) = 0 Then
      ' Have the other control do the navigation
'      Frm_LCI_Dictionary.Show , MDIForm1

'      Form1.WebsterPro1.Navigate URL, mFlags, nHandle, TargetName, TextToPost, ExtraHeaders
      Form1.WebsterPro2.Navigate URL, mFlags, nHandle, TargetName, TextToPost, ExtraHeaders

      ' Cancel our own navigation action
      Cancel = True
    End If

End Sub






