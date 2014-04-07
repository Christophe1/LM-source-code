VERSION 5.00
Object = "{22BBD0E3-74FD-11D1-B2C7-00A0C98B5B55}#2.8#0"; "WebPro32.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_JCI_Learn_and_Revise 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LinguaMaster Junior Certificate Gaeilge Folens Edition 2011"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Frm_JCI_Learn_and_Revise.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Start_Instruction 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   3480
      Picture         =   "Frm_JCI_Learn_and_Revise.frx":08CA
      ScaleHeight     =   6255
      ScaleWidth      =   8415
      TabIndex        =   10
      Top             =   2280
      Width           =   8415
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      Picture         =   "Frm_JCI_Learn_and_Revise.frx":1D3B2
      ScaleHeight     =   1215
      ScaleWidth      =   3135
      TabIndex        =   9
      Top             =   600
      Width           =   3135
   End
   Begin JCIrish.AutoResize Resize 
      Left            =   5520
      Tag             =   "NO"
      Top             =   3840
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin WebsterProLib.WebsterPro jc_ir_WebsterPro 
      Height          =   6255
      Left            =   3480
      TabIndex        =   2
      Top             =   2280
      Width           =   8370
      _Version        =   131080
      _ExtentX        =   14764
      _ExtentY        =   11033
      _StockProps     =   109
      BackColor       =   -2147483647
      PageURL         =   "Webster://Internal/315"
      BevelStyleOuter =   0
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
         Name            =   "Arial"
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
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AnchorColor     =   32768
      ButtonMask      =   2147418112
      ScrollbarStyleHorizontal=   0
      MenuControl     =   0
      ActiveColor     =   33023
      HoverColor      =   33023
      AnchorUnderline =   2
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox textargs 
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash sflash_jc_ir_audio_1 
      Height          =   1575
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   8535
      _cx             =   4209359
      _cy             =   4197082
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ExactFit"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      Picture         =   "Frm_JCI_Learn_and_Revise.frx":213AC
      ScaleHeight     =   375
      ScaleWidth      =   4335
      TabIndex        =   4
      Top             =   0
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash sflash_jci_learn_and_revise 
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
      _cx             =   5953
      _cy             =   11456
      FlashVars       =   ""
      Movie           =   " "
      Src             =   " "
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Menu_Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   11280
      MouseIcon       =   "Frm_JCI_Learn_and_Revise.frx":236A3
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Exit LinguaMaster."
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Menu_Label3 
      BackColor       =   &H8000000E&
      Caption         =   "How to use LinguaMaster"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7080
      MouseIcon       =   "Frm_JCI_Learn_and_Revise.frx":239AD
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "How to benefit most from LinguaMaster."
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Menu_Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4800
      MouseIcon       =   "Frm_JCI_Learn_and_Revise.frx":23CB7
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Print the transcript for the year selected."
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Frm_JCI_Learn_and_Revise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this has to do with the sound files, defining them

Private currentTrack(1) As Variant

'this line is needed so that the sound file loads
'properly
Private formNo As Integer


'set new track#Array() for each track


Private track3Array(6) As String
Private track4Array(6) As String
Private track5Array(6) As String
Private track6Array(6) As String
Private track7Array(6) As String
Private track8Array(6) As String
Private track9Array(6) As String
Private track10Array(6) As String
Private track11Array(6) As String


'Public g_currentTrack As String
Private tempCommand As String
Private tempArgs As String


Private Sub Form_Load()

'If Dir("date.txt") = "" Then
'frm_Date_lock.Show , Frm_JCI_Learn_and_Revise
'End If

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'indicates which swf movies to load

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


'sflash_jc_ir_audio.Playing = True


jc_ir_WebsterPro.BackColor = vbWhite

sflash_jci_learn_and_revise.Playing = True
sflash_jci_learn_and_revise.Movie = App.Path & "\flash\jc_irish_learn_and_revise.swf"


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'this gives times, breaks, track details of each track
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


    
    '2002     SAMPLE A
    track3Array(0) = "3"
    track3Array(1) = "25,149,265,432"
    track3Array(2) = "10,20,152,293,509"
    track3Array(3) = "10,20,152,293,509"
    track3Array(4) = "2002"
    track3Array(5) = "450000"
    
    
    '2003     SAMPLE G
    track4Array(0) = "4"
    track4Array(1) = "27,140,270,468"
    track4Array(2) = "10,20,152,293,509"
    track4Array(3) = "10,20,152,293,509"
    track4Array(4) = "2003"
    track4Array(5) = "480000"
    
    
        
    '2004     SAMPLE F
    track5Array(0) = "5"
    track5Array(1) = "27,165,300,525"
    track5Array(2) = "10,20,152,293,509"
    track5Array(3) = "10,20,152,293,509"
    track5Array(4) = "2004"
    track5Array(5) = "550000"
    
    

    '2005     SAMPLE E
    track6Array(0) = "6"
    track6Array(1) = "30,156,275,475"
    track6Array(2) = "40,40,174,307,544"
    track6Array(3) = "40,40,174,307,544"
    track6Array(4) = "2005"
    track6Array(5) = "500000"


    
    
    '2006     SAMPLE D
    track7Array(0) = "7"
    track7Array(1) = "24,150,277,461"
    track7Array(2) = "27,27,169,310,555"
    track7Array(3) = "27,27,169,310,555"
    track7Array(4) = "2006"
    track7Array(5) = "460000"



    '2007     SAMPLE C
    track8Array(0) = "8"
    track8Array(1) = "10,160,270,395"
    track8Array(2) = "33,33,204,360,643"
    track8Array(3) = "33,33,204,360,643"
    track8Array(4) = "2007"
    track8Array(5) = "440000"
    
    

    '2008     SAMPLE B
    track9Array(0) = "9"
    track9Array(1) = "20,174,321,552"
    track9Array(2) = "33,33,204,360,643"
    track9Array(3) = "33,33,204,360,643"
    track9Array(4) = "2008"
    track9Array(5) = "560000"
    
    
    '2009      SEC SAMPLE
    track10Array(0) = "10"
    track10Array(1) = "35,158,271,475"
    track10Array(2) = "33,33,204,360,643"
    track10Array(3) = "33,33,204,360,643"
    track10Array(4) = "2009"
    track10Array(5) = "480000"
    
    
    '2010     2010
    track11Array(0) = "11"
    track11Array(1) = "30,110,201,482"
    track11Array(2) = "32,32,199,340,600"
    track11Array(3) = "32,32,199,340,600"
    track11Array(4) = "2010"
    track11Array(5) = "350000"
    
    
End Sub







'*******************************************************
'These are the Cuid buttons, A, B and C, on the Playbar.
'*******************************************************

Private Sub sflash_jc_ir_audio_1_FSCommand(ByVal command As String, ByVal args As String)
    
    If command = "loaded" Then
    setupTrack1 tempCommand, tempArgs
        
    End If
    
Select Case textargs.Text





    
    '2002     SAMPLE A

Case 3

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2002text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2002text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2002text.aye#third", NavCreateFromText, 0, "", Text1, ""



End Select


    '2003     SAMPLE G

Case 4

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2003text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2003text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2003text.aye#third", NavCreateFromText, 0, "", Text1, ""



End Select



    '2004     SAMPLE F

Case 5

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2004text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2004text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2004text.aye#third", NavCreateFromText, 0, "", Text1, ""


End Select


 '2005     SAMPLE E
 
Case 6

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2005text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2005text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2005text.aye#third", NavCreateFromText, 0, "", Text1, ""



End Select




'2006   SAMPLE D

Case 7

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2006text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2006text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2006text.aye#third", NavCreateFromText, 0, "", Text1, ""



End Select


'2007        SAMPLE C

Case 8

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2007text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2007text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2007text.aye#third", NavCreateFromText, 0, "", Text1, ""


End Select


'2008       SAMPLE B

Case 9

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2008text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2008text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2008text.aye#third", NavCreateFromText, 0, "", Text1, ""


End Select


'2009   SEC SAMPLE

Case 10

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2009text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2009text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2009text.aye#third", NavCreateFromText, 0, "", Text1, ""



End Select


'2010

Case 11

Select Case args

Case 1

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2009text.aye#first", NavCreateFromText, 0, "", Text1, ""

Case 2

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2009text.aye#second", NavCreateFromText, 0, "", Text1, ""

Case 3

Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\2009text.aye#third", NavCreateFromText, 0, "", Text1, ""



End Select


End Select

End Sub



'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'this passes the command and args variables to the Form_intro form

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub sflash_jci_learn_and_revise_FSCommand(ByVal command As String, ByVal args As String)

        tempCommand = command
    tempArgs = args
    
    Select Case args


    Case 3
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2002.swf"

    Case 4
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2003.swf"
    
    Case 5
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2004.swf"

    Case 6
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2005.swf"
    
    
    Case 7
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2006.swf"

    Case 8
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2007.swf"
    
    Case 9
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2008.swf"
    
        
    Case 10
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2009.swf"
    
            
    Case 11
    sflash_jc_ir_audio_1.Playing = True
    sflash_jc_ir_audio_1.Movie = App.Path & "\Flash\jc_ir_Player_2010.swf"
    
    
    'we use textargs.Text to know which year button is clicked,
    '1993=0, 1994=1, etc....When we know which year is
    'clicked, we can determine which Cuid of which year
    'to jump to.
    End Select
    
    
    Start_Instruction.Visible = False
    
    textargs.Text = args
    
End Sub




'this is the code for bringing up the dictionary box,
'and showing the translation, when a word is clicked


Private Sub jc_ir_Websterpro_BeforeNavigate(URL As String, ByVal mFlags As Long, ByVal nHandle As WebsterProLib.ObjectHandle, TargetName As String, TextToPost As String, ExtraHeaders As String, Cancel As Boolean)

   ' If flip-flop mode and this is a not a container action (i.e. this is a user action)
   If (mFlags And NavContainerAction) = 0 Then
      ' Have the other control do the navigation
      Frm_JCI_Dictionary.Show
      Frm_JCI_Dictionary.WebsterPro_Dictionary.Navigate URL, mFlags, nHandle, TargetName, TextToPost, ExtraHeaders
      
      ' Cancel our own navigation action
      Cancel = True
    End If

End Sub




    'this protect_from_rgt_click module',
    'which protects text from being copied on
    'right click
Private Sub jc_ir_WebsterPro_KeyUp(KeyCode As Integer, Shift As Integer)
Clipboard.Clear

End Sub

'the transcripts are loaded into here. Encrypted.
'as soon as it happens, websterpro grabs it makes it
'visible in html form.
Private Sub Text1_Change()
jc_ir_WebsterPro.Navigate "file:///" & App.Path & "\Aural\Transcripts\", NavCreateFromText, 0, "", Text1, ""
End Sub


'--------------------CODE FOR MENU LABELS-------------
'*****************************************
'*****************************************
'*****************************************
'*****************************************

Private Sub Menu_Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Menu_Label1.ForeColor = &H80FF&

End Sub

'print the transcript
Private Sub Menu_Label1_Click()

On Error GoTo ErrHandler
Printer.Print Frm_JCI_Learn_and_Revise.jc_ir_WebsterPro.DoPrint(True, FromPage, ToPage)

ErrHandler:
Exit Sub

End Sub

Private Sub Menu_Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Menu_Label2.ForeColor = &H80FF&

End Sub

Private Sub Menu_Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Menu_Label3.ForeColor = &H80FF&

End Sub

Private Sub Menu_Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Menu_Label4.ForeColor = &H80FF&

End Sub

'this is the Exit button
Private Sub Menu_Label4_Click()
'Unload Me
    Dim Form As Form
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next Form
End Sub


'this turns the menu_labels back to green,
'when the mouse has moved off the labels.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Menu_Label1.ForeColor = &HC000&
'Menu_Label2.ForeColor = &HC000&
Menu_Label3.ForeColor = &HC000&
Menu_Label4.ForeColor = &HC000&
End Sub

Private Sub jc_ir_WebsterPro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Menu_Label1.ForeColor = &HC000&
'Menu_Label2.ForeColor = &HC000&
Menu_Label3.ForeColor = &HC000&
Menu_Label4.ForeColor = &HC000&
End Sub


'this opens the Help file, on the 'How to Use.htm' page
'the number '1' loads into the text box of the help.exe,
'telling help.exe which file to load.
Private Sub Menu_Label3_Click()

Call ShellExecute(hWnd, "Open", "Help\Help.exe", "1", App.Path, 1)

End Sub

'Private Sub Menu_Label2_Click()
'Form_More.Show , Frm_JCI_Learn_and_Revise
'End Sub











'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'Function that sets up the player.swf to play the right track
'in the correct Form.

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function setupSwf(sNo As Integer, sArray As String, sTrackNo As Integer, sTrackID As String, sTrackLength As String)
    Select Case formNo
        Case 1
            Frm_JCI_Learn_and_Revise.sflash_jc_ir_audio_1.Visible = True
            Frm_JCI_Learn_and_Revise.sflash_jc_ir_audio_1.SetVariable "control.setupNo", sNo
            Frm_JCI_Learn_and_Revise.sflash_jc_ir_audio_1.SetVariable "control.setupSegArray", sArray
            Frm_JCI_Learn_and_Revise.sflash_jc_ir_audio_1.SetVariable "control.trackNo", sTrackNo
            Frm_JCI_Learn_and_Revise.sflash_jc_ir_audio_1.SetVariable "control.trackID", sTrackID
            Frm_JCI_Learn_and_Revise.sflash_jc_ir_audio_1.SetVariable "control.trackLength", sTrackLength
            Frm_JCI_Learn_and_Revise.sflash_jc_ir_audio_1.SetVariable "control.flag", "false"
        Case 2
           Frm_JCI_Test_Yourself.sflash_jc_ir_audio_2.Visible = True
            Frm_JCI_Test_Yourself.sflash_jc_ir_audio_2.SetVariable "control.setupNo", sNo
            Frm_JCI_Test_Yourself.sflash_jc_ir_audio_2.SetVariable "control.setupSegArray", sArray
            Frm_JCI_Test_Yourself.sflash_jc_ir_audio_2.SetVariable "control.trackNo", sTrackNo
            Frm_JCI_Test_Yourself.sflash_jc_ir_audio_2.SetVariable "control.trackID", sTrackID
            Frm_JCI_Test_Yourself.sflash_jc_ir_audio_2.SetVariable "control.trackLength", sTrackLength
            Frm_JCI_Test_Yourself.sflash_jc_ir_audio_2.SetVariable "control.flag", "false"
    End Select
    
End Function



'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'Setup Function for Form_Learn_and_Revise

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function setupTrack1(command As String, args As String)

    formNo = 1

    Select Case args
    'add case for each button in the Learn_and_Revise.swf
    


'start of 2002

Case 3
currentTrack(1) = track3Array()
Call open_transcript(2002)


'the year below is placed in the Text2 box
'and corresponding dictionary_print file
'is opened in the Dictionary_frm, when opened
Text2.Text = "2002"

'start of 2003

Case 4
currentTrack(1) = track4Array()
Call open_transcript(2003)

'the year below is placed in the Text2 box
'and corresponding dictionary_print file
'is opened in the Dictionary_frm, when opened
Text2.Text = "2003"


'start of 2004

Case 5
currentTrack(1) = track5Array()
Call open_transcript(2004)


'the year below is placed in the Text2 box
'and corresponding dictionary_print file
'is opened in the Dictionary_frm, when opened
Text2.Text = "2004"



'start of 2005

Case 6
currentTrack(1) = track6Array()
Call open_transcript(2005)


'the year below is placed in the Text2 box
'and corresponding dictionary_print file
'is opened in the Dictionary_frm, when opened
Text2.Text = "2005"


'start of 2006

Case 7
currentTrack(1) = track7Array()
Call open_transcript(2006)


'the year below is placed in the Text2 box
'and corresponding dictionary_print file
'is opened in the Dictionary_frm, when opened
Text2.Text = "2006"



''start of 2007

Case 8
currentTrack(1) = track8Array()
Call open_transcript(2007)

Text2.Text = "2007"

''start of 2008

Case 9
currentTrack(1) = track9Array()
Call open_transcript(2008)

Text2.Text = "2008"


''start of 2009

Case 10
currentTrack(1) = track10Array()
Call open_transcript(2009)

'the year below is placed in the Text2 box
'and corresponding dictionary_print file
'is opened in the Dictionary_frm, when opened
Text2.Text = "2009"

''start of 2010

Case 11
currentTrack(1) = track11Array()
Call open_transcript(2010)

'the year below is placed in the Text2 box
'and corresponding dictionary_print file
'is opened in the Dictionary_frm, when opened
Text2.Text = "2010"


        
        
End Select

    setupSwf 0, CStr(currentTrack(1)(1)), CStr(currentTrack(1)(0)), CStr(currentTrack(1)(4)), CStr(currentTrack(1)(5))
    
End Function




