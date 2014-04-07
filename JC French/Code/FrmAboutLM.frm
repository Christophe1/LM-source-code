VERSION 5.00
Begin VB.Form FrmAboutLM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Lingua-Master"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "FrmAboutLM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      Picture         =   "FrmAboutLM.frx":08CA
      ScaleHeight     =   795
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Content and proof reading supplied by: Babette Levassor and Therese Boyle."
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label Label7 
      Caption         =   "Design and Artwork: In Your Eye Design"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Programmers: Christophe Harris, Mark Searcy, Naveem Swamy,                              Hilary Moloney"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "Project Manager: Christophe Harris"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Credits:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "www.rosk.ie"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      DrawMode        =   4  'Mask Not Pen
      X1              =   120
      X2              =   5040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Lingua-Master (tm) Copyright © 2003 Rosk Education Systems. The term Lingua-Master is a trademark of Rosk Education-Systems Ltd."
      Height          =   975
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmAboutLM.frx":14FD
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmAboutLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
Unload Me
End Sub




