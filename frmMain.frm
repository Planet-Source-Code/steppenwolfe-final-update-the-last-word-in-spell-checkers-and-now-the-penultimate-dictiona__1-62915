VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spell Check"
   ClientHeight    =   11325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "Translate"
      Height          =   345
      Left            =   5700
      TabIndex        =   64
      Top             =   4560
      Width           =   945
   End
   Begin VB.CommandButton cmdDictionary 
      Caption         =   "Dictionary"
      Height          =   345
      Left            =   4620
      TabIndex        =   63
      Top             =   4560
      Width           =   915
   End
   Begin VB.Frame fmDictionary 
      Caption         =   "Definition Dictionary Options"
      Height          =   1425
      Index           =   1
      Left            =   180
      TabIndex        =   50
      Top             =   9660
      Width           =   7545
      Begin VB.CheckBox chkTooltip 
         Caption         =   "Tooltip Definitions"
         Height          =   195
         Left            =   5250
         TabIndex        =   61
         Top             =   990
         Width           =   1665
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "Format Results"
         Height          =   225
         Left            =   5250
         TabIndex        =   60
         Top             =   660
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   2820
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   57
         Top             =   600
         Width           =   2055
         Begin VB.OptionButton optProgress 
            Caption         =   "No Progress (Faster)"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   59
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optProgress 
            Caption         =   "Load with Progress Bar"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   58
            Top             =   60
            Width           =   2025
         End
      End
      Begin VB.CheckBox chkSearch 
         Caption         =   "Adjectives"
         Height          =   225
         Index           =   2
         Left            =   1200
         TabIndex        =   55
         Top             =   660
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkSearch 
         Caption         =   "Iterative Return"
         Height          =   225
         Index           =   3
         Left            =   1200
         TabIndex        =   54
         Top             =   990
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkSearch 
         Caption         =   "Nouns"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   52
         Top             =   660
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox chkSearch 
         Caption         =   "Verbs"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   51
         Top             =   990
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.Label lblDictionary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advanced Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5250
         TabIndex        =   62
         Top             =   330
         Width           =   1320
      End
      Begin VB.Label lblDictionary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Up Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   56
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label lblDictionary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return Results:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   53
         Top             =   330
         Width           =   1125
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Correction Options"
      Height          =   1065
      Index           =   4
      Left            =   180
      TabIndex        =   38
      Top             =   7620
      Width           =   7545
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   1
         Left            =   90
         ScaleHeight     =   855
         ScaleWidth      =   7365
         TabIndex        =   39
         Top             =   180
         Width           =   7365
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   765
            Left            =   1410
            ScaleHeight     =   765
            ScaleWidth      =   5895
            TabIndex        =   41
            Top             =   30
            Width           =   5895
            Begin VB.OptionButton optCorrectstyle 
               Caption         =   "Line"
               Height          =   225
               Index           =   3
               Left            =   4530
               TabIndex        =   47
               Top             =   270
               Width           =   1185
            End
            Begin VB.ComboBox cbType 
               Height          =   315
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   46
               Text            =   "Combo1"
               Top             =   240
               Width           =   915
            End
            Begin VB.OptionButton optCorrectstyle 
               Caption         =   "Sentence"
               Height          =   225
               Index           =   2
               Left            =   3390
               TabIndex        =   44
               Top             =   270
               Width           =   1005
            End
            Begin VB.OptionButton optCorrectstyle 
               Caption         =   "Word"
               Height          =   225
               Index           =   1
               Left            =   2460
               TabIndex        =   43
               Top             =   270
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton optCorrectstyle 
               Caption         =   "Interval:"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   42
               Top             =   270
               Width           =   1425
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "Non Timer Methods:"
               Height          =   195
               Index           =   7
               Left            =   2490
               TabIndex        =   48
               Top             =   30
               Width           =   1440
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "Time: Min."
               Height          =   195
               Index           =   6
               Left            =   1080
               TabIndex        =   45
               Top             =   60
               Width           =   735
            End
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Auto Correct"
            Height          =   225
            Index           =   3
            Left            =   150
            TabIndex        =   40
            Top             =   300
            Width           =   1245
         End
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Dictionary Selection"
      Height          =   765
      Index           =   3
      Left            =   180
      TabIndex        =   33
      Top             =   8790
      Width           =   7545
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   90
         ScaleHeight     =   375
         ScaleWidth      =   7365
         TabIndex        =   34
         Top             =   270
         Width           =   7365
         Begin VB.OptionButton optDictionary 
            Caption         =   "Full List (2430k)"
            Height          =   225
            Index           =   2
            Left            =   4830
            TabIndex        =   37
            Top             =   90
            Width           =   1455
         End
         Begin VB.OptionButton optDictionary 
            Caption         =   "Standard (480k)"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   36
            Top             =   90
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optDictionary 
            Caption         =   "Medium (recommended) (1200k)"
            Height          =   225
            Index           =   1
            Left            =   1890
            TabIndex        =   35
            Top             =   90
            Width           =   2775
         End
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "General Options"
      Height          =   1155
      Index           =   2
      Left            =   210
      TabIndex        =   16
      Top             =   6360
      Width           =   7515
      Begin VB.CheckBox chkOptions 
         Caption         =   "Common Word Extensions"
         Height          =   225
         Index           =   6
         Left            =   2700
         TabIndex        =   49
         Top             =   690
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Highlight Corrections"
         Height          =   225
         Index           =   5
         Left            =   5250
         TabIndex        =   21
         Top             =   690
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Highlight Errors"
         Height          =   225
         Index           =   4
         Left            =   5250
         TabIndex        =   20
         Top             =   360
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Enable Menu Suggestions"
         Height          =   225
         Index           =   2
         Left            =   2700
         TabIndex        =   19
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Enable Menu Replacement"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   690
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Sort Results by Accuracy %"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Value           =   1  'Checked
         Width           =   2325
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Suggestion Options"
      Height          =   1155
      Index           =   1
      Left            =   2220
      TabIndex        =   11
      Top             =   5130
      Width           =   5505
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   3990
         ScaleHeight     =   765
         ScaleWidth      =   795
         TabIndex        =   26
         Top             =   240
         Width           =   795
         Begin VB.OptionButton optTolerance 
            Caption         =   "Low"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   28
            Top             =   270
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.OptionButton optTolerance 
            Caption         =   "High"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   27
            Top             =   570
            Width           =   645
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Tolerance"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Index           =   0
         Left            =   2160
         ScaleHeight     =   825
         ScaleWidth      =   1335
         TabIndex        =   22
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton optSort 
            Caption         =   "Ascending"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   24
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Descending"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   570
            Width           =   1185
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Relevance Order"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Precision Filter"
         Height          =   225
         Left            =   330
         TabIndex        =   15
         Top             =   780
         Width           =   1335
      End
      Begin VB.ComboBox cbFreq 
         Height          =   315
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "70"
         Top             =   450
         Width           =   885
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Accuracy %"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   13
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Database Options"
      Height          =   1155
      Index           =   0
      Left            =   210
      TabIndex        =   10
      Top             =   5130
      Width           =   1815
      Begin VB.PictureBox picComp 
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   150
         ScaleHeight     =   705
         ScaleWidth      =   1605
         TabIndex        =   30
         Top             =   270
         Width           =   1605
         Begin VB.OptionButton optCompress 
            Caption         =   "High Compression"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   32
            Top             =   390
            Width           =   1635
         End
         Begin VB.OptionButton optCompress 
            Caption         =   "Fast Compression"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   31
            Top             =   90
            Value           =   -1  'True
            Width           =   1575
         End
      End
   End
   Begin VB.ListBox lstSuggest 
      Height          =   1620
      Index           =   1
      ItemData        =   "frmMain.frx":0000
      Left            =   4050
      List            =   "frmMain.frx":0002
      TabIndex        =   7
      Top             =   2790
      Width           =   3645
   End
   Begin VB.CommandButton cmdControls 
      Caption         =   "Replace All"
      Height          =   345
      Index           =   1
      Left            =   1350
      TabIndex        =   6
      Top             =   4560
      Width           =   945
   End
   Begin VB.CommandButton cmdControls 
      Caption         =   "Skip Word"
      Height          =   345
      Index           =   2
      Left            =   2430
      TabIndex        =   5
      Top             =   4560
      Width           =   945
   End
   Begin VB.CommandButton cmdControls 
      Caption         =   "Exception"
      Height          =   345
      Index           =   3
      Left            =   3540
      TabIndex        =   4
      Top             =   4560
      Width           =   915
   End
   Begin VB.CommandButton cmdControls 
      Caption         =   "Replace"
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   945
   End
   Begin VB.ListBox lstSuggest 
      Height          =   1620
      Index           =   0
      ItemData        =   "frmMain.frx":0004
      Left            =   240
      List            =   "frmMain.frx":0006
      TabIndex        =   2
      Top             =   2790
      Width           =   3645
   End
   Begin VB.CommandButton cmdControls 
      Caption         =   "Spelling"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   6780
      TabIndex        =   1
      Top             =   4560
      Width           =   945
   End
   Begin RichTextLib.RichTextBox txtBody 
      Height          =   2115
      Left            =   210
      TabIndex        =   0
      Top             =   330
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   3731
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0008
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Text Body"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Suggestions"
      Height          =   195
      Index           =   3
      Left            =   4050
      TabIndex        =   9
      Top             =   2580
      Width           =   870
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Not Found"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2580
      Width           =   750
   End
   Begin VB.Menu mnuControl1 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuControl 
         Caption         =   "Cut"
         Index           =   0
      End
      Begin VB.Menu mnuControl 
         Caption         =   "Copy"
         Index           =   1
      End
      Begin VB.Menu mnuControl 
         Caption         =   "Paste"
         Index           =   2
      End
      Begin VB.Menu mnuControl 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuControl 
         Caption         =   "Replace Selected"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuControl 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuControl 
         Caption         =   "- Suggestions -"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuControl 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuControl 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuControl 
         Caption         =   ""
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuControl 
         Caption         =   ""
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuControl 
         Caption         =   ""
         Index           =   11
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'** Spell Checker - John Underhill Oct 16 2005
'** original post: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62915&lngWId=1&txtForceRefresh=1016200523424785279

'** For a comment (or a job.. ;o) email: steppenwolfe_2000@yahoo.com
'** You can use this code in any way you like, provided you keep this comment field intact.

'** This tool is for educational purposes, the use of word lists and dictionaries and any respective
'** copyright responsibilities are up to the user to investigate in their country.
'** Wordlists originated from crossword puzzle and scrabble sites, and are a compilation of several common word lists
'** no copyright notice or instructions were posted at the source of origin,
'** a list of which, can be found here: http://www.net-comber.com/wordurls.html, so I am assuming they are fair game.
'** the dictionary is also a complilation. The main body of which was taken from a
'** website that claimed it is a retired document and public domain.

'** Note** The first time this loads it takes some time, because all three dictionaries are
'** being compressed on the first run. The First_Run sub is in mMain, and can be commented
'** out after the initial run.
'
'** This program has several unique features in it, including some nice tolerance features
'** using api to build match percentage profiles and sorting the results by relevance, api compression,
'** rtb menu tricks, and some fast sorting and comparison routines..
'** Thanks go out to..
'** Philippe Lord for the incredible array sorting and hash table routines
'** Ion Alex Ionescu for his demonstration of API compression, very impressive guy that Alex..
'** Cyber Chris for his spell checker, of all the examples on PSC, his was the best,
'** and the seed for this project (although I wrote all this from scratch, as you should do with my example.. ;o)
'** Anyways, if you like it vote, if you can think of improvements, or spot mistakes, post a comment, after all
'** this is just the beta, (about 12 hrs of work = 5am this morning.. ,:~O}

'** Oct 17 2005
'** Added some automation to it this morning, now several different auto check features.
'** also fixed the filter function and went bug hunting..

'** Oct 18 2005
'** For those of you still using 98/ME, (you have my sympathies..), added the Huffman class
'** as an alternate to the api compression. Class courtesy of Fredrik Qvarfort, and left unaltered.

'** Oct 19 2005
'** Added some filter changes, added right click auto selects word at cursor pos and filters punctuation
'*** Working on a definition dictionary, 270,000 words with full definitions, should be done in a day or so

'** Oct 21 2005
'** Added the 'Penultimate Dictionary - 221,000 word definitions!!!
'** optional tooltip definitions, lightning fast dictionary class, multiple results filter et al..
'** Thanks go out to Mario Flores Gonzalez for his incredible ToolTip class, best I've seen..

'** Added a switch that turns off the extensions search when the larger dictionaries are loaded.
'** Added a word in word fix for highlight corrections routine.
'** Rebuilt highlight corrections without seltext to greatly improve performance on large documents.

'** Added a basic set of translation dictionaries, not as many words as I'd like, but as good as
'** I could find so far.. translates english words to french/italian/spanish, reverse dictionary in the works..

'** Included sample project with basic spell check as a class module to demonstrate how easily this
'** project could be exported to active-x dll, or inline class, also object independance means
'** could be used with any object that uses strings with only a few modifications (and a little imagination..).

'** I also included an example in frmTranslate of how any file can be reformatted and used with
'** the mTranslate collection routines. So, for those that want a thesaurus, or a dictionary in another language
'** just find the file you need and format it as described, and it can be reused by the mTranslate collection routines..

'** Have fun! This was a lot of work, so if you like it, plan to use it, or think it is good code..
'** Then show your appreciation and throw a vote my way..
'** Cheers :o)
'** John


Option Explicit

'//dictionary with events
Public Event Process(percent As Long)
Public WithEvents cDictionary    As clsDictionary
Attribute cDictionary.VB_VarHelpID = -1

'//dictionary dimns
Private lGCount                  As Long
Private iPerc                    As Integer

'//tooltip with events
Public WithEvents cTip           As ExToolTip
Attribute cTip.VB_VarHelpID = -1
Public Event Status(bStatus As Boolean)
Private bLocal                   As Boolean

'//declares
Private bPosition                As Boolean
Private mMove                    As Boolean

'//mouse to word stuff
Private Const CHAR_POS           As Long = &HD7

Private Type tPtr
    cx                           As Long
    cy                           As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByRef lParam As Any) As Long

'//spell check declares
Private lTMem       As Long
Private lSLen       As Long
Private lSSen       As Long
Private lWLen       As Long
Public lWait        As Long
Public bBack        As Boolean

Private Sub chkOptions_Click(Index As Integer)
'//control array calls
    Select Case Index
        Case 0
            If chkOptions(0).Value Then
                optSort(0).Enabled = True
                optSort(1).Enabled = True
                chkFilter.Enabled = True
            Else
                optSort(0).Enabled = False
                optSort(1).Enabled = False
                chkFilter.Value = 0
                chkFilter.Enabled = False
            End If

        Case 1
                If chkOptions(1).Value = 1 Then
                    mnuControl(4).Enabled = True
                Else
                    mnuControl(4).Enabled = False
                End If
            
            optSort(1).Value = True
            chkOptions(0).Value = 1
            optSort(0).Enabled = True
            optSort(1).Enabled = True
            
        Case 2
            optSort(1).Value = True
            chkOptions(0).Value = 1
            optSort(0).Enabled = True
            optSort(1).Enabled = True
            
        Case 3
            If chkOptions(3).Value = 1 Then
                chkOptions(4).Value = 1
            Else
                chkOptions(4).Value = 0
            End If

    End Select

End Sub

Private Sub cmdControls_Click(Index As Integer)
'//command button calls
Dim sMatch  As String
Dim sWord   As String
Dim iPos    As Long

    Select Case Index

        Case 0
            '//replace selected
            sWord = lstSuggest(0).List(lstSuggest(0).ListIndex)
            sMatch = lstSuggest(1).List(lstSuggest(1).ListIndex)
            If (Len(sMatch) = 0) Or (Len(sWord) = 0) Then Exit Sub
            With txtBody
                iPos = InStr(iPos + 1, .Text, sWord, vbTextCompare)
                If iPos > 0 Then
                    .SelStart = iPos - 1
                    .SelLength = Len(sWord)
                    If chkOptions(5) Then
                        .SelColor = vbBlue
                    End If
                    .SelText = sMatch
                End If
            End With
            
        Case 1
            '//replace all
            sWord = lstSuggest(0).List(lstSuggest(0).ListIndex)
            sMatch = lstSuggest(1).List(lstSuggest(1).ListIndex)
            If (Len(sMatch) = 0) Or (Len(sWord) = 0) Then Exit Sub
            With txtBody
                iPos = InStr(iPos + 1, .Text, sWord, vbTextCompare)
                While iPos > 0
                    If iPos > 0 Then
                        .SelStart = iPos - 1
                        .SelLength = Len(sWord)
                        If chkOptions(5) Then
                            .SelColor = vbBlue
                        End If
                        .SelText = sMatch
                        iPos = iPos + Len(sWord)
                        iPos = InStr(iPos + 1, .Text, sWord)
                    End If
                Wend
            End With
            
        Case 2
            '//skip this word
            If chkOptions(4) Then
                sWord = lstSuggest(0).List(lstSuggest(0).ListIndex)
                If Len(sWord) = 0 Then Exit Sub
                With txtBody
                    iPos = InStr(iPos + 1, .Text, sWord, vbTextCompare)
                    While iPos > 0
                        If iPos > 0 Then
                            .SelStart = iPos - 1
                            .SelLength = Len(sWord)
                            .SelBold = False
                            .SelColor = &H80000012
                            iPos = iPos + Len(sWord)
                            iPos = InStr(iPos + 1, .Text, sWord)
                        End If
                    Wend
                End With
            End If
            
        Case 3
            '//add word to database
            sWord = lstSuggest(0).List(lstSuggest(0).ListIndex)
            If Len(sWord) = 0 Then Exit Sub
            Database_Add sWord
            lstSuggest(0).RemoveItem (lstSuggest(0).ListIndex)
            
        Case 4
            '//check spelling
            chkOptions(3).Value = 0
            Auto_Correct

    End Select

End Sub

Private Sub cmdDictionary_Click()
    '//load definition dictionary
    frmDictionary.Show vbModeless, Me
End Sub

Private Sub cmdTranslate_Click()
    '//load translation interface
    frmTranslate.Show vbModeless, Me
End Sub

Private Sub Form_Load()

Dim i As Integer
    '//add filter prc
    For i = 50 To 90 Step 5
        cbFreq.AddItem (i)
    Next i
    '//add time increments
    For i = 1 To 60 Step 1
        cbType.AddItem (i)
    Next i
    lWait = 1
    cbType.Text = lWait
    
    Load_Text
    
    '//dictionary tooltip class
    Set cTip = New ExToolTip


End Sub

Public Sub Dictionary_Init()
    '//load the dictionary w/progress
    If optProgress(0).Value Then
        With frmProgress
            .pBar1.Min = 0
            .pBar1.Max = TMR_MAX
        End With
        Set cDictionary = New clsDictionary
        cDictionary.Dictionary_Init 1
        Unload frmProgress
    Else
        '//dict without prog
        Set cDictionary = New clsDictionary
        cDictionary.Dictionary_Init 2
    End If
    
End Sub
Private Sub Load_Text()

    txtBody.Text = "This is just a bunch of sample text to test this new applicatiun." & vbNewLine & _
                   "Here I intruduce sume sample text, (and punctuation), to " & vbNewLine & _
                   "test the functions of the progrim. My spelling is reilly a bit" & vbNewLine & _
                   "better thun this. But you get the poant, rijht?"
                   
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '//just to be sure..
    Kill_Timer
    '//compress database
    Database_Cleanup
    '//record dictionary choice
    Set_Options
    '//tooltip
    Set cTip = Nothing
    
End Sub

Private Sub lstSuggest_Click(Index As Integer)

Dim sWord As String

On Error Resume Next

    Select Case Index
        Case 0
            '//click on word calls word suggestion routines
            sWord = lstSuggest(0).List(lstSuggest(0).ListIndex)
            '//fix # of concurrent wildcards
            If optTolerance(1).Value Then
                Word_Suggest sWord, 0
            Else
                Word_Suggest sWord, 1
                '//2 wildcard method creates duplicates
                Remove_Duplicates
            End If
            
            '//sort suggest relevance
            If chkOptions(0).Value Then
                    Suggest_Sort sWord
                End If
            
            '//use min filter tolerance
            If chkFilter Then
                Dim i As Integer
                Dim sMatch As String
                For i = lstSuggest(1).ListCount To 1 Step -1
                    sMatch = lstSuggest(1).List(i - 1)
                    If Not Match_Tolerance(sWord, sMatch, cbFreq.Text) Then
                        lstSuggest(1).RemoveItem (i - 1)
                    End If
                Next i
            End If
            
        Case 1
        
    End Select
    
On Error GoTo 0

End Sub

Private Sub Remove_Duplicates()
'//remove suggestion duplicates from list
'//created by 2 wildcard matching
Dim i       As Integer
Dim j       As Integer
Dim sMatch  As String
Dim sWord   As String
Dim item    As ListItem
Dim x       As Integer

On Error Resume Next

    For i = lstSuggest(1).ListCount To 1 Step -1
        x = 0
        sWord = lstSuggest(1).List(i - 1)
            For j = lstSuggest(1).ListCount To 0 Step -1
                If lstSuggest(1).List(j) = sWord Then
                    x = x + 1
                    If x > 1 Then
                        lstSuggest(1).RemoveItem (i - 1)
                    End If
                End If
            Next j
        Next i

On Error GoTo 0

End Sub

Private Sub mnuControl_Click(Index As Integer)
'//menu routines
Dim sMatch  As String
Dim sWord   As String

    Select Case Index
    
        Case 0
            '//cut
            Clipboard.Clear
            Clipboard.SetText txtBody.SelRTF
            txtBody.SelText = vbNullString
        Case 1
            '//copy
            Clipboard.Clear
            Clipboard.SetText txtBody.SelRTF
        Case 2
            '//paste
            txtBody.SelRTF = Clipboard.GetText
        Case 4
            '//correct
            Correct_Item
            '//suggestion items
        Case 7, 8, 9, 10, 11
            With txtBody
                sMatch = mnuControl(Index).Caption
                sWord = .SelText
                '//check for spaces
                If Right$(sWord, 1) = Chr(32) Then
                    sMatch = sMatch & Chr(32)
                ElseIf Left$(sWord, 1) = Chr(32) Then
                    sMatch = Chr(32) & sMatch
                End If
                '//highlite
                If chkOptions(5) Then
                    .SelColor = vbBlue
                Else
                    .SelBold = False
                    .SelColor = &H80000012
                End If
                '//replace
                .SelText = sMatch
            End With
            
    End Select
    
End Sub

Private Sub optCorrectstyle_Click(Index As Integer)
    '//if not timed method chosen, kill the api timer
    If Index > 0 Then Kill_Timer
End Sub

Private Sub optDictionary_Click(Index As Integer)
'//spell check list choice
Dim sTemp   As String
Dim sPath   As String

On Error GoTo Handler

    sPath = App.Path & DBEC_PATH
    '//copy chosen list to 'list.db'
    Select Case Index
        Case 0
            sTemp = App.Path & STEN_PATH
            CopyFile sTemp, sPath, 0
            chkOptions(6).Value = 1
        Case 1
            sTemp = App.Path & MRMW_PATH
            CopyFile sTemp, sPath, 0
            chkOptions(6).Value = 0
        Case 2
            sTemp = App.Path & WBCL_PATH
            CopyFile sTemp, sPath, 0
            chkOptions(6).Value = 0
    End Select
    
    '//decompress and load the new database
    Database_Decompress
    Database_Load
    
Handler:

End Sub

Private Sub txtBody_KeyUp(KeyCode As Integer, Shift As Integer)
'//hmmn.. doesn't always work..
If KeyCode = 8 Then bBack = True Else: bBack = False
End Sub

Private Sub txtBody_MouseUp(Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
'//get mouse event and display appropriate menu items
Dim i As Integer

    If Button <> vbRightButton Then Exit Sub
    '//hide suggestions
    For i = 7 To 11
        mnuControl(i).Visible = False
    Next i
    '//disable/enable options based on suggestion listings
    If lstSuggest(0).ListCount = 0 Then
        For i = 4 To 6
            mnuControl(i).Enabled = False
        Next i
    Else
        For i = 4 To 6
            mnuControl(i).Enabled = True
        Next i
    End If
    
    If chkOptions(1).Value = 1 And lstSuggest(0).ListCount > 0 Then
        mnuControl(4).Enabled = True
    Else
        mnuControl(4).Enabled = False
    End If
    
    If (chkOptions(2).Value = 1) And (mnuControl(6).Enabled = True) Then
        Suggest_MenuItem
    Else
        mnuControl(6).Enabled = False
    End If

    Me.PopupMenu mnuControl1

End Sub

Private Sub Correct_Item()
'//isolate selected word and replace/highlight
Dim sWord   As String
Dim sMatch  As String
Dim i       As Integer
Dim bLSpce  As Boolean
Dim bRSpce  As Boolean

    sWord = txtBody.SelText
    '//test for spaces & errors
    If Right$(sWord, 1) = Chr(32) Then
        sWord = Left$(sWord, (Len(sWord) - 1))
        bRSpce = True
    ElseIf Left$(sWord, 1) = Chr(32) Then
        sWord = Left$(sWord, Len(sWord) - 1)
        bLSpce = True
    ElseIf Len(sWord) = 0 Then
        sWord = Get_Position
        If Len(sWord) = 0 Then
            mnuControl(4).Enabled = False
            Exit Sub
        End If
    End If
            
    For i = 0 To lstSuggest(0).ListCount
    'Debug.Print lstSuggest(0).List(i)
        If sWord = lstSuggest(0).List(i) Then
            lstSuggest(0).Selected(i) = True
            sMatch = lstSuggest(1).List(0)
            
            If Len(sMatch) = 0 Then
                MsgBox "A Replacement Word could Not be Found!!", vbExclamation, "No Replacement"
                Exit Sub
            End If
            
            With txtBody
                '//check highlite
                If chkOptions(5) Then
                    .SelColor = vbBlue
                Else
                    .SelBold = False
                    .SelColor = &H80000012
                End If
                '//reset leading/trailing space to selection
                If bLSpce Then
                    sMatch = Chr(32) & sMatch
                ElseIf bRSpce Then
                    sMatch = sMatch & Chr(32)
                End If
                .SelText = sMatch
            End With
            Exit For
        End If
    Next i

End Sub

Private Sub Suggest_MenuItem()
'//format word and call suggestion routines
Dim sWord   As String
Dim sMatch  As String
Dim i       As Integer
Dim j       As Integer
Dim k       As Integer
Dim iPos    As Long
Dim sTemp   As String

    sWord = txtBody.SelText
    '//test for spaces & errors
    If Right$(sWord, 1) = Chr(32) Then
        sWord = Left$(sWord, (Len(sWord) - 1))
    ElseIf Left$(sWord, 1) = Chr(32) Then
        sWord = Left$(sWord, Len(sWord) - 1)
    ElseIf Len(sWord) = 0 Then
        '//if not highlighted
        sWord = Get_Position
        If Len(sWord) = 0 Then
            mnuControl(4).Enabled = False
            mnuControl(6).Enabled = False
            Exit Sub
        End If
    End If
            
    For i = 0 To lstSuggest(0).ListCount
    'Debug.Print lstSuggest(0).List(i)
        If sWord = lstSuggest(0).List(i) Then
            lstSuggest(0).Selected(i) = True
            k = 7
            For j = 0 To 4
                If lstSuggest(1).ListCount = 0 Then
                    Exit Sub
                ElseIf Len(lstSuggest(1).List(j)) = 0 Then
                    Exit Sub
                End If
                mnuControl(k).Visible = True
                mnuControl(k).Caption = lstSuggest(1).List(j)
                k = k + 1
            Next j
        End If
    Next i
    
End Sub

Private Function Get_Position() As String
'//get the word position and filter out punctuation
Dim x       As Long
Dim l       As Long
Dim m       As Long
Dim y       As Long
Dim sTemp   As String

On Error Resume Next

    With txtBody
        m = .SelStart
        sTemp = Left$(.Text, (m + 1))
        x = InStrRev(sTemp, Chr(32))
        .SelStart = x
        l = InStr(m, .Text, Chr(32)) - 1
        If l = -1 Then l = 0
        
        '//check for punctuation
        y = InStr(m, .Text, vbNewLine) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        '// )
        y = InStr(m, .Text, Chr(29)) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        '// !
        y = InStr(m, .Text, Chr(33)) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        '// "
        y = InStr(m, .Text, Chr(34)) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        '// '
        y = InStr(m, .Text, Chr(45)) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        '// ,
        y = InStr(m, .Text, Chr(44)) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        '// -
        y = InStr(m, .Text, Chr(45)) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        '// .
        y = InStr(m, .Text, Chr(46)) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        '// ?
        y = InStr(m, .Text, Chr(63)) - 1
        If (y < l) And (y > 0) Or (l = 0) And (y > 0) Then l = y
        
        .SelLength = l - x
        Get_Position = .SelText
    End With
    
    
On Error GoTo 0
    
End Function

Private Sub txtBody_LostFocus()

    If (chkOptions(3).Value = 1) And (optCorrectstyle(0).Value) Then
        Kill_Timer
    End If

End Sub

Private Sub Timer_Correct()
 '//the timer is a little cranky, so make sure
 '//you turn off the app by using the close button..
Dim i       As Long
Dim sText   As String
Dim lVal    As Long
Dim lPos    As Long

On Error Resume Next

    '// if the text changes, evaluate length
    If Not Len(txtBody.Text) - lTMem > 24 Then Exit Sub
    With txtBody
        lWait = (CLng(cbType.Text) * 1000)
        i = Len(.Text)
        If i - lTMem > 0 Then
            lPos = i - lTMem
                '//get number of words added
                If Word_Count(lPos) > 10 Then
                    lTMem = i
                    '// over 10 then correct
                    Auto_Correct
                Else
                    '//wait
                    Start_Timer lWait
                End If
        End If
    End With
    
On Error GoTo 0

End Sub

Private Sub Word_Correct()
'//correct by word
Dim aText(0)    As String
Dim sTemp       As String
Dim sFirst      As String

On Error Resume Next

    With txtBody
        If Not Right$(.Text, 1) = Chr(32) Then
            Exit Sub
        Else
            sFirst = Left$(.Text, Len(.Text) - 1)
            sTemp = Mid$(sFirst, InStrRev(sFirst, Chr(32)) + 1)
            If Len(sTemp) = 0 Then sTemp = sFirst
            aText(0) = sTemp
            Search_Control aText(0), 1
        End If
    End With
    
On Error GoTo 0
    
End Sub

Private Sub Line_Correct()
'//correct by line break
Dim sText   As String
Dim aText() As String
Dim aRow()  As String
Dim l       As Long
Dim x       As Long

On Error Resume Next

    If Not Asc(Right$(txtBody.Text, 1)) = 10 Then
        Exit Sub
    Else
        With txtBody
            sText = Left$(.Text, Len(.Text) - 1)
            aRow = Split(sText, vbNewLine)
            x = UBound(aRow)
            aText = Split(aRow(UBound(aRow)), Chr(32))
            For l = 0 To UBound(aText)
                Search_Control aText(l)
                DoEvents
            Next l
            lSLen = Len(.Text)
        End With
    End If

On Error GoTo 0

End Sub

Private Sub Sentence_Correct()
'//correct by sentence (punctuation mark)
Dim lPos    As Long
Dim sTemp   As String
Dim aText() As String
Dim sText   As String
Dim l       As Long
Dim m       As Long
Dim x       As Integer
Dim sCheck  As String

On Error Resume Next

    With txtBody
        x = Asc(Right$(.Text, 1))
        If x = 46 Or x = 63 Or x = 33 Then
            '//figure out the location of the last sentence
            sText = Left$(.Text, Len(.Text) - 1)
            sTemp = Mid$(sText, InStrRev(sText, Chr(46)) + 1)
            l = Len(sTemp)
            sCheck = Mid$(sText, InStrRev(sText, Chr(63)) + 1)
            m = Len(sCheck)
            If (m < l) And (m > 0) Then sTemp = sCheck
            sCheck = Mid$(sText, InStrRev(sText, Chr(33)) + 1)
            m = Len(sCheck)
            l = Len(sTemp)
            If (m < l) And (m > 0) Then sTemp = sCheck
            If Len(sTemp) = 0 Then Exit Sub
            aText = Split(sTemp, Chr(32))
            
            For l = 0 To UBound(aText)
                Search_Control aText(l)
                DoEvents
            Next l
        End If
    End With

On Error GoTo 0

End Sub

Private Sub txtBody_Change()
'//pass to correction routines if selected
    If Not chkOptions(3).Value = 1 Then Exit Sub
    'If bBack Then Exit Sub
    Select Case True
        Case optCorrectstyle(0)
            Timer_Correct
        Case optCorrectstyle(1)
            Word_Correct
        Case optCorrectstyle(2)
            Sentence_Correct
        Case optCorrectstyle(3)
            Line_Correct
    End Select
    'bBack = False
    
End Sub

Private Function Word_Count(ByVal lPos As Long) As Long
'//get total word count
Dim aWord() As String
Dim sComp As String

    With txtBody
        sComp = Right$(.Text, lPos + 1)
        aWord = Split(sComp, Chr(32))
        Word_Count = UBound(aWord)
    End With
    
End Function

Public Sub Auto_Correct()
'//get each word, and rebuild string with correction highlights
Dim aRows() As String
Dim aText() As String
Dim sText   As String
Dim i       As Long
Dim j       As Long
Dim sFrmt   As String
Dim sTemp   As String
Dim sTfrt   As String
Dim iPos    As Long

On Error GoTo Handler

    With frmMain.txtBody
        If Len(.Text) = 0 Then Exit Sub
        If Len(.Text) > 5000 Then chkOptions(4).Value = 0
            '//reset correction coloring
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelColor = &H0
            lstSuggest(0).Clear
            lstSuggest(1).Clear
            sText = .Text
            sTBody = .Text
            aRows = Split(sText, vbNewLine)
                For i = 0 To UBound(aRows)
                    aText = Split(aRows(i), Chr(32))
                    For j = 0 To UBound(aText)
                        Search_Control aText(j)
                        DoEvents
                    Next j
                    DoEvents
                Next i
                
                '//though the highlight functions appears to be a little roundabout
                '//in method, it is actually much faster then using seltext to highlight
                '//corrections..
                If chkOptions(4).Value = 1 Then
                    '//replace linebreaks with rtb tag
                    sTemp = sTBody
                    iPos = InStr(1, sTBody, vbCrLf)
                    Do While iPos > 0
                        sFrmt = "\par "
                        sTfrt = Replace(sTemp, vbCrLf, sFrmt, iPos)
                        sTBody = Left$(sTemp, iPos - 1) & sTfrt
                        iPos = iPos + Len(sTfrt)
                        iPos = InStr(iPos + 1, sTBody, vbCrLf)
                    Loop
                    '//add rtb header data
                    '//might want to split this string up with vars for font type and size..
                    sTBody = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Tahoma;}}" & vbCrLf & _
                         "{\colortbl ;\red255\green0\blue0;}" & vbCrLf & _
                         "\viewkind4\uc1\pard\lang1033\f0\fs17" & vbCrLf & _
                         sTBody & "}"
                    '//paste in results (tried adding as text, didn't work. is a better way maybe? hmmn..)
                    .Text = vbNullString
                    Clipboard.Clear
                    Clipboard.SetText sTBody
                    .SelRTF = Clipboard.GetText
                End If
    End With

Exit Sub

Handler:
Kill_Timer
frmMain.chkOptions(3).Value = 0

End Sub

Private Sub cbType_Click()
    '//timer interval
    lWait = CLng(cbType.Text)
End Sub

Private Sub lstSuggest_DblClick(Index As Integer)
    '//display suggestions
    Select Case Index
        Case 1
            cmdControls(0).Value = True
    End Select
    
End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                     Dictionary Controls
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                                                        

Public Sub cDictionary_Process(ByVal percent As Long)
'//capture loading timer event

On Error Resume Next

    RaiseEvent Process(percent)
    lGCount = lGCount + percent
    With frmProgress
        If percent = TMR_MAX Then
            .pBar1.Value = TMR_MAX
            .lblLoading(0).Caption = "Progress 100 % Complete!"
            Unload Me
        End If

        iPerc = iPerc + 1
        .pBar1.Value = lGCount
        .lblLoading(0).Caption = "Progress " & iPerc & " %"
        .lblLoading(1).Caption = "Loaded " & lGCount & " Entries"
    End With

On Error GoTo 0

End Sub

Public Sub cTip_Status(bStatus As Boolean)
'//return availability of class w/ destroy event

    RaiseEvent Status(bStatus)
    mMove = bStatus

End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

    If mMove Then
        mMove = False
    End If

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                     Definition Popup Controls
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub txtBody_MouseMove(Button As Integer, _
                                    Shift As Integer, _
                                    x As Single, _
                                    y As Single)

'//this feature still needs some work..
'//I would create a pause timer, and
'//lower the tooltip timer to 10 ms
'//if the same word is selected for
'//say 1 sec, then pass to tooltip..

Dim sWord   As String
Dim sDefine As String


    If chkTooltip.Value = 1 Then
        If Not mMove Then
            mMove = True
            '//get word from pointer location
            sWord = Point_Word(txtBody, x, y)
            '//filter the obvious
            If Len(sWord) > 3 Then
                '//fetch definition
                sDefine = Point_Define(sWord)
                If Not Len(sDefine) = 0 Then
                    '//format for tooltip
                    sDefine = Point_Format(sDefine)
                    If sDefine = vbNullString Then
                        Exit Sub
                    End If
                    cTip.ShowToolTip txtBody.hwnd, sWord, sDefine, 0, 95
                Else
                    Exit Sub
                End If
            End If
        End If

        mMove = False

    End If

End Sub

Private Function Point_Format(ByVal sDefine As String) As String
'//format definition for tooltip
Dim iPos    As Long
Dim iSaf    As Long
Dim sFormat As String
Dim sTemp   As String

On Error Resume Next

    sTemp = sDefine
    iPos = 1
    Do Until iPos = 0
        iPos = InStr(50, sTemp, Chr$(32))
        sFormat = sFormat & Left$(sTemp, iPos) & vbNewLine
        sTemp = Right$(sTemp, Len(sTemp) - iPos)
        iSaf = iSaf + 1
        If iSaf > 7 Then
            Exit Function
        End If
    Loop
    sFormat = Left$(sFormat, Len(sFormat) - 2) & sTemp
    Point_Format = sFormat

On Error GoTo 0

End Function

Private Function Point_Word(RTB As RichTextBox, _
                            ByVal x As Long, _
                            ByVal y As Long) As String

'//get selected word with mouse tracking api
Dim lPos  As Long
Dim lStrt As Long
Dim sWord As String
Dim lLen  As Long
Dim sTemp As String
Dim Ptr   As tPtr
Dim xPos  As Long


On Error GoTo Handler

    Ptr.cx = x \ Screen.TwipsPerPixelX
    Ptr.cy = y \ Screen.TwipsPerPixelY
    lPos = SendMessage(RTB.hwnd, CHAR_POS, 0&, Ptr)
    If lPos = 0 Then Exit Function

    With RTB
        sTemp = .Text
        lStrt = InStrRev(sTemp, Chr$(32), lPos)
        lLen = InStr(lPos, sTemp, Chr$(32))
        lLen = lLen - lStrt
        sWord = Mid$(sTemp, lStrt, lLen)
        If InStr(1, sWord, vbNewLine) > 0 Then
            sWord = Mid$(sWord, InStr(sWord, vbNewLine) + 2)
        End If
        If Len(sWord) = 0 Then GoTo Handler
        Point_Word = Filter_Punctuation(sWord)
    End With

Exit Function

Handler:
    Point_Word = vbNullString

End Function

Public Function Point_Define(ByVal sWord As String) As String
'//search collection for selected definitions
Dim sSearch As String
Dim sResult As String

    sSearch = LCase$(sWord) & " ()"
    sResult = cDictionary.Dictionary_Fetch(sSearch)
    If Not sResult = "WNF" Then
        Point_Define = sResult
        Exit Function
    End If

    sSearch = sWord & " (n.)"
    sResult = cDictionary.Dictionary_Fetch(sSearch)
    If Not sResult = "WNF" Then
        Point_Define = sResult
        Exit Function
    End If

    sSearch = sWord & " (v.)"
    sResult = cDictionary.Dictionary_Fetch(sSearch)
    If Not sResult = "WNF" Then
        Point_Define = sResult
        Exit Function
    End If

    sSearch = sWord & " (a.)"
    sResult = cDictionary.Dictionary_Fetch(sSearch)
    If Not sResult = "WNF" Then
        Point_Define = sResult
        Exit Function
    End If

    Point_Define = vbNullString

End Function

