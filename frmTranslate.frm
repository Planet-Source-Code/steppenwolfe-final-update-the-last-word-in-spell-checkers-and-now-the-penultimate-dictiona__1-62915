VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTranslate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Translate: French"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
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
   ScaleHeight     =   2565
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "Translate"
      Height          =   345
      Left            =   3990
      TabIndex        =   3
      Top             =   2010
      Width           =   1215
   End
   Begin VB.TextBox txtWord 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   540
      Width           =   2895
   End
   Begin RichTextLib.RichTextBox rtMain 
      Height          =   885
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   1561
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmTranslate.frx":0000
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
   Begin VB.ComboBox cbTranslate 
      Height          =   315
      Left            =   3180
      TabIndex        =   0
      Top             =   540
      Width           =   2025
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Language:"
      Height          =   195
      Index           =   1
      Left            =   3180
      TabIndex        =   5
      Top             =   330
      Width           =   1245
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word/Phrase:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   330
      Width           =   1005
   End
End
Attribute VB_Name = "frmTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbTranslate_Click()
'//language selection
    cbTranslate.Locked = False
    Select Case cbTranslate.Text
        Case "French"
            Me.Caption = "Translate To: French"
            Translate_Init 1
        Case "Italian"
            Me.Caption = "Translate To: Italian"
            Translate_Init 2
        Case "Spanish"
            Me.Caption = "Translate To: Spanish"
            Translate_Init 3
    End Select
    'cbTranslate.Locked = True
    
End Sub

Private Sub cmdTranslate_Click()
'//translate selected
Dim sWord As String

    With txtWord
        rtMain.Text = vbNullString
        If Len(.Text) > 0 Then
            sWord = .Text
            rtMain.Text = Translate_Fetch(sWord)
        End If
    End With

End Sub

Private Sub Form_Load()

    '//load the file
    cbTranslate.AddItem "French", 0
    cbTranslate.AddItem "Italian", 1
    cbTranslate.AddItem "Spanish", 2
    cbTranslate.Text = "French"
    '//load default french
    Translate_Init 1
    
End Sub

Private Sub Sample_Format_Step_1()

'//you can reuse the translation routine in any way
'//you like, you only need to format the file correctly
'//this first routine takes a file that was originally
'//formatted for a database with a fixed length record
'//type of 84 chars long. Step 1 strips out each record
'//and formats them 1 record per line..

Dim sPath1      As String
Dim sPath2      As String
Dim sResult     As String
Dim l           As Long
Dim lIncr       As Long
Dim sTemp       As String
Dim DLM         As String
Dim aResult()   As String
Dim sLeft       As String
Dim sRight      As String
Dim x           As Long

    DLM = Chr(32)

    sPath1 = App.Path & "\input.txt"
    sPath2 = App.Path & "\output.dic"
    '//get string from file
    Open sPath1 For Binary As #1
        sResult = Space$(LOF(1))
    Get #1, , sResult
    Close #1

    '//get each record and seperate by line
    '//in this case fixed length record (84 bytes)
    '//but may be any kind of record seperator
    '//ex. vbNewLine, or dbase seperator
    Open sPath2 For Binary As #1
    x = 1
    Do
        sLeft = Mid$(sResult, x, 84)
        sRight = Trim$(sLeft) & vbNewLine
        Put #1, , sRight
        x = x + 84
        DoEvents
    Loop Until sLeft = vbNullString
    Close #1

    
End Sub

Private Sub Sample_Format_Step_2()
'//step 2 splits each line into left and right
'//(in this case word: search word or key in the collection
'//and item, or in this case the translation entry and
'//item in the collection..)
'//each key and item are seperated by a unit seperator (Chr(30))
'//and each pairing is seperated by a record seperator (Chr(31))
'//you only need to figure out how original file is seperated
'//to rebuild any file with these routines..
'//for example a dutch or cantonese translation dictionary
'//or a thesaurus, you only need to find the file you want to add,
'//format it, (add unicode support if necessary), and reuse
'//the collection routines in mTranslate to implement..

Dim sPath1      As String
Dim sPath2      As String
Dim sResult     As String
Dim l           As Long
Dim lIncr       As Long
Dim sTemp       As String
Dim DLM         As String
Dim aResult()   As String
Dim sLeft       As String
Dim sRight      As String
Dim x As Long

    DLM = Chr(32) & Chr(32)

    sPath1 = App.Path & "\output.txt"
    sPath2 = App.Path & "\output.dic"
    '//get string from file
    Open sPath1 For Binary As #1
        sResult = Space$(LOF(1))
    Get #1, , sResult
    Close #1
    
    '//reformat file with dbase seperators
    Open sPath2 For Binary As #1
    aResult = Split(sResult, vbNewLine)
    For l = 0 To UBound(aResult)
        sLeft = Trim$(Left$(aResult(l), 31))
        sRight = Trim$(Mid$(aResult(l), 31))
        'Debug.Print sLeft
        'Debug.Print sRight
        Put #1, , sLeft & Chr(31) & sRight & Chr(30)
        DoEvents
    Next l
    Close #1
    
    
End Sub

