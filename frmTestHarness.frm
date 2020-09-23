VERSION 5.00
Begin VB.Form frmTestHarness 
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   4620
      Width           =   1575
   End
   Begin VB.ListBox lstSuggest 
      Height          =   1620
      Left            =   4350
      TabIndex        =   4
      Top             =   2790
      Width           =   3765
   End
   Begin VB.ListBox lstMatch 
      Height          =   1620
      Left            =   120
      TabIndex        =   3
      Top             =   2790
      Width           =   3735
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Spelling"
      Height          =   375
      Left            =   6300
      TabIndex        =   2
      Top             =   4620
      Width           =   1845
   End
   Begin VB.TextBox txtTest 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   390
      Width           =   7995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Suggestions"
      Height          =   195
      Index           =   2
      Left            =   4350
      TabIndex        =   7
      Top             =   2580
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Not Found"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2580
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Text"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   885
   End
End
Attribute VB_Name = "frmTestHarness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//simple demostration of how my spellcheck project could be made into a class or active-x library
'//with just a few modifications. To simplfy demonstration, I removed compression, tooltips, dictionary etc.
'//and placed relative routines in a class, added some properties and events and call them from a form.
'//This is only a simple demonstration, the full project is at:
'//http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62915&lngWId=1&txtForceRefresh=1016200523424785279
'//so if you want to vote, vote for that project please.
'//this example was put together rather quickly, about half an hour, so if you plan to use it in this form,
'//please rewrite it properly.. ie place all the proper properties and methods in the class

Private WithEvents cPublic     As clsPublic
Attribute cPublic.VB_VarHelpID = -1
Public Event iERR(iErrnum As Integer)
Public Event sMTCH(sSuggest As String)
Public Event sNFD(sReturn As String)


Private Sub cmdCheck_Click()
    Get_Words
End Sub

Private Sub cmdReplace_Click()

Dim sWord   As String
Dim sRplce  As String

    sWord = lstMatch.List(lstMatch.ListIndex)
    sRplce = lstSuggest.List(lstSuggest.ListIndex)
    If Len(sWord) = 0 Then
        MsgBox "Please highlight a word from the list", vbExclamation, "No Selection!"
        Exit Sub
    End If
    
    If Len(sRplce) = 0 Then
        MsgBox "Please highlight a replacement word from the list", vbExclamation, "No Selection!"
        Exit Sub
    End If
    
    Replace_Word sWord, sRplce

End Sub

Private Sub Replace_Word(ByVal sWord As String, _
                         ByVal sRplce As String)
'//simple demonstration of replace. There is a better way to track
'//changes, by running string through do loop and identifying each
'//instance of the word, (see full example Word_Highlight routine for an example)
'//using string parse and replacing body of content with corrected version,
'//it doesn't matter what object you bind the class to, as there is no
'//reliance on specific object properties
Dim sTemp   As String
Dim sTfrt   As String

On Error Resume Next

        With txtTest
            sTemp = .Text
            sTfrt = Replace(sTemp, sWord, sRplce, 1, 1)
            .Text = sTfrt
        End With
        
On Error GoTo 0

End Sub

Private Sub cPublic_iERR(iErrnum As Integer)
'//use event to destroy class if dbase is missing
On Error Resume Next

    RaiseEvent iERR(iErrnum)
    cPublic.Destroy
    Err.Raise 52, , "The Database File appears to be Missing or Corrupt!" & vbNewLine & _
     "Please Reinstall the Application to use this feature."

On Error GoTo 0

End Sub

Private Sub cPublic_sMTCH(sSuggest As String)
'//event raised if suggestions found
On Error Resume Next

    RaiseEvent sMTCH(sSuggest)
    lstSuggest.AddItem sSuggest

On Error GoTo 0

End Sub

Private Sub cPublic_sNFD(sReturn As String)
'//event raised if word not found in hash table
On Error Resume Next

    RaiseEvent sNFD(sReturn)
    lstMatch.AddItem sReturn

On Error GoTo 0

End Sub

Private Sub Form_Load()

    Set cPublic = New clsPublic

    txtTest.Text = "This is just a banch of sumple text to test this new applicatiun."
    '//load class
    cPublic.Init

End Sub

Private Sub Get_Words()
'//parse words from text and pass to class
'//for spell check
Dim aRows() As String
Dim aText() As String
Dim sText   As String
Dim i       As Long
Dim j       As Long

On Error GoTo Handler

    With frmTestHarness.txtTest
        If Len(.Text) = 0 Then Exit Sub
        If Len(.Text) > 5000 Then Exit Sub
        sText = .Text
        aRows = Split(sText, vbNewLine)
        For i = 0 To UBound(aRows)
            aText = Split(aRows(i), Chr$(32))
            For j = 0 To UBound(aText)
                cPublic.Search aText(j)
                DoEvents
            Next j
            DoEvents
        Next i
    End With

Handler:

End Sub

Private Sub lstMatch_Click()
'//d-click on list gets suggestions from class
Dim sWord  As String
Dim i      As Integer
Dim sMatch As String

On Error Resume Next

    lstSuggest.Clear
    sWord = lstMatch.List(lstMatch.ListIndex)
    If Len(sWord) = 0 Then Exit Sub
    cPublic.Word_Suggest sWord, 0

On Error GoTo 0

End Sub

