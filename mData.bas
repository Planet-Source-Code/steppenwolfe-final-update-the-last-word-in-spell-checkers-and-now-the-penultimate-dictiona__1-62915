Attribute VB_Name = "mData"
Option Explicit

Private Type Dictionary
    Filename          As String
    Words             As Long
    Seperator         As String
End Type

Private dData     As Dictionary
Public sWords()   As String
Public lWords()   As Long

Public Sub Database_Init(ByVal iState As Integer)
'//check db status
    Select Case iState
    Case 1
        Database_Decompress
        Database_Load
    Case 2
        Database_Load
    Case Else
        GoTo Handler
    End Select

Exit Sub

Handler:
    Err.Raise 52, , "The Database File appears to be Missing or Corrupt!" & vbNewLine & _
     "Please Reinstall the Application to use this feature."

End Sub

Public Sub Database_Compress()
'//compress db and replace file
Dim eCLevel  As eRatio
Dim sCPath   As String
Dim sPath    As String
Dim cHuffman As New clsHuffman

    sCPath = App.Path & DBEC_PATH
    sPath = App.Path & DATA_PATH

    '//alternative to api compression for Win 9x
    If bOVersion Then
        Set cHuffman = New clsHuffman
        cHuffman.EncodeFile sPath, sCPath
        Kill sPath
        Set cHuffman = Nothing
        Exit Sub
    End If

    With frmMain
        Select Case True
        Case .optCompress(0).Value
            eCLevel = cLow
        Case .optCompress(1).Value
            eCLevel = cHigh
        End Select
    End With
    Compress_File sPath, sCPath, eCLevel
    Kill sPath

End Sub

Public Sub Database_Decompress()
'//decompress and make target db
Dim sCPath   As String
Dim sPath    As String
Dim cHuffman As New clsHuffman

    sCPath = App.Path & DBEC_PATH
    sPath = App.Path & DATA_PATH

    '//alternative to api compression for Win 9x
    If bOVersion Then
        Set cHuffman = New clsHuffman
        cHuffman.DecodeFile sCPath, sPath
        Kill sCPath
        Set cHuffman = Nothing
        Exit Sub
    End If

    Decompress_File sCPath, sPath
    Database_Compact
    Kill sCPath

End Sub

Public Sub Database_Load()
'//load db into hash table
Dim sPath   As String

On Error GoTo Handler

    sPath = App.Path & DATA_PATH
    '//build structure
    dData.Seperator = vbNewLine
    dData.Filename = sPath
    sWords = Split(Database_Extract(dData.Filename), dData.Seperator)
    ReDim Preserve sWords(UBound(sWords) - 1)
    dData.Words = CLng(UBound(sWords))

    BuildHashTable sWords(), lWords()

Exit Sub

Handler:
    Err.Raise 52, , "The Database File appears to be Missing or Corrupt!" & vbNewLine & _
     "Please Reinstall the Application to use this feature."

End Sub

Private Function Database_Extract(ByVal sFile As String) As String
'//extract db to string
    Open sFile For Binary As #1
    Database_Extract = Space$(LOF(1))
    Get #1, , Database_Extract
    Close #1

End Function

Private Sub Database_Compact()
'//needed because the compression/decompression adds
'//spaces to the file, not sure why, but I am looking into it
Dim sPath As String
Dim sTmp  As String

On Error GoTo Handler

    sPath = App.Path & DATA_PATH
    '//put data to string
    sTmp = Database_Extract(sPath)
    sTmp = Left$(sTmp, InStr(1, sTmp, vbNullChar))

    Open sPath For Output As #1
    Print #1, sTmp
    Close #1

Handler:

End Sub

Public Sub Database_Add(ByVal sWord As String)
'//add exception to current wordlist
Dim l       As Long
Dim aTemp() As String
Dim sPath   As String

On Error GoTo Handler

    sPath = App.Path & DATA_PATH
    aTemp = Split(Database_Extract(sPath), vbNewLine)
    bDimn = True
    AddToStringArray aTemp(), sWord, -1
    TriQuickSortString aTemp()
    Kill sPath
    Open sPath For Output As #1
    For l = 0 To UBound(aTemp)
        If Not aTemp(l) = vbNullString Then
            Print #1, aTemp(l)
        End If
    Next l
    Close #1

Handler:

End Sub

Public Sub Database_Cleanup()
'//existence check and compress database
Dim sPath  As String
Dim sCPath As String

    sPath = App.Path & DATA_PATH
    sCPath = App.Path & DBEC_PATH

    If File_Exist(sPath) Then
        Database_Compress
    ElseIf File_Exist(sCPath) Then
        Exit Sub
    Else
        GoTo Handler
    End If
    
Exit Sub

Handler:
Err.Raise 52, , "The Database File appears to be Missing or Corrupt!" & vbNewLine & "Please Reinstall the Application to use this feature."

End Sub
