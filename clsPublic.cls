VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event iERR(iErrnum As Integer)
Public Event sNFD(sReturn As String)
Public Event sMTCH(sSuggest As String)
Private CFm_sWord                As String

'//database
Private Const DATA_PATH          As String = "\list.db"

Private Type Dictionary
    Filename                         As String
    Words                            As Long
    Seperator                        As String
End Type

Private dData                    As Dictionary

'//sorting
Public Enum SortOrder
    SortAscending = 0
    SortDescending = 1
End Enum

Public Enum RemoveFrom
    RemoveArray = 0
    RemoveIndex = 1
End Enum

Private Const ERROR_NOT_FOUND    As Long = &H80000000
Private CFm_bDimn                As Boolean

Private sWords()                 As String
Private lWords()                 As Long
Private Declare Function StrCmpLogicalP Lib "Shlwapi.dll" Alias "StrCmpLogicalW" (ByVal ptr1 As Long, _
                                                                                  ByVal ptr2 As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)


Public Sub Init()
    Database_Load
End Sub

Public Sub Destroy()
    'Set clsPublic = Nothing
End Sub

Public Sub Search(clWord As String)
    Search_Control clWord
End Sub

Public Property Get clWord() As String
    clWord = CFm_sWord
End Property

Public Property Let clWord(ByVal PropVal As String)
    CFm_sWord = PropVal
End Property

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                         Database Routines
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Function Database_Extract(ByVal sFile As String) As String

'//put data to string

    Open sFile For Binary As #1
    Database_Extract = Space$(LOF(1))
    Get #1, , Database_Extract
    Close #1

End Function

Public Sub Database_Load()

Dim sPath   As String

On Error GoTo Handler

    sPath = App.Path & DATA_PATH
    '//build hashtable structure
    dData.Seperator = vbNewLine
    dData.Filename = sPath
    sWords = Split(Database_Extract(dData.Filename), dData.Seperator)
    ReDim Preserve sWords(UBound(sWords) - 1)
    dData.Words = CLng(UBound(sWords))
    BuildHashTable sWords(), lWords()

Exit Sub

Handler:
    RaiseEvent iERR(51)

End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                         Frequency Routines
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function StrCmpLogical(str1 As String, _
                              str2 As String) As Long

    StrCmpLogical = StrCmpLogicalP(ByVal StrPtr(str1), ByVal StrPtr(str2))

End Function

Public Function Match_Tolerance(ByVal sWord As String, _
                                ByVal sComp As String, _
                                ByVal iTol As Integer) As Boolean

Dim i        As Integer
Dim l        As Long
Dim sChr1    As String
Dim sChr2    As String
Dim lCount   As Long
Dim dPercent As Double

On Error Resume Next

    For i = 1 To Len(sWord)
        sChr1 = Left$(sWord, i)
        sChr1 = Mid$(sChr1, i, 1)
        If sChr1 = vbNullString Then Exit Function
        sChr2 = Left$(sComp, i)
        sChr2 = Mid$(sChr2, i, 1)
        If sChr1 = vbNullString Then
            lCount = lCount + 1
            GoTo Skip
        End If
        l = StrCmpLogical(sChr1, sChr2)
        If l <> 0 Then
            lCount = lCount + 1
        End If
Skip:
    Next i
    If Len(sComp) > Len(sWord) Then
        lCount = lCount + (Len(sComp) - Len(sWord))
    End If

    dPercent = (100 / Len(sWord))
    dPercent = (Len(sWord) - lCount) * dPercent

    If CInt(dPercent) >= CInt(iTol) Then
        Match_Tolerance = True
    End If

On Error GoTo 0

End Function

Public Sub Suggest_Sort(ByRef aWord() As String, _
                        ByVal sWord As String)

Dim aWord()   As String
Dim i         As Integer
Dim j         As Integer
Dim l         As Long
Dim sChr1     As String
Dim sChr2     As String
Dim lCount    As Long
Dim dPercent  As Double
Dim iTol      As Double
Dim aResult() As String

On Error Resume Next

    bDimn = False

    '//check each char for a score value
    '//using StrCmpLogicalP api
    For j = 0 To UBound(aWord)
        For i = 1 To Len(sWord)
            sChr1 = Left$(sWord, i)
            sChr1 = Mid$(sChr1, i, 1)

            If sChr1 = vbNullString Then Exit Sub

            sChr2 = Left$(aWord(j), i)
            sChr2 = Mid$(sChr2, i, 1)

            If sChr1 = vbNullString Then
                lCount = lCount + 1
                GoTo Skip
            End If

            l = StrCmpLogical(sChr1, sChr2)
            If l <> 0 Then
                lCount = lCount + 1
            End If
Skip:
        Next i

        '//if compare exceeds base length add points
        If Len(aWord(j)) > Len(sWord) Then
            lCount = lCount + (Len(aWord(j)) - Len(sWord))
        End If

        '//convert score total to percentage
        dPercent = (100 / Len(sWord))
        dPercent = (Len(sWord) - lCount) * dPercent

        '//compare percentages and add to sorted array
        If CInt(dPercent) >= CInt(iTol) Then
            '//add to top
            AddToStringArray aResult, aWord(j), -1
        Else
            '//add to bottom
            AddToStringArray aResult, aWord(j), 0
        End If
        iTol = dPercent
    Next j

    i = 0
    ReDim aWord(0 To UBound(aResult)) As String
    aWord() = aResult()

On Error GoTo 0

End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                         Search Routines
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Function Filter_Word(ByVal sWord As String) As Boolean

Dim i As Integer

    '//filter addresses and dates
    Select Case True
    Case InStr(sWord, "@") > 0
        Exit Function
    Case InStr(sWord, "http") > 0
        Exit Function
    Case InStr(sWord, "ftp") > 0
        Exit Function
    Case InStr(sWord, "1st") > 0
        Exit Function
    Case InStr(sWord, "2nd") > 0
        Exit Function
    Case InStr(sWord, "3rd") > 0
        Exit Function
    Case InStr(sWord, "th") > 0
        For i = 4 To 29
            If InStr(sWord, i & "th") > 0 Then Exit Function
        Next i
    End Select

    Filter_Word = True

End Function

Private Function Filter_Punctuation(ByVal sWord As String) As String

Dim iLen As Integer
Dim sChr As String

On Error Resume Next

    '"`~#$%^&*(_-+=</'[{>}]|?., -" and db markers = 0 to 63

    iLen = Len(sWord)
    If Not iLen = 0 Then
        '//filter punctuation marks
        '//filter low
        sChr = Left$(sWord, 1)
        Do While Asc(sChr) < 64
            sChr = Left$(sWord, 1)
            If Asc(sChr) > 63 Then
                Exit Do
            End If
            If sChr = vbNullString Then Exit Function
            iLen = Len(sWord)
            sWord = Right$(sWord, (iLen - 1))
        Loop

        '//filter high
        sChr = Right$(sWord, 1)
        Do While Asc(sChr) < 64
            sChr = Right$(sWord, 1)
            If Asc(sChr) > 63 Then
                Exit Do
            End If
            If sChr = vbNullString Then Exit Function
            iLen = Len(sWord)
            sWord = Left$(sWord, (iLen - 1))
        Loop

        Filter_Punctuation = sWord
    End If
    
On Error GoTo 0

End Function

Public Function Word_Compare(ByVal sWord As String) As Boolean

Dim sTmp As String

    Select Case True
        '//word search
    Case HashSearch(sWords(), lWords(), sWord) <> -1
        Word_Compare = True

        '//common extension
    Case Right$(sWord, 1) = "s"
        sTmp = Left$(sWord, (Len(sWord) - 1))
        If HashSearch(sWords(), lWords(), sTmp) <> -1 Then
            Word_Compare = True
        End If
    End Select

End Function

Public Sub Word_Wildcards(ByVal sWord, _
                          ByRef aWord() As String, _
                          Optional ByVal iTol As Integer = 0)

Dim i    As Integer
Dim iLen As Integer
Dim sChr As String
Dim sTmp As String

    '//the tolerance factor alters the number of concurrent wildcards
    '//placed in a word, this modifies sensitivity, and increases lag time
    '//but more results will be returned
    Select Case iTol
    Case 0
        '//assume first letter is correct
        '//and add one letter wildcard to end
        iLen = Len(sWord)
        ReDim aWord(1 To iLen) As String
        For i = 1 To iLen - 1
            sChr = Mid$(sWord, i, 1)
            sTmp = Replace$(sWord, sChr, "*", i, 1, vbBinaryCompare)
            aWord(i) = Left$(sWord, (i - 1)) & sTmp
            'Debug.Print aWord(i)
        Next i
    Case 1
        '//use two wildcards for every letter
        iLen = Len(sWord)
        ReDim aWord(1 To iLen) As String
        For i = 1 To iLen - 1
            sChr = Mid$(sWord, (i + 1), 2)
            sTmp = Replace$(sWord, sChr, "**", i, 1, vbBinaryCompare)
            aWord(i) = Left$(sWord, (i - 1)) & sTmp
            'Debug.Print aWord(i)
        Next i
    End Select

End Sub

Public Function Word_Suggest(ByVal sWord As String, _
                             Optional ByVal iTol As Integer) As String


Dim aWord() As String
Dim i       As Integer
Dim j       As Integer
Dim sMatch  As String
Dim k       As Integer

    Select Case iTol
        '//97 - 122 lowercase chars
    Case 0
        Word_Wildcards sWord, aWord()
        For i = 1 To UBound(aWord)
            For j = 97 To 122
                sMatch = Replace$(aWord(i), "*", Chr$(j))
                If HashSearch(sWords(), lWords(), sMatch) <> -1 Then
                    If Len(sMatch) > 0 Then
                        RaiseEvent sMTCH(sMatch)
                    End If
                End If
            Next j
        Next i
        '//two wildcards used
    Case 1
        Word_Wildcards sWord, aWord(), 1
        For i = 1 To UBound(aWord) - 1
            For j = 97 To 122
                For k = 97 To 122
                    sMatch = Replace$(aWord(i), "**", Chr$(j) & Chr$(k))
                    If HashSearch(sWords(), lWords(), sMatch) <> -1 Then
                        If Len(sMatch) > 0 Then
                            RaiseEvent sMTCH(sMatch)
                        End If
                    End If
                Next k
            Next j
            DoEvents
        Next i
    End Select

End Function

Public Sub Search_Control(ByVal sWord As String)

    sWord = LCase$(sWord)

    If Filter_Word(sWord) Then
        sWord = Filter_Punctuation(sWord)
        '//filter a and I
        If Len(sWord) < 2 Then Exit Sub

        '//check spelling
        If Not Word_Compare(sWord) Then
            RaiseEvent sNFD(sWord)
        End If

    End If

End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                         Sorting Routines
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Property Get bDimn() As Boolean
    bDimn = CFm_bDimn
End Property

Public Property Let bDimn(ByVal PropVal As Boolean)
    CFm_bDimn = PropVal
End Property

Public Sub BuildHashTable(ByRef sArray() As String, _
                          ByRef iHashArray() As Long)

Dim iLBound  As Long
Dim iUBound  As Long
Dim iUBound2 As Long
Dim iMax     As Long
Dim iIndex   As Long
Dim i        As Long

    iLBound = LBound(sArray)
    iUBound = UBound(sArray)
    iMax = (iUBound + 1) * 4
    ReDim iHashArray(0 To iMax - 1) As Long
    iUBound2 = UBound(iHashArray)

    For i = LBound(iHashArray) To iUBound2
        iHashArray(i) = ERROR_NOT_FOUND
    Next i

    For i = iLBound To iUBound
        iIndex = GetFastXorHash(sArray(i)) Mod iMax
        Do Until iHashArray(iIndex) = ERROR_NOT_FOUND
            iIndex = (iIndex + 1) Mod iMax
        Loop
        iHashArray(iIndex) = i
    Next i

End Sub

Private Function GetFastXorHash(ByVal sString As String, _
                                Optional ByVal iLenToHash As Long = -1) As Long

Dim iUBound   As Long
Dim iBuffer() As Long
Dim i         As Long

    If sString = vbNullString Then
        GetFastXorHash = -1
    Else
        If iLenToHash = -1 Then
            iLenToHash = Len(sString)
        End If
        If iLenToHash > Len(sString) Then
            iLenToHash = Len(sString)
        End If
        iUBound = iLenToHash \ 4 + 1
        ReDim iBuffer(iUBound) As Long
        CopyMemory iBuffer(0), ByVal sString, iLenToHash
        For i = 0 To iUBound
            GetFastXorHash = GetFastXorHash Xor iBuffer(i) Xor i
        Next i
        GetFastXorHash = GetFastXorHash And &H7FFFFFFF
    End If

End Function

Public Function HashSearch(ByRef sArray() As String, _
                           ByRef iHashArray() As Long, _
                           ByVal sFind As String) As Long

Dim iMax        As Long
Dim bInitialize As Boolean
Dim i           As Long

On Error GoTo Handler

    If UBound(iHashArray) = -1 Then
        bInitialize = True
    Else
        If iHashArray(LBound(iHashArray)) = iHashArray(UBound(iHashArray)) Then
            bInitialize = True
        End If
    End If
    If bInitialize Then
        BuildHashTable sArray, iHashArray
    End If
    iMax = UBound(iHashArray) + 1
    i = GetFastXorHash(sFind) Mod iMax
    Do Until iHashArray(i) = ERROR_NOT_FOUND
        If sArray(iHashArray(i)) = sFind Then
            HashSearch = iHashArray(i)
            Exit Function
        End If
        i = (i + 1) Mod iMax
    Loop
    HashSearch = -1

Exit Function

Handler:
    HashSearch = 20

End Function

Private Sub TriQuickSortString2(ByRef sArray() As String, _
                                ByVal iSplit As Long, _
                                ByVal iMin As Long, _
                                ByVal iMax As Long)
                                
Dim i     As Long
Dim j     As Long
Dim sTemp As String

    If (iMax - iMin) > iSplit Then
        i = (iMax + iMin) / 2

        If sArray(iMin) > sArray(i) Then
            SwapStrings sArray(iMin), sArray(i)
        End If
        If sArray(iMin) > sArray(iMax) Then
            SwapStrings sArray(iMin), sArray(iMax)
        End If
        If sArray(i) > sArray(iMax) Then
            SwapStrings sArray(i), sArray(iMax)
        End If

        j = iMax - 1
        SwapStrings sArray(i), sArray(j)
        i = iMin
        CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4

        Do
            Do
                i = i + 1
            Loop While sArray(i) < sTemp

            Do
                j = j - 1
            Loop While sArray(j) > sTemp

            If j < i Then
                Exit Do
            End If
            SwapStrings sArray(i), sArray(j)
        Loop

        SwapStrings sArray(i), sArray(iMax - 1)
        TriQuickSortString2 sArray, iSplit, iMin, j
        TriQuickSortString2 sArray, iSplit, i + 1, iMax
    End If

    i = 0
    CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4

End Sub

Private Sub SwapStrings(ByRef s1 As String, _
                        ByRef s2 As String)

Dim i   As Long

    i = StrPtr(s1)
    If i = 0 Then
        CopyMemory ByVal VarPtr(i), ByVal VarPtr(s1), 4
    End If

    CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
    CopyMemory ByVal VarPtr(s2), i, 4

End Sub

Private Sub InsertionSortString(ByRef sArray() As String, _
                                ByVal iMin As Long, _
                                ByVal iMax As Long)

Dim i     As Long
Dim j     As Long
Dim sTemp As String

    For i = iMin + 1 To iMax
        CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(i)), 4
        j = i

        Do While j > iMin
            If sArray(j - 1) <= sTemp Then
                Exit Do
            End If

            CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - 1)), 4
            j = j - 1
        Loop
        CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sTemp), 4
    Next i

    i = 0
    CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4

End Sub

Public Sub AddToStringArray(ByRef sArray() As String, _
                            ByVal sStringToAdd As String, _
                            Optional ByVal iPos As Long = -1)

Dim iUBound As Long
Dim iTemp   As Long

On Error Resume Next

    '//check for dimensioning
    If Not CFm_bDimn Then
        ReDim sArray(0) As String
        sArray(0) = sStringToAdd
        CFm_bDimn = True
        Exit Sub
    Else
        iUBound = UBound(sArray)
    End If

    '//if adding at the end
    If (iPos > iUBound) Or (iPos = -1) Then
        ReDim Preserve sArray(iUBound + 1)
        sArray(iUBound + 1) = sStringToAdd
        Exit Sub
    End If

    If iPos < 0 Then
        iPos = 0
    End If

    iUBound = iUBound + 1
    ReDim Preserve sArray(iUBound)

    CopyMemory ByVal VarPtr(sArray(iPos + 1)), ByVal VarPtr(sArray(iPos)), (iUBound - iPos) * 4

    iTemp = 0
    CopyMemory ByVal VarPtr(sArray(iPos)), iTemp, 4

    sArray(iPos) = sStringToAdd

On Error GoTo 0

End Sub

Public Sub TriQuickSortString(ByRef sArray() As String)

Dim iLBound As Long
Dim iUBound As Long

    iLBound = LBound(sArray)
    iUBound = UBound(sArray)

    TriQuickSortString2 sArray, 4, iLBound, iUBound
    InsertionSortString sArray, iLBound, iUBound

End Sub
