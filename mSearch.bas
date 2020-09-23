Attribute VB_Name = "mSearch"
Option Explicit

Public sTBody As String
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Sub Search_Control(ByVal sWord As String, _
                          Optional ByVal iMode As Integer)
'//root of spell check routine
    sWord = LCase$(sWord)

    If Not Filter_Word(sWord) Then Exit Sub
    sWord = Filter_Punctuation(sWord)
    '//filter a and I
    If Len(sWord) < 2 Then Exit Sub
    '//check spelling
    If Not Word_Compare(sWord) Then
        With frmMain
            .lstSuggest(0).AddItem (sWord)
            If .chkOptions(4) Then
                Word_Highlight sWord, iMode
            End If
        End With
    End If

End Sub

Private Function Filter_Word(ByVal sWord As String) As Boolean
'//filter addresses and dates
Dim i As Integer

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
            For i = 4 To 30
                If InStr(sWord, i & "th") > 0 Then
                    Exit Function
                End If
            Next i
    End Select
    
    Filter_Word = True
    
End Function

Public Function Filter_Punctuation(ByVal sWord As String) As String
'//filter ascii less then chr 64 (non letter vals)
Dim iLen    As Integer
Dim sChr    As String

On Error Resume Next

'"`~#$%^&*(_-+=</'[{>}]|?., -" and db markers = 0 to 63

    iLen = Len(sWord)
    If iLen = 0 Then Exit Function
    
    '//filter punctuation marks
    '//filter low
    sChr = Left$(sWord, 1)
    Do While Asc(sChr) < 64
        sChr = Left$(sWord, 1)
        If Asc(sChr) > 63 Then Exit Do
        If sChr = vbNullString Then Exit Function
        iLen = Len(sWord)
        sWord = Right$(sWord, (iLen - 1))
    Loop
    
    '//filter high
    sChr = Right$(sWord, 1)
    Do While Asc(sChr) < 64
        sChr = Right$(sWord, 1)
        If Asc(sChr) > 63 Then Exit Do
        If sChr = vbNullString Then Exit Function
        iLen = Len(sWord)
        sWord = Left$(sWord, (iLen - 1))
    Loop

    Filter_Punctuation = sWord

On Error GoTo 0

End Function

Public Function Word_Compare(ByVal sWord As String) As Boolean
'//call to hash table lookups
'//with optional suffix attatchment
Dim sTmp As String

    '//word search
    If HashSearch(sWords(), lWords(), sWord) <> -1 Then
        Word_Compare = True
        If frmMain.chkOptions(6).Value = 0 Then Exit Function
        
    '//common extensions
    ElseIf Right$(sWord, 2) = "'s" Then
        sTmp = Left$(sWord, (Len(sWord) - 2))
        If HashSearch(sWords(), lWords(), sTmp) <> -1 Then
            Word_Compare = True
        End If
            
    ElseIf Right$(sWord, 4) = "ally" Then
        sTmp = Left$(sWord, (Len(sWord) - 4))
        If HashSearch(sWords(), lWords(), sTmp) <> -1 Then
            Word_Compare = True
        End If
            
    ElseIf Right$(sWord, 4) = "ity" Then
        sTmp = Left$(sWord, (Len(sWord) - 3))
        If HashSearch(sWords(), lWords(), sTmp) <> -1 Then
            Word_Compare = True
        End If
            
    ElseIf Right$(sWord, 2) = "ed" Then
        sTmp = Left$(sWord, (Len(sWord) - 2))
        If HashSearch(sWords(), lWords(), sTmp) <> -1 Then
            Word_Compare = True
        End If
            
    ElseIf Right$(sWord, 3) = "ion" Then
        sTmp = Left$(sWord, (Len(sWord) - 3)) & "e"
        If HashSearch(sWords(), lWords(), sTmp) <> -1 Then
            Word_Compare = True
        End If
        
    ElseIf Right$(sWord, 3) = "ies" Then
        sTmp = Left$(sWord, (Len(sWord) - 3))
        If HashSearch(sWords(), lWords(), sTmp) <> -1 Then
            Word_Compare = True
        End If
    End If

End Function

Public Sub Word_Wildcards(ByVal sWord, _
                          ByRef aWord() As String, _
                          Optional ByVal iTol As Integer = 0)
'//creates words with 1 or two concurrent wildcards for each letter
'//ex. cest = *est, c*st, ce*t, les*
'//two concurrent wildcards: **st, c**t, ce**
'//each wildcard word is processed and replaced with ascii
'//positive results return suggestion matches
Dim i       As Integer
Dim iLen    As Integer
Dim sChr    As String
Dim sTmp    As String

'//the tolerance factor alters the number of concurrent wildcards
'//placed in a word, this modifies sensitivity, and increases lag time
'//but more results will be returned

    Select Case iTol
        Case 0
            '//assume first letter is correct
            '//and add one letter wildcard to end
            iLen = Len(sWord)
            ReDim aWord(1 To iLen)
                For i = 1 To iLen - 1
                    sChr = Mid$(sWord, i, 1)
                    sTmp = Replace$(sWord, sChr, "*", i, 1, vbBinaryCompare)
                    aWord(i) = Left$(sWord, (i - 1)) & sTmp
                    'Debug.Print aWord(i)
                Next i
        Case 1
            '//use two concurrent wildcards for every letter
            iLen = Len(sWord)
            ReDim aWord(1 To iLen)
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
'//for each wildcard char returned from above
'//replace with ascii char, and lookup in hash table
'//might seem slow, but table lookup is almost instant
'//so even max lookups (26 * 26 for each wildcard pairing)
'//happens in small fraction of a second
'//if you wanted to get fancy, could use hueristic
'//approach, ex. chars are mapped to each other
'//in relative table by keyboard position
'//N,S,E,W. Since many errors are typing errors
'//first routine runs search against substitutions
'//chars derived by relative keyboard mapping.
'//second search uses phoenetic mapping in same way
'//common sounds are mapped to actual spelling
'//ex. f to ph, si to psy etc. and those maps run through with
'//replace routine similar to this one.
'//but from what I can see this method catches most
'//suggestions, if you want to explore though..

Dim aWord() As String
Dim i       As Integer
Dim j       As Integer
Dim sMatch  As String

    With frmMain
        .lstSuggest(1).Clear
    End With
    
    Select Case iTol
        '//97 - 122 lowercase chars
        Case 0
            Word_Wildcards sWord, aWord()
            For i = 1 To UBound(aWord)
                For j = 97 To 122
                    sMatch = Replace$(aWord(i), "*", Chr(j))
                    If HashSearch(sWords(), lWords(), sMatch) <> -1 Then
                        If Len(sMatch) > 0 Then
                            With frmMain
                                .lstSuggest(1).AddItem (sMatch)
                            End With
                        End If
                    End If
                Next j
            Next i
        '//two wildcards used in concurrent running arrays
        Case 1
            Dim k As Integer
            Word_Wildcards sWord, aWord(), 1
            For i = 1 To UBound(aWord) - 1
                For j = 97 To 122
                    For k = 97 To 122
                        sMatch = Replace$(aWord(i), "**", Chr(j) & Chr(k))
                        If HashSearch(sWords(), lWords(), sMatch) <> -1 Then
                            If Len(sMatch) > 0 Then
                                With frmMain
                                    .lstSuggest(1).AddItem (sMatch)
                                End With
                            End If
                        End If
                    Next k
                Next j
                DoEvents
            Next i
    End Select
    
End Function

Private Sub Word_Highlight(ByVal sWord As String, _
                           ByVal iMode As Integer)
'//rewrote the multi select portion without seltext
'//for a single word match speed is not an issue
'//but with a large document, seltext is not viable
'//this method inserts rtb highlight markers around
'//match words, and at the end of the cycle, drops
'//the new text and a header back into the textbox
Dim iPos    As Long
Dim sTemp   As String
Dim sFrmt   As String
Const HGTF  As String = "\cf0 "
Const HGTT  As String = "\cf1 "
Dim sTfrt   As String
Dim sTstw   As String
Dim sTwrd   As String

On Error Resume Next

    Select Case iMode
        
        Case 1
            '//correct one word
            If Len(sWord) = 0 Then Exit Sub
                With frmMain.txtBody
                    sTemp = sTBody
                    iPos = InStrRev(sTemp, sWord, (Len(sTemp)), vbTextCompare) - 1
                    If iPos > 0 Then
                        .Locked = True
                        .SelStart = iPos
                        .SelLength = Len(sWord)
                        '.SelBold = True
                        .SelColor = vbRed
                        .SelLength = 0
                        .SelStart = Len(sTemp)
                        .SelColor = 0
                        .Locked = False
                    End If
                End With
        
        Case Else
            '//correct all words
            If Len(sWord) = 0 Then Exit Sub
                With frmMain.txtBody
                    sTemp = sTBody
                    iPos = InStr(iPos + 1, sTemp, sWord, vbTextCompare)
                    
                    While iPos > 0
                        '//filter word in word scenarios
                        sTwrd = Mid$(sTemp, iPos, Len(sWord))
                        sTstw = Mid$(sTemp, iPos - 1, Len(sWord))
                        sTstw = Left$(sTstw, 1)
                        If (LCase$(sTwrd)) = sWord And (sTstw = Chr(32)) Then
                            sFrmt = HGTT & sWord & HGTF
                            sTfrt = Replace(sTemp, sWord, sFrmt, iPos, 1)
                            sTBody = Left$(sTemp, iPos - 1) & sTfrt
                            iPos = iPos + Len(sTfrt)
                            iPos = InStr(iPos + 1, sTBody, sWord)
                        Else
                            iPos = iPos + 1
                        End If
                    DoEvents
                    Wend
                End With

    End Select
   
On Error GoTo 0

End Sub
