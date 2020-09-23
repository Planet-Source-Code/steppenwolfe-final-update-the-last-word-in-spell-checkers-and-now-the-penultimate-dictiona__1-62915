Attribute VB_Name = "mFrequency"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)
                                                                     
Private Declare Function StrCmpLogicalP Lib "Shlwapi.dll" Alias "StrCmpLogicalW" (ByVal ptr1 As Long, _
                                                                                  ByVal ptr2 As Long) As Long

Public Function StrCmpLogical(str1 As String, _
                              str2 As String) As Long

    StrCmpLogical = StrCmpLogicalP(ByVal StrPtr(str1), ByVal StrPtr(str2))

End Function

Public Sub Suggest_Sort(ByVal sWord As String)
'//scores by positioning of char in asccii table
'// -1 lower/0 equal/1 higher
'//each character in word compared against
'//base word and converted to percentages for
'//numerical match positioning
Dim aWord()     As String
Dim i           As Integer
Dim j           As Integer
Dim l           As Long
Dim sChr1       As String
Dim sChr2       As String
Dim lCount      As Long
Dim dPercent    As Double
Dim iTol        As Double
Dim aResult()   As String

On Error Resume Next
    
    bDimn = False
    
    With frmMain.lstSuggest(1)
        If .ListCount = 0 Then Exit Sub
        i = .ListCount
        ReDim aWord(0 To (i - 1))
        For i = 0 To .ListCount - 1
            aWord(i) = .List(i)
        Next i
    End With
    
    '//check each char for a score value
    '//using StrCmpLogicalP api
    For j = 0 To UBound(aWord)
        For i = 1 To Len(sWord)
            sChr1 = Left$(sWord, i)
            sChr1 = Mid$(sChr1, i, 1)
        
            If sChr1 = vbNullString Then
                Exit Sub
            End If
        
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
        
        With frmMain
            i = 0
            .lstSuggest(1).Clear
            If Not .optSort(0).Value Then
                '//descending order
                For i = 0 To UBound(aResult)
                    .lstSuggest(1).AddItem aResult(i)
                Next i
            Else
                '//ascending
                For i = UBound(aResult) To 0 Step -1
                    .lstSuggest(1).AddItem aResult(i)
                Next i
            End If
        End With

On Error GoTo 0

End Sub

Public Function Match_Tolerance(ByVal sWord As String, _
                                ByVal sComp As String, _
                                ByVal iTol As Integer) As Boolean
'//same as above, only one word
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
        If sChr1 = vbNullString Then
            Exit Function
        End If
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

