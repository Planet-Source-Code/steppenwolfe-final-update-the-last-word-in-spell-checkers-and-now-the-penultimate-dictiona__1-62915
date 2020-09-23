Attribute VB_Name = "mTable"
Option Explicit
'Tri Sort credits go to Philippe Lord -> awesome routines

Private Const ERROR_NOT_FOUND As Long = &H80000000

Public bDimn    As Boolean

Public Enum SortOrder
   SortAscending = 0
   SortDescending = 1
End Enum

Public Enum RemoveFrom
   RemoveArray = 0
   RemoveIndex = 1
End Enum


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, _
                                                                     ByRef lpSource As Any, _
                                                                     ByVal iLen As Long)

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

Public Sub TriQuickSortString(ByRef sArray() As String)

   Dim iLBound As Long
   Dim iUBound As Long
   Dim i       As Long
   Dim j       As Long
   Dim sTemp   As String
   
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   TriQuickSortString2 sArray, 4, iLBound, iUBound
   InsertionSortString sArray, iLBound, iUBound

End Sub

Private Sub TriQuickSortString2(ByRef sArray() As String, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String

   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2
      
      If sArray(iMin) > sArray(i) Then SwapStrings sArray(iMin), sArray(i)
      If sArray(iMin) > sArray(iMax) Then SwapStrings sArray(iMin), sArray(iMax)
      If sArray(i) > sArray(iMax) Then SwapStrings sArray(i), sArray(iMax)
      
      j = iMax - 1
      SwapStrings sArray(i), sArray(j)
      i = iMin
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4 ' sTemp = sArray(j)
      
      Do
         Do
            i = i + 1
         Loop While sArray(i) < sTemp
         
         Do
            j = j - 1
         Loop While sArray(j) > sTemp
         
         If j < i Then Exit Do
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
    If i = 0 Then CopyMemory ByVal VarPtr(i), ByVal VarPtr(s1), 4
   
    CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
    CopyMemory ByVal VarPtr(s2), i, 4
    
End Sub
   
Private Sub InsertionSortString(ByRef sArray() As String, ByVal iMin As Long, ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String

   For i = iMin + 1 To iMax
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(i)), 4 ' sTemp = sArray(i)
      j = i
      
      Do While j > iMin
         If sArray(j - 1) <= sTemp Then Exit Do

         CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - 1)), 4 ' sArray(j) = sArray(j - 1)
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
    If Not bDimn Then
        ReDim sArray(0)
        sArray(0) = sStringToAdd
        bDimn = True
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
   
   If iPos < 0 Then iPos = 0
   
   iUBound = iUBound + 1
   ReDim Preserve sArray(iUBound)
   
   CopyMemory ByVal VarPtr(sArray(iPos + 1)), ByVal VarPtr(sArray(iPos)), (iUBound - iPos) * 4
   
   iTemp = 0
   CopyMemory ByVal VarPtr(sArray(iPos)), iTemp, 4
   
   sArray(iPos) = sStringToAdd

On Error GoTo 0

End Sub
