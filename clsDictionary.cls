VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CDEN_DICT                As String = "\sed.dic"
Private Const FIELD_LEN                As Integer = 1024

Public Event Process(ByVal percent As Long)
Private CFm_cList                           As New Collection

Public Property Get cList() As Collection

    Set cList = CFm_cList

End Property

Public Property Set cList(PropVal As Collection)

    Set CFm_cList = PropVal

End Property

Public Sub Dictionary_Init(ByVal iProgress As Integer)

'//using search words for keys in a collection
'//is a very fast method of retieving data
'//I tried an array here, but the split function
'//took almost 30 seconds to load the file!
'//also tested this with an indexed database, but
'//the results return on a collection is so fast
'//that this is the only way to go..
'//added a safe exit to this loop with a counter
'//that uses an event raise to update the progress bar

Dim sPath   As String
Dim sResult As String
Dim l       As Long
Dim sKey    As String
Dim sItem   As String
Dim lIncr   As Long
Dim sTemp   As String
Dim Pos1    As Long
Dim Pos2    As Long

On Error Resume Next

    Set CFm_cList = New Collection
    sPath = App.Path & CDEN_DICT

    Open sPath For Binary As #1
    sResult = Space$(LOF(1))
    Get #1, , sResult
    Close #1

    Select Case iProgress
        '//use progress bar
    Case 1
        lIncr = TMR_MAX / 100
        frmProgress.Visible = True

        Do
            Pos1 = Pos2 + 5
            Pos2 = InStr(Pos1, sResult, vbLf)
            If Pos2 <> 0 Then sTemp = Mid$(sResult, Pos1, Pos2 - Pos1)
            sKey = Left$(sTemp, InStr(1, sTemp, vbTab) - 1)
            sItem = Mid$(sTemp, InStr(1, sTemp, vbTab) + 1)
            cList.Add sItem, sKey
            l = l + 1
            If l = lIncr Then
                RaiseEvent Process(lIncr)
                l = 0
            End If
            If l > 221000 Then GoTo Handler
            DoEvents
        Loop Until Pos2 = 0

        '//no progress bar
        '//user can start searching while it loads
        '//hope they search lower letters though!
    Case 2
        '//I know, object call in class is bad
        '//should probably put this in event on form
        '//but I am not picky, and it works..
        frmMain.Visible = True
        Do
            Pos1 = Pos2 + 5
            Pos2 = InStr(Pos1, sResult, vbLf)
            If Pos2 <> 0 Then sTemp = Mid$(sResult, Pos1, Pos2 - Pos1)
            sKey = Left$(sTemp, InStr(1, sTemp, vbTab) - 1)
            sItem = Mid$(sTemp, InStr(1, sTemp, vbTab) + 1)
            cList.Add sItem, sKey
            If l > 221000 Then GoTo Handler
            DoEvents
        Loop Until Pos2 = 0
    End Select

    Exit Sub

Handler:
    Err.Raise 52, , "The dictionary file appears to be Missing or invalid!"

On Error GoTo 0

End Sub

Public Function Dictionary_Fetch(ByVal sItem As String) As Variant
'//collection lookup
Dim sTemp As String

On Error Resume Next

    sItem = LCase$(Trim$(sItem))
    sTemp = cList.item(sItem)
    If Len(sTemp) > 0 Then
        Dictionary_Fetch = sTemp
    ElseIf Dictionary_Filters(sItem) > vbNullString Then
        Dictionary_Fetch = Dictionary_Filters(sItem)
    Else
        Dictionary_Fetch = "WNF"
    End If
    
On Error GoTo 0

End Function

Private Function Dictionary_Filters(ByVal sItem As String) As String
'//standard suffix filters
Dim sTemp   As String
Dim sResult As String

    Select Case True
    Case Right$(sItem, 2) = "'s"
        sTemp = Left$(sItem, (Len(sItem) - 2))
        sResult = cList.item(sTemp)
        If Len(sResult) > 0 Then
            Dictionary_Filters = sResult
        End If

    Case Right$(sItem, 4) = "ally"
        sTemp = Left$(sItem, (Len(sItem) - 4))
        sResult = cList.item(sTemp)
        If Len(sResult) > 0 Then
            Dictionary_Filters = sResult
        End If

    Case Right$(sItem, 3) = "ity"
        sTemp = Left$(sItem, (Len(sItem) - 3))
        sResult = cList.item(sTemp)
        If Len(sResult) > 0 Then
            Dictionary_Filters = sResult
        End If

    Case Right$(sItem, 2) = "ed"
        sTemp = Left$(sItem, (Len(sItem) - 2))
        sResult = cList.item(sTemp)
        If Len(sResult) > 0 Then
            Dictionary_Filters = sResult
        End If

    Case Right$(sItem, 3) = "ion"
        sTemp = Left$(sItem, (Len(sItem) - 3))
        sResult = cList.item(sTemp)
        If Len(sResult) > 0 Then
            Dictionary_Filters = sResult
        End If

    Case Right$(sItem, 3) = "ies"
        sTemp = Left$(sItem, (Len(sItem) - 3))
        sResult = cList.item(sTemp)
        If Len(sResult) > 0 Then
            Dictionary_Filters = sResult
        End If

    Case Else
        Dictionary_Filters = vbNullString
    End Select

End Function
