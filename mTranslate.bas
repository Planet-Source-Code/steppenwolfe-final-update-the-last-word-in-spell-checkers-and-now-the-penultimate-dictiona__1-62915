Attribute VB_Name = "mTranslate"
Option Explicit

Private Const TRSD_FRE As String = "\efd.dic"
Private Const TRSD_ITA As String = "\eid.dic"
Private Const TRSD_SPA As String = "\esd.dic"
Public cTrl    As Collection

Public Sub Translate_Init(ByVal iVolume As Integer)
'//load collection with key(search word), and item (translation)
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

    '//destroy the old collection
    Set cTrl = Nothing
    Set cTrl = New Collection

    Select Case iVolume
    '//french
    Case 1
        sPath = App.Path & TRSD_FRE
    '//german
    Case 2
        sPath = App.Path & TRSD_ITA
        If Len(Dir(sPath)) > 0 Then
        End If
    '//italian
    Case 3
        sPath = App.Path & TRSD_SPA
    Case Else
        GoTo Handler
    End Select

    Open sPath For Binary As #1
        sResult = Space$(LOF(1))
    Get #1, , sResult
    Close #1

    Do
    Pos1 = Pos2 + 5
    Pos2 = InStr(Pos1, sResult, Chr(30))
        If Pos2 <> 0 Then
            sTemp = Mid$(sResult, Pos1, Pos2 - Pos1)
        End If
        sKey = Left$(sTemp, InStr(1, sTemp, Chr(31)) - 1)
        sItem = Mid$(sTemp, InStr(1, sTemp, Chr(31)) + 1)
        cTrl.Add sItem, sKey
        l = l + 1
        If l > 15000 Then
            GoTo Handler
        End If
        DoEvents
    Loop Until Pos2 = 0

Exit Sub

Handler:
Err.Raise 52, , "The dictionary file appears to be Missing or invalid!"
On Error GoTo 0

End Sub

Public Function Translate_Fetch(ByVal sItem As String) As Variant
'//collection lookup
Dim sTemp As String

On Error Resume Next

    sItem = LCase$(Trim$(sItem))
    sTemp = cTrl.item(sItem)
    If Len(sTemp) > 0 Then
        Translate_Fetch = sTemp
    Else
        Translate_Fetch = "Not Found"
    End If
    
On Error GoTo 0

End Function


