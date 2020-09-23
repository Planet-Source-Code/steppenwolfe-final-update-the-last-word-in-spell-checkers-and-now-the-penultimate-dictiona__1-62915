Attribute VB_Name = "mMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Const TMR_MAX    As Long = 220814
Private bFirst          As Boolean
'//os version
Private Const VER_PLATFORM_WIN32s             As Integer = 0
Private Const VER_PLATFORM_WIN32_WINDOWS      As Integer = 1
Private Const VER_PLATFORM_WIN32_NT           As Integer = 2

'//version structure
Private Type OSVersion
    dwOSVersionInfoSize                           As Long
    dwMajorVersion                                As Long
    dwMinorVersion                                As Long
    dwBuildNumber                                 As Long
    dwPlatformId                                  As Long
    szCSDVersion                                  As String * 128
End Type

'//os compression check switch
Public bOVersion As Boolean

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersion) As Boolean



Public Const DATA_PATH As String = "\list.db"
Public Const DBEC_PATH As String = "\list.edb"
'//word lists
Public Const STEN_PATH As String = "\sten.edb"
Public Const MRMW_PATH As String = "\mrmw.edb"
Public Const WBCL_PATH As String = "\wbcl.edb"

'//initially named with '.txt' extension for psc upload
Public Const FTRN_PATH As String = "\list.txt"
Public Const MOVEFILE_REPLACE_EXISTING       As Long = &H1

Public Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, _
                                                                        ByVal lpNewFileName As String, _
                                                                        ByVal dwFlags As Long) As Long
                                                                        
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                                                                   ByVal lpNewFileName As String, _
                                                                   ByVal bFailIfExists As Long) As Long

Private Sub First_Run()
'//you can rem this sub after db is renamed to compressed ext. 'edb'
'//basically this checks for file and renames from txt to db extension
Dim sTemp As String
Dim sPath As String

    '//default dictionary
    sTemp = App.Path & "\sten.txt"
    sPath = App.Path & STEN_PATH
    If File_Exist(sTemp) Then
        Compress_File sTemp, sPath, 0
        Kill sTemp
        bFirst = True
    End If
    
    '//medium
    sTemp = App.Path & "\mrmw.txt"
    sPath = App.Path & MRMW_PATH
    If File_Exist(sTemp) Then
        Compress_File sTemp, sPath, 0
        Kill sTemp
        bFirst = True
    End If
    
    '//collegiate
    sTemp = App.Path & "\wbcl.txt"
    sPath = App.Path & WBCL_PATH
    If File_Exist(sTemp) Then
        Compress_File sTemp, sPath, 0
        Kill sTemp
        bFirst = True
    End If
    
    '//obsolete - left as a demonstration
    'sTemp = App.Path & "\dic1.txt"
    'If File_Exist(sTemp) Then
    '    Rebuild_Dbase
    'End If
    
    '//definition dictionary
    sTemp = App.Path & "\sed.txt"
    sPath = App.Path & "\sed.dic"
    If File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
        bFirst = True
    End If
    
    '//french translation
    sTemp = App.Path & "\efd.txt"
    sPath = App.Path & "\efd.dic"
    If File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
        bFirst = True
    End If
    
    '//italian translation
    sTemp = App.Path & "\eid.txt"
    sPath = App.Path & "\eid.dic"
    If File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
        bFirst = True
    End If
    
    '//spanish translation
    sTemp = App.Path & "\esd.txt"
    sPath = App.Path & "\esd.dic"
    If File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
        bFirst = True
    End If
    
    sTemp = App.Path & STEN_PATH
    sPath = App.Path & DBEC_PATH
    If Not File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
    End If
    
End Sub

Private Sub First_RunAlt()
'//you can rem this sub after db is renamed to compressed ext. 'edb'
'//same as above but uses huffman rather then nt api for compression
Dim sTemp As String
Dim sPath As String
Dim cHuffman As New clsHuffman

    Set cHuffman = New clsHuffman
        
    '//default dictionary
    sTemp = App.Path & "\sten.txt"
    sPath = App.Path & STEN_PATH
    If File_Exist(sTemp) Then
        cHuffman.EncodeFile sTemp, sPath
        Kill sTemp
    End If
    
    '//medium
    sTemp = App.Path & "\mrmw.txt"
    sPath = App.Path & MRMW_PATH
    If File_Exist(sTemp) Then
        cHuffman.EncodeFile sTemp, sPath
        Kill sTemp
    End If
    
    '//collegiate
    sTemp = App.Path & "\wbcl.txt"
    sPath = App.Path & WBCL_PATH
    If File_Exist(sTemp) Then
        cHuffman.EncodeFile sTemp, sPath
        Kill sTemp
    End If
    
    '//added because PSC upload would not accept 18M dictionary file
    sTemp = App.Path & "\dic1.txt"
    If File_Exist(sTemp) Then
        Rebuild_Dbase
    End If
    
    '//definition dictionary
    sTemp = App.Path & "\sed.txt"
    sPath = App.Path & "\sed.dic"
    If File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
        bFirst = True
    End If
    
    '//french translation
    sTemp = App.Path & "\efd.txt"
    sPath = App.Path & "\efd.dic"
    If File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
        bFirst = True
    End If
    
    '//italian translation
    sTemp = App.Path & "\eid.txt"
    sPath = App.Path & "\eid.dic"
    If File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
        bFirst = True
    End If
    
    '//spanish translation
    sTemp = App.Path & "\esd.txt"
    sPath = App.Path & "\esd.dic"
    If File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
        bFirst = True
    End If
    
    sTemp = App.Path & STEN_PATH
    sPath = App.Path & DBEC_PATH
    If Not File_Exist(sTemp) Then
        MoveFileEx sTemp, sPath, MOVEFILE_REPLACE_EXISTING
    End If
    
End Sub

'///Rem End Here///

Public Sub Main()
'//check for file and status, then load selected spelling list
'//first run is slow, because of one-time compression of all word lists
Dim sTemp As String
Dim sPath As String

    sPath = App.Path & DBEC_PATH
    
    InitCommonControls
    
    Identify_OS
    '//you can rem this call after db's are renamed
    If Not bOVersion Then
        First_Run
    Else
        First_RunAlt
    End If
    
    Select Case Get_DBase
        Case 0
            sTemp = App.Path & STEN_PATH
            CopyFile sTemp, sPath, 0
        Case 1
            sTemp = App.Path & MRMW_PATH
            CopyFile sTemp, sPath, 0
        Case 2
            sTemp = App.Path & WBCL_PATH
            CopyFile sTemp, sPath, 0
    End Select
    
    If File_Exist(App.Path & DBEC_PATH) Then
        Database_Init 1
    ElseIf File_Exist(App.Path & DATA_PATH) Then
        Database_Init 2
    Else
        Database_Init 0
    End If
    
    Get_Options
    With frmMain
        .Show
        .SetFocus
        .Dictionary_Init
    End With
    
        
End Sub

Private Sub Rebuild_Dbase()
'//solved my upload timeout problems, but left this as a demonstration
'//of recombining a binary file

Dim sTpath1 As String
Dim sTpath2 As String
Dim sTpath3 As String
Dim sTpath4 As String
Dim sPath   As String
Dim sTemp   As String

    sTpath1 = App.Path & "\dic1.txt"
    sTpath2 = App.Path & "\dic2.txt"
    sTpath3 = App.Path & "\dic3.txt"
    sTpath4 = App.Path & "\dic4.txt"
    sPath = App.Path & "\sed.txt"

    Open sPath For Binary As #1
        Open sTpath1 For Binary As #2
            sTemp = Space$(LOF(2))
            Get #2, , sTemp
            Put #1, , sTemp
        Close #2
        Kill sTpath1
        
        Open sTpath2 For Binary As #2
            sTemp = Space$(LOF(2))
            Get #2, , sTemp
            Put #1, , sTemp
        Close #2
        Kill sTpath2
        
        Open sTpath3 For Binary As #2
            sTemp = Space$(LOF(2))
            Get #2, , sTemp
            Put #1, , sTemp
        Close #2
        Kill sTpath3
        
        Open sTpath4 For Binary As #2
            sTemp = Space$(LOF(2))
            Get #2, , sTemp
            Put #1, , sTemp
        Close #2
        Kill sTpath4
    Close #1

End Sub

Public Sub Set_Options()
'//set all options settings in the registry
    If bFirst Then Exit Sub
    With frmMain
        '//spelling options
        SaveSetting App.EXEName, "Dictionary", "optdic0", .optDictionary(0).Value
        SaveSetting App.EXEName, "Dictionary", "optdic1", .optDictionary(1).Value
        SaveSetting App.EXEName, "Dictionary", "optdic2", .optDictionary(2).Value
        SaveSetting App.EXEName, "Dictionary", "optcmp0", .optCompress(0).Value
        SaveSetting App.EXEName, "Dictionary", "optcmp1", .optCompress(1).Value
        SaveSetting App.EXEName, "Dictionary", "chkfilt", .chkFilter.Value
        SaveSetting App.EXEName, "Dictionary", "optsrt0", .optSort(0).Value
        SaveSetting App.EXEName, "Dictionary", "optsrt1", .optSort(1).Value
        SaveSetting App.EXEName, "Dictionary", "opttlr0", .optTolerance(0).Value
        SaveSetting App.EXEName, "Dictionary", "optsrt1", .optTolerance(1).Value
        SaveSetting App.EXEName, "Dictionary", "chkopt0", .chkOptions(0).Value
        SaveSetting App.EXEName, "Dictionary", "chkopt1", .chkOptions(1).Value
        SaveSetting App.EXEName, "Dictionary", "chkopt2", .chkOptions(2).Value
        SaveSetting App.EXEName, "Dictionary", "chkopt3", .chkOptions(3).Value
        SaveSetting App.EXEName, "Dictionary", "chkopt4", .chkOptions(4).Value
        SaveSetting App.EXEName, "Dictionary", "chkopt5", .chkOptions(5).Value
        SaveSetting App.EXEName, "Dictionary", "optcrt0", .optCorrectstyle(0).Value
        SaveSetting App.EXEName, "Dictionary", "optcrt1", .optCorrectstyle(1).Value
        SaveSetting App.EXEName, "Dictionary", "optcrt2", .optCorrectstyle(2).Value
        SaveSetting App.EXEName, "Dictionary", "optcrt3", .optCorrectstyle(3).Value
        '//dictionary options
        SaveSetting App.EXEName, "Dictionary", "chksrc0", .chkSearch(0).Value
        SaveSetting App.EXEName, "Dictionary", "chksrc1", .chkSearch(1).Value
        SaveSetting App.EXEName, "Dictionary", "chksrc2", .chkSearch(2).Value
        SaveSetting App.EXEName, "Dictionary", "chksrc3", .chkSearch(3).Value
        SaveSetting App.EXEName, "Dictionary", "optprg0", .optProgress(0).Value
        SaveSetting App.EXEName, "Dictionary", "optprg1", .optProgress(1).Value
        SaveSetting App.EXEName, "Dictionary", "chkfrt", .chkFormat.Value
        SaveSetting App.EXEName, "Dictionary", "chkttp", .chkTooltip.Value
    End With
    
End Sub

Public Sub Get_Options()
'//fetch control settings from last run
    With frmMain
        '//spell check values
        .optDictionary(0).Value = GetSetting(App.EXEName, "Dictionary", "optdic0", "False")
        .optDictionary(1).Value = GetSetting(App.EXEName, "Dictionary", "optdic1", "False")
        .optDictionary(2).Value = GetSetting(App.EXEName, "Dictionary", "optdic2", "False")
        .optCompress(0).Value = GetSetting(App.EXEName, "Dictionary", "optcmp0", "False")
        .optCompress(1).Value = GetSetting(App.EXEName, "Dictionary", "optcmp1", "False")
        .chkFilter.Value = GetSetting(App.EXEName, "Dictionary", "chkfilt", "0")
        .optSort(0).Value = GetSetting(App.EXEName, "Dictionary", "optsrt0", "False")
        .optSort(1).Value = GetSetting(App.EXEName, "Dictionary", "optsrt1", "False")
        .optTolerance(0).Value = GetSetting(App.EXEName, "Dictionary", "opttlr0", "False")
        .optTolerance(1).Value = GetSetting(App.EXEName, "Dictionary", "opttlr1", "False")
        .chkOptions(0).Value = GetSetting(App.EXEName, "Dictionary", "chkopt0", "0")
        .chkOptions(1).Value = GetSetting(App.EXEName, "Dictionary", "chkopt1", "0")
        .chkOptions(2).Value = GetSetting(App.EXEName, "Dictionary", "chkopt2", "0")
        .chkOptions(3).Value = GetSetting(App.EXEName, "Dictionary", "chkopt3", "0")
        .chkOptions(4).Value = GetSetting(App.EXEName, "Dictionary", "chkopt4", "0")
        .chkOptions(5).Value = GetSetting(App.EXEName, "Dictionary", "chkopt5", "0")
        .optCorrectstyle(0).Value = GetSetting(App.EXEName, "Dictionary", "optcrt0", "False")
        .optCorrectstyle(1).Value = GetSetting(App.EXEName, "Dictionary", "optcrt1", "False")
        .optCorrectstyle(2).Value = GetSetting(App.EXEName, "Dictionary", "optcrt2", "False")
        .optCorrectstyle(3).Value = GetSetting(App.EXEName, "Dictionary", "optcrt3", "False")
        '//dictionary values
        .chkSearch(0).Value = GetSetting(App.EXEName, "Dictionary", "chksrc0", "0")
        .chkSearch(1).Value = GetSetting(App.EXEName, "Dictionary", "chksrc1", "0")
        .chkSearch(2).Value = GetSetting(App.EXEName, "Dictionary", "chksrc2", "0")
        .chkSearch(3).Value = GetSetting(App.EXEName, "Dictionary", "chksrc3", "0")
        .optProgress(0).Value = GetSetting(App.EXEName, "Dictionary", "optprg0", "False")
        .optProgress(1).Value = GetSetting(App.EXEName, "Dictionary", "optprg1", "False")
        .chkFormat.Value = GetSetting(App.EXEName, "Dictionary", "chkfrt", "0")
        .chkTooltip.Value = GetSetting(App.EXEName, "Dictionary", "chkttp", "0")
        '//disable api compression options for non nt systems
        If bOVersion Then
            .optCompress(0).Enabled = False
            .optCompress(1).Enabled = False
        End If
    End With
        
End Sub

Private Function Get_DBase() As Integer
'//get spell check list based on registry setting
    With frmMain
        If GetSetting(App.EXEName, "Dictionary", "optdic0", "True") = "True" Then
            Get_DBase = 0
            Exit Function
        ElseIf GetSetting(App.EXEName, "Dictionary", "optdic1", "False") = "True" Then
            Get_DBase = 1
            Exit Function
        ElseIf GetSetting(App.EXEName, "Dictionary", "optdic2", "False") = "True" Then
            Get_DBase = 2
            Exit Function
        Else
            Get_DBase = 0
        End If
    End With
    
End Function

Public Sub Identify_OS()
'//os version check for compression choice
Dim rOsVersion As OSVersion

    rOsVersion.dwOSVersionInfoSize = Len(rOsVersion)
    If GetVersionEx(rOsVersion) Then
        If Not rOsVersion.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            bOVersion = True
        End If
    End If

End Sub

Public Function File_Exist(ByVal sFile As String) As Boolean
'//file check
    If Len(Dir(sFile)) > 0 Then
        File_Exist = True
    End If

End Function

Private Sub Split_Dbase()
'//used to split the dbase into smaller files (upload problems)
Dim sTpath1 As String
Dim sTpath2 As String
Dim sTpath3 As String
Dim sTpath4 As String
Dim sPath   As String
Dim sTemp   As String
Dim iPos    As Long
Dim iLen    As Long

    sTpath1 = App.Path & "\dic1.txt"
    sTpath2 = App.Path & "\dic2.txt"
    sTpath3 = App.Path & "\dic3.txt"
    sTpath4 = App.Path & "\dic4.txt"
    sPath = App.Path & "\sed.txt"
    
    Open sPath For Binary As #1
        sTemp = Space$(LOF(1))
        Get #1, , sTemp
    Close #1
    iLen = Len(sTemp)


    Open sPath For Binary As #1
        Open sTpath1 For Binary As #2
            sTemp = Space$(5000000)
            iPos = 5000000
            Get #1, , sTemp
            Put #2, , sTemp
        Close #2
        
        Open sTpath2 For Binary As #2
            sTemp = Space$(5000000)
            iPos = iPos + 5000000
            Get #1, , sTemp
            Put #2, , sTemp
        Close #2
        
        Open sTpath3 For Binary As #2
            sTemp = Space$(5000000)
            iPos = iPos + 5000000
            Get #1, , sTemp
            Put #2, , sTemp
        Close #2
        
        iPos = iLen - iPos
        
        Open sTpath4 For Binary As #2
            sTemp = Space$(iPos)
            Get #1, , sTemp
            Put #2, , sTemp
        Close #2
    Close #1
    
End Sub
