VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDictionary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "English Dictionary"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
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
   ScaleHeight     =   4800
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdControl 
      Caption         =   "Search"
      Height          =   345
      Index           =   1
      Left            =   4800
      TabIndex        =   3
      Top             =   210
      Width           =   885
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   4545
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Ok"
      Height          =   345
      Index           =   0
      Left            =   4800
      TabIndex        =   1
      Top             =   4350
      Width           =   885
   End
   Begin RichTextLib.RichTextBox txtDictionary 
      Height          =   3465
      Left            =   90
      TabIndex        =   0
      Top             =   720
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   6112
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDictionary.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdControl_Click(Index As Integer)
'//control array calls
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Get_Results
    End Select

End Sub

Private Sub Get_Results()
'//formats result and passes to collection lookup
'//seems like a long routine.. but lookups are so fast
'//you don't even notice a lag
Dim sText    As String
Dim sSearch  As String
Dim sResult  As String
Dim sContent As String

    sText = Trim$(txtSearch.Text)
    If Not Len(sText) = 0 Then
        txtDictionary.Text = vbNullString
        With frmMain
        '//first search - any
        sSearch = sText & " ()"
        sResult = .cDictionary.Dictionary_Fetch(sSearch)

        If Not sResult = "WNF" Then
            If .chkFormat.Value = 1 Then
                Format_Results sSearch & "  " & sResult
            Else
                txtDictionary.Text = sSearch & "  " & sResult
            End If
        End If

        '//second search
        sSearch = sText & " (1)"
        sResult = .cDictionary.Dictionary_Fetch(sSearch)

        If Not sResult = "WNF" Then
            If .chkFormat.Value = 1 Then
                Format_Results sSearch & "  " & sResult
            Else
                If Len(txtDictionary.Text) > 0 Then
                    sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                Else
                    sContent = sText & "  " & sResult
                End If
                txtDictionary.Text = sContent
            End If
        End If

        '//third search
        sSearch = sText & " (2)"
        sResult = .cDictionary.Dictionary_Fetch(sSearch)

        If Not sResult = "WNF" Then
            If .chkFormat.Value = 1 Then
                Format_Results sSearch & "  " & sResult
            Else
                If Len(txtDictionary.Text) > 0 Then
                    sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                Else
                    sContent = sText & "  " & sResult
                End If
                txtDictionary.Text = sContent
            End If
        End If

        If .chkSearch(0).Value = 1 Then
            '//first search - noun
            sSearch = sText & " (n.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If
        End If

        If .chkSearch(1).Value = 1 Then
            '//first search - verb
            sSearch = sText & " (v.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If
        End If

        If .chkSearch(2).Value = 1 Then
            '//first search - adjective
            sSearch = sText & " (a.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If
        End If

        '//any additional possible references
        If .chkSearch(3).Value = 1 Then
            sSearch = sText & " (adv.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If

            sSearch = sText & " (prep.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If

            sSearch = sText & " (pl.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If

            sSearch = sText & " (imp. & p. p.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If

            sSearch = sText & " (p. pr. & vb. n.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If

            sSearch = sText & " (v. i.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If

            sSearch = sText & " (v. t.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If

            sSearch = sText & " (adj.)"
            sResult = .cDictionary.Dictionary_Fetch(sSearch)

            If Not sResult = "WNF" Then
                If .chkFormat.Value = 1 Then
                    Format_Results sSearch & "  " & sResult
                Else
                    If Len(txtDictionary.Text) > 0 Then
                        sContent = txtDictionary.Text & vbNewLine & vbNewLine & sSearch & "  " & sResult
                    Else
                        sContent = sText & "  " & sResult
                    End If
                    txtDictionary.Text = sContent
                End If
            End If
        End If
        End With
    End If

        If Len(txtDictionary.Text) = 0 Then
            txtDictionary.Text = "Word Not Found"
        End If
    
End Sub

Private Sub Format_Results(ByVal sResult As String)
'//extracts word from current cursor position
Dim sTemp As String
Dim iPos  As Integer
Dim i     As Integer

    i = Len(txtDictionary.Text)
    iPos = InStr(1, sResult, Chr$(41))

    With txtDictionary
        sTemp = .Text
        If Len(sTemp) > 0 Then
            .SelStart = (i + 1)
            .SelText = vbNewLine & vbNewLine & sResult
            .SelStart = (i + 4)
            .SelLength = iPos
            '.SelColor = vbBlue
            .SelBold = True
        Else
            .SelText = sResult
            .SelStart = 0
            .SelLength = iPos
            '.SelColor = vbBlue
            .SelBold = True
        End If

        .SelLength = 0
        .SelStart = 0
        .SelColor = &H0
        .SelBold = False
    End With

End Sub
