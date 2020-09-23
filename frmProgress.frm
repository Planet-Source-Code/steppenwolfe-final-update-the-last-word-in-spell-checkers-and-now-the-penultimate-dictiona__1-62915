VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
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
   ScaleHeight     =   1125
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      Begin MSComctlLib.ProgressBar pBar1 
         Height          =   165
         Left            =   120
         TabIndex        =   1
         Top             =   510
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblLoading 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Dictionary File.."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   1995
      End
      Begin VB.Label lblLoading 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loaded: 0 %"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   330
         Width           =   930
      End
      Begin VB.Label lblLoading 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Dictionary File.."
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   750
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


