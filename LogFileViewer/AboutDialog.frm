VERSION 5.00
Begin VB.Form AboutDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "AboutDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OkButton 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   810
   End
   Begin VB.Label CopyrightLabel 
      AutoSize        =   -1  'True
      Caption         =   "[Legal Copyright]"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label VersionLabel 
      AutoSize        =   -1  'True
      Caption         =   "[Version]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   930
   End
   Begin VB.Label TitleLabel 
      AutoSize        =   -1  'True
      Caption         =   "[Application Title]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2145
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "AboutDialog.frx":000C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "AboutDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: Skip
Option Explicit

Private Sub Form_Load()

    On Error GoTo ERROR_HANDLER
    
    TitleLabel.Caption = App.Title
    VersionLabel.Caption = AppVersion
    CopyrightLabel.Caption = App.LegalCopyright
    
    Exit Sub
    
ERROR_HANDLER:
    ' ignore errors that occur when window gets too small
    Debug.Print "Form_Load: " & Err.Description
    Resume Next

End Sub

Private Sub OkButton_Click()

    Hide
    
End Sub
