VERSION 5.00
Begin VB.Form frmWorm 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Wormloch..."
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   2550
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   840
      Width           =   915
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Crawl"
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   840
      Width           =   1275
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Text            =   "http://"
      Top             =   420
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "Start searching at:"
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2475
   End
End
Attribute VB_Name = "frmWorm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnd_Click()
  Unload Me
End Sub

Private Sub cmdStart_Click()
  Dim frmStat As New frmStatistik
  frmStat.strURL = txtURL
  frmStat.Show vbModal
End Sub
