VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmStatistik 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Statistiken..."
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3630
   StartUpPosition =   1  'Fenstermitte
   Begin SHDocVwCtl.WebBrowser wbrSource 
      Height          =   345
      Left            =   1845
      TabIndex        =   7
      Top             =   1035
      Width           =   315
      ExtentX         =   556
      ExtentY         =   609
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtCurrentURL 
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1680
      Width           =   3315
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1620
      Top             =   300
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1020
      Width           =   1275
   End
   Begin VB.TextBox txtEmails 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtLinks 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtSites 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Current URL"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   1500
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Nr. of E-Mails found"
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nr. of Links in Cache"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   840
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Nr. of Sites searched"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frmStatistik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strURL As String
Public mcolFinished As New Collection
Public mcolCache As New Collection
Public Property Get vbct()
  vbct = Chr(34)
End Property


Private Sub cmdStop_Click()
  Close #1
  Open "C:\Statistic.txt" For Append As #1
  Print #1, "Number of Sites searched: " & txtSites
  Print #1, "Number of emails found: " & txtEmails
  Print #1, ""
  Close #1
  Unload Me
End Sub

Private Sub Form_Load()
  Timer1.Interval = 1
  mcolCache.Add strURL, strURL
  Open "C:\EMails.txt" For Append As #1
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Call WormTheWeb
  Timer1.Enabled = False
End Sub

Private Sub WormTheWeb()
  Dim vURL As Variant
  Dim strFile As String
  Dim i As Long
  
  On Error Resume Next
  
  Do While Not mcolCache.Count = 0
    vURL = mcolCache(1)
    strURL = vURL
    txtCurrentURL = strURL
    wbrSource.Navigate strURL
    Do While wbrSource.Busy
      DoEvents
    Loop
    
    If Err = 0 Then
      For i = 0 To wbrSource.Document.links.length - 1
        strURL = wbrSource.Document.links(i).href
        If InStr(1, "mailto:", strURL, vbTextCompare) > 0 Then
          Print #1, strURL
        Else
          If Not AllreadySearched(strURL) Then
            txtLinks = Val(txtLinks) + 1
            mcolCache.Add strURL
          End If
        End If
        DoEvents
      Next i
    Else
      Err.Clear
    End If
    Call mcolCache.Remove(1)
    strURL = vURL
    mcolFinished.Add strURL, strURL
    txtLinks = Val(txtLinks) - 1
    txtSites = Val(txtSites) + 1
    DoEvents
  Loop
End Sub

Private Function AllreadySearched(vURL As Variant) As Boolean
  
  On Error Resume Next
  
  mcolFinished.Add strURL, strURL
  If Err = 0 Then
    AllreadySearched = False
  Else
    AllreadySearched = True
    Err.Clear
  End If
  
End Function

