VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Select All Web Broswer"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "G.O."
      Height          =   315
      Left            =   11100
      TabIndex        =   2
      Top             =   60
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "http://www.glsetup.com/readme.htm"
      Top             =   60
      Width           =   11055
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7995
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   11745
      ExtentX         =   20717
      ExtentY         =   14102
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rafael Bonventi
'using the menu mouse left click when using the web browser control with vb.


Private Sub Command1_Click()

    WebBrowser1.Navigate Text1.Text
    
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    WebBrowser1.SetFocus
    If WebBrowser1.LocationURL <> Text1.Text Then
        Exit Sub
    Else
        Call WebBrowser1.ExecWB(OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT)
    End If
    
End Sub

