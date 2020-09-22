VERSION 5.00
Begin VB.Form mainFrm 
   Caption         =   "The linking APP"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   3000
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "mainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'you must add the DLL as a REFERENCE before you can use it.
' todo that you click the PROJECT menu item
' then you click on  "References"
' browse for the DLL

Private sMY_DLL As New MY_DLL

Private Sub Command1_Click()
    sMY_DLL.ShowAbout
End Sub

Private Sub Form_Load()
    Dim Returned As Long
    With List1
        .AddItem sMY_DLL.ReturnLowerCase("MY FIRST DLL")
        .AddItem sMY_DLL.ReturnUpperCase("my first dll")
        sMY_DLL.AddNumbers 34, 43, Returned
        .AddItem Returned
    End With
    sMY_DLL.SizePicBox Picture1, 1500, 1500
End Sub
