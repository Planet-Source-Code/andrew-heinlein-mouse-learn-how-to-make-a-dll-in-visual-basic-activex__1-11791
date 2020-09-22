VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MY_DLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This is just a simple class moduel.


'set the functions you want to use in linking program to PUBLIC:
Public Function ReturnUpperCase(ByVal LCaseString As String) As String
    ReturnUpperCase = UCase(LCaseString)
End Function

Public Function ReturnLowerCase(ByVal UcaseString As String) As String
    ReturnLowerCase = LCase(UcaseString)
End Function

'subs work the same way:
Public Sub AddNumbers(ByVal iX As Long, ByVal iY As Long, iReturn As Long)
    'since your passing in "iReturn" not By Value, you may alter it.
    iReturn = CLng(iX + iY)
End Sub

'Want to use a picture box or another object as a parameter?
'you cant use "as picturebox" since its a specific object
'but you can use this:

Public Sub SizePicBox(ByVal PictureBx As Object, ByVal Widt As Integer, ByVal Higt As Integer)
    PictureBx.Width = Widt
    PictureBx.Height = Higt
End Sub

'You can use forms, moduels, messageboxs... anything you want... its
'exatly like an Active X control.

Public Sub ShowAbout()
    Load Form1
    Form1.Show
End Sub

'People who use the dll wont see this function cause its PRIVATE:
Private Function HiddenFunction(ByVal Param1 As Long, ByVal Param2 As Long) As String
    HiddenFunction = CStr(Param1) & CStr(Param2)
End Function

'NOTE: When you *compile* this project, VB will automatically REGISTER the DLL!
'if you want to fix something you will have to UNREGISTER the old on first
'before compiling otherwise you will have a shitload of "your dll's" registered
'and only one to show for it.

'Remember... this is exatly like an active x control.  if you add it into a program
'if will have to be registered before the linked program will work.
