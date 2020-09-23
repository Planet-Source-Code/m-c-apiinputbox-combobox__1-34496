VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show me the trick"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "AUTOTYPE ABILITY ADDED !"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ver 1.0
'first boguous attempt
'ver 2.0 (excelency acchived)
' in ver 1.0 nothing could be typed in text portion of combo box
' this is now ok
' in ver 1.0 wierd thing happend to Num lock, caps lock on keyboard
' this is now ok
' added ability to swallow selection on ENTER key pressed
' horizontal and vertical scroll added as and only if needed
' autotype ability added
' this last one vas especialy hard to figure out as API combobox doesn't
' get WM_CHAR message





Private Sub Command1_Click()
Dim tt() As Variant
tt = Array("alfa", "beta", "gama", "delta", "altruit", "omegaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")
CreateComboBoxInputBox "API InputBox & ComboBox ver. 2.0", "Instead of typing, now user can select from combo box. Try to type in word alfa", tt

End Sub

Private Sub Command2_Click()
MsgBox "1. '?' button glues as it should, attempting to get events from it crashed app." & Chr(10) & _
"    I am interested in any solutions to this. " & Chr(10) & _
"2. more comments are in declaratin section module" & Chr(10) & _
"3. If you realy like this and you know the stuff then tell me:" & Chr(10) & _
"3.1.  How to change APIInputBox & ComboBox main window icon - how to get it from resource" & Chr(10) & _
"3.2.  How to put picture into API created window (allso from resource), I read somewhere that it can be puted into EDIT class ?" & Chr(10) & _
"" & Chr(10) & _
"Thanks for any ereplay to: kozlicki@yahoo.com"


End Sub

