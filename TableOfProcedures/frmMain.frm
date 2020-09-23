VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Create a Table of Procedures"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   3600
      Left            =   6450
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CommandButton cmdClipboard 
      Height          =   480
      Left            =   2415
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   30
      Width           =   1740
   End
   Begin VB.CommandButton cmdFile 
      Height          =   465
      Left            =   45
      Picture         =   "frmMain.frx":0A9B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   1740
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   30
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   615
      Width           =   4170
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************
'**                        Create a Table of Procedures
'**                               Version 1.0.0
'**                               By Ken Foster
'**                                 Nov  2004
'**                       Freeware--- no copyrights claimed
'*******************************************************************

'******************* Table of Procedures *************
'   Private Sub Form_Load
'   Private Sub cmdFile_Click
'   Private Sub cmdClipboard_Click
'   Public Function Trun_St
'   Private Sub Form_Resize
'   Private Sub Form_Unload
'******************* End of Table ********************

Const ar As Integer = 11 'number of last element in array
Dim arry(ar)

Private Sub Form_Load()
   arry(0) = "Private Sub"
   arry(1) = "Public Sub"
   arry(2) = "Sub"
   arry(3) = "Function"
   arry(4) = "Friend"
   arry(5) = "Static"
   arry(6) = "Property"
   arry(7) = "Private Function"
   arry(8) = "Public Function"
   arry(9) = "Public Property Get"
   arry(10) = "Public Property Let"
   arry(11) = "Public Property Set"
End Sub

Private Sub cmdFile_Click()
   Dim instring As String  'line loaded from file
   Dim sStg As String  'holds words from instring to compare
   Dim ShV As String  'short version of instring that is loaded into listbox
   Dim Y As Integer  'counter
   Dim LA As Integer  'length of word in array
   
   List1.Clear
   ShowOpen
   If cmndlg.FileName = "" Then Exit Sub
   On Error GoTo Clnup
   List1.AddItem "'***************** Table of Procedures *************"
   Open cmndlg.FileName For Input As #1  ' file opened for reading
   While Not EOF(1)
      Line Input #1, instring
      sStg = Mid$(instring, 1, 1)
      For Y = 0 To ar  'get words in array
         LA = Len(arry(Y)) 'find length of array word
         sStg = Mid$(instring, 1, LA)  'find a word
            If sStg = arry(Y) Then  ' if word found then add line to textbox
               ShV = Trun_St(instring)  'shortened version of instring
               List1.AddItem "'   " & ShV   ' add to listbox
            End If
      Next Y
   Wend
   Close #1
   List1.AddItem "'***************** End of Table ********************"
   
   If List1.ListCount = 2 Then  'No data loaded
      List1.Clear
      List1.AddItem "                     No Procedures Available"
   End If
   
Clnup:
   Close #1
   
End Sub

Private Sub cmdClipboard_Click()
Dim x As Integer
'copy list1 to text1
'makes it easier to paste to clipboard
  Text1.Text = ""
  For x = 0 To List1.ListCount - 1
      List1.ListIndex = x
      Text1.Text = Text1.Text & List1.Text & vbCrLf
  Next x
   Clipboard.Clear
   Clipboard.SetText Text1
End Sub

Public Function Trun_St(txt As String) As String
   Dim x As Integer
   Dim Chg As String
   
   For x = 1 To Len(txt)
      Chg = Mid$(txt, x, 1)
      If Chg = "(" Then  ' if true then stop
         Trun_St = Left$(txt, x - 1)
         Exit Function
      End If
   Next x
End Function

Private Sub Form_Resize()
List1.Width = frmMain.Width - 200
List1.Top = 615
List1.Height = frmMain.Height - 1100
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
