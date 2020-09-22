VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Barahonta WordPad 2006  [Version 1.0.2]"
   ClientHeight    =   8070
   ClientLeft      =   1920
   ClientTop       =   1890
   ClientWidth     =   10575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10575
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   2040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   1200
      TabIndex        =   24
      Top             =   2760
      Width           =   7935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Barahonta WordPad 2006"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Image Image1 
         Height          =   4395
         Left            =   0
         Picture         =   "Form1.frx":08CA
         Top             =   0
         Width           =   7875
      End
   End
   Begin VB.CommandButton Command23 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8760
      Picture         =   "Form1.frx":96EE
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Press To Open The About Window"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      Picture         =   "Form1.frx":9FB8
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Press To Open Help Topics Window"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Bullets"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Picture         =   "Form1.frx":A882
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Press To Turn Bullets On/Off"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Font +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      Picture         =   "Form1.frx":B14C
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Press To + Font Size"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Font -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      Picture         =   "Form1.frx":BA16
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Press To - Font Size"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      Picture         =   "Form1.frx":C2E0
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Press To Select All Text In The File"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Redo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Picture         =   "Form1.frx":CBAA
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Press To Redo Last Action"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Undo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      Picture         =   "Form1.frx":D474
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Press To Undo Last Action"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Font Color"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      Picture         =   "Form1.frx":DD3E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Press To Change The Font Color"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Font Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "Form1.frx":E608
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Press To Change The Font Type"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "====>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      Picture         =   "Form1.frx":EED2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Press To Change Text Direction To Right"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "====="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      Picture         =   "Form1.frx":F79C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Press To Change Text Direction To Center"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<===="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      Picture         =   "Form1.frx":10066
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Press To Change Text Direction To Left"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Underline"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Picture         =   "Form1.frx":10930
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Press To Underline Selected Font"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Italic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      Picture         =   "Form1.frx":111FA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Press To Italic Selected Font"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      Picture         =   "Form1.frx":11AC4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Press To Bold Selected String"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Picture         =   "Form1.frx":1238E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Press To Save Your File"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Picture         =   "Form1.frx":12C58
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Press To Open A Saved Document"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "Form1.frx":13522
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Press To Open A Blank Document"
      Top             =   120
      Width           =   855
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10610
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":13DEC
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      Picture         =   "Form1.frx":13E7C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   6600
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Paste"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      Picture         =   "Form1.frx":14746
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Press To Paste Copied/Cuted Strings"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Cut"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Picture         =   "Form1.frx":15010
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Press To Cut Selected String"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      Picture         =   "Form1.frx":158DA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Press To Copy Selected String"
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnunew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopn 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuundo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuredo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuall 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "Format"
      Begin VB.Menu mnufontt 
         Caption         =   "Font Type"
      End
      Begin VB.Menu mnufntcolor 
         Caption         =   "Font Color"
      End
   End
   Begin VB.Menu mnuh 
      Caption         =   "Help"
      Begin VB.Menu mnuhelppp 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu mnabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const maxUndo = 50
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(maxUndo) As String
Dim stackBK(maxUndo) As String
Dim i As Integer
Private Sub Command1_Click()
Dim ans As Single
ans = MsgBox("Do you want to save the changes made in this file?", vbYesNoCancel + vbQuestion, "Save File")
If ans = vbYes Then
Dialog.Filter = "Rich Text Document(.RTF)|*.rtf"
Dialog.ShowSave
Text1.SaveFile Dialog.FileName
Text1.Text = ""
Text1.SetFocus
End If
If ans = vbNo Then
Text1.Text = ""
Text1.SetFocus
End If
If ans = vbCancel Then
Text1.SetFocus
End If
End Sub

Private Sub Command10_Click()
On Error Resume Next
Dialog.ShowColor
Text1.SelColor = Dialog.Color
Text1.SetFocus
End Sub

Private Sub Command11_Click()
With Dialog
On Error Resume Next
.CancelError = True
.FontName = "Fonts"
.Flags = cdlCFEffects Or cdlCFBoth
.ShowFont
Text1.SelFontName = .FontName
Text1.SelBold = .FontBold
Text1.SelItalic = .FontItalic
Text1.SelFontSize = .FontSize
Text1.SelStrikeThru = .FontStrikethru
Text1.SelUnderline = .FontUnderline
Text1.SelColor = .Color
Text1.SetFocus
End With
End Sub

Private Sub Command12_Click()
Dim ans As Single
ans = MsgBox("Do you want to save the changes made in this file?", vbYesNoCancel + vbQuestion, "Save File")
If ans = vbYes Then
Dialog.Filter = "Rich Text Document(.RTF)|*.rtf"
Dialog.ShowSave
Text1.SaveFile Dialog.FileName
End
End If
If ans = vbNo Then
End
End If
If ans = vbCancel Then
Text1.SetFocus
End If
End Sub

Private Sub Command13_Click()
    Text1.SelText = Clipboard.GetText
    Text1.SetFocus
End Sub

Private Sub Command14_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
    Text1.SetFocus
End Sub

Private Sub Command15_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SetFocus
End Sub

Private Sub Command16_Click()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
End Sub

Private Sub Command17_Click()
    If gintIndex < maxUndo Then
        gblnIgnoreChange = True
        gintIndex = gintIndex + 1
        On Error Resume Next
        Text1.TextRTF = gstrStack(gintIndex)
        gblnIgnoreChange = False
    End If
End Sub

Private Sub Command18_Click()
    If gintIndex = 0 Then Exit Sub
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    Text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub Command19_Click()
On Error Resume Next
Text1.SelFontSize = Text1.SelFontSize - 4
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Dialog.Filter = "All Supported Formats|*.txt;*.rtf"
Dialog.ShowOpen
Text1.LoadFile Dialog.FileName
Text1.SetFocus
End Sub

Private Sub Command20_Click()
On Error Resume Next
Text1.SelFontSize = Text1.SelFontSize + 4
Text1.SetFocus
End Sub

Private Sub Command21_Click()
Text1.SelBullet = Not Text1.SelBullet
Text1.SetFocus
End Sub

Private Sub Command22_Click()
Form2.Show vbModal
End Sub

Private Sub Command24_Click()

End Sub

Private Sub Command23_Click()
Form3.Show vbModal
End Sub

Private Sub Command3_Click()
Dialog.Filter = "Rich Text Document(.RTF)|*.rtf"
Dialog.ShowSave
Text1.SaveFile Dialog.FileName
End Sub

Private Sub Command4_Click()
Text1.SelBold = Not Text1.SelBold
Text1.SetFocus
End Sub

Private Sub Command5_Click()
Text1.SelItalic = Not Text1.SelItalic
Text1.SetFocus
End Sub

Private Sub Command6_Click()
Text1.SelUnderline = Not Text1.SelUnderline
Text1.SetFocus
End Sub

Private Sub Command7_Click()
Text1.SelAlignment = 0
Text1.SetFocus
End Sub

Private Sub Command8_Click()
Text1.SelAlignment = 2
Text1.SetFocus
End Sub

Private Sub Command9_Click()
Text1.SelAlignment = 1
Text1.SetFocus
End Sub

Private Sub d_Click()

End Sub

Private Sub mnufrmat_Click()

End Sub

Private Sub mnufind_Click()
Form2.Show
End Sub

Private Sub Form_Resize()
Text1.Width = Form1.Width - 110
Text1.Height = Form1.Height - 2850
Command12.Left = Form1.Width - 650
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub Image1_Click()
Image1.Visible = False
Frame1.Visible = False
Label1.Visible = False
End Sub

Private Sub mnabout_Click()
Form3.Show vbModal
End Sub

Private Sub mnuall_Click()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
End Sub

Private Sub mnucopy_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SetFocus
End Sub

Private Sub mnucut_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
    Text1.SetFocus
End Sub

Private Sub mnuexit_Click()
Dim ans As Single
ans = MsgBox("Do you want to save the changes made in this file?", vbYesNoCancel + vbQuestion, "Save File")
If ans = vbYes Then
Dialog.Filter = "Rich Text Document(.RTF)|*.rtf"
Dialog.ShowSave
Text1.SaveFile Dialog.FileName
End
End If
If ans = vbNo Then
End
End If
If ans = vbCancel Then
Text1.SetFocus
End If
End Sub

Private Sub mnufntcolor_Click()
Dialog.ShowColor
Text1.SelColor = Dialog.Color
Text1.SetFocus
End Sub

Private Sub mnufontt_Click()
With Dialog
On Error Resume Next
.CancelError = True
.FontName = "Fonts"
.Flags = cdlCFEffects Or cdlCFBoth
.ShowFont
Text1.SelFontName = .FontName
Text1.SelBold = .FontBold
Text1.SelItalic = .FontItalic
Text1.SelFontSize = .FontSize
Text1.SelStrikeThru = .FontStrikethru
Text1.SelUnderline = .FontUnderline
Text1.SelColor = .Color
End With
End Sub

Private Sub mnuhelppp_Click()
Form2.Show vbModal
End Sub

Private Sub mnunew_Click()
Dim ans As Single
ans = MsgBox("Do you want to save the changes made in this file?", vbYesNoCancel + vbQuestion, "Save File")
If ans = vbYes Then
Dialog.Filter = "Rich Text Document(.RTF)|*.rtf"
Dialog.ShowSave
Text1.SaveFile Dialog.FileName
Text1.Text = ""
Text1.SetFocus
End If
If ans = vbNo Then
Text1.Text = ""
Text1.SetFocus
End If
If ans = vbCancel Then
Text1.SetFocus
End If
End Sub

Private Sub mnuopn_Click()
Dialog.Filter = "All Supported Formats|*.txt;*.rtf"
Dialog.ShowOpen
Text1.LoadFile Dialog.FileName
Text1.SetFocus
End Sub

Private Sub mnupaste_Click()
    Text1.SelText = Clipboard.GetText
    Text1.SetFocus
End Sub

Private Sub mnuredo_Click()
    If gintIndex < maxUndo Then
        gblnIgnoreChange = True
        gintIndex = gintIndex + 1
        On Error Resume Next
        Text1.TextRTF = gstrStack(gintIndex)
        gblnIgnoreChange = False
            End If
End Sub

Private Sub mnusave_Click()
Dialog.Filter = "Rich Text Document(.RTF)|*.rtf"
Dialog.ShowSave
Text1.SaveFile Dialog.FileName
End Sub

Private Sub mnuundo_Click()
    If gintIndex = 0 Then Exit Sub
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    Text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
    End Sub

Private Sub Text1_Change()
Dim g As Integer
Dim b As Integer
Dim i As Integer
g = maxUndo
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
                If gintIndex >= maxUndo + 1 Then
                    For b = 0 To maxUndo
                stackBK(b) = gstrStack(b)
            Next b
                        For i = 0 To maxUndo
                If g >= 1 Then
                g = g - 1
                gstrStack(g) = stackBK(g + 1)
                End If
            Next i
                        gintIndex = maxUndo
                    End If
        gstrStack(gintIndex) = Text1.TextRTF
    End If
End Sub

Private Sub Timer1_Timer()
Image1.Visible = False
Frame1.Visible = False
Label1.Visible = False
End Sub
