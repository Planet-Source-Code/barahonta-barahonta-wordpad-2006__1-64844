VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help Topics"
   ClientHeight    =   4920
   ClientLeft      =   4065
   ClientTop       =   4020
   ClientWidth     =   6990
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Form2.frx":0000
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   4470
      Left            =   120
      Picture         =   "Form2.frx":0189
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
