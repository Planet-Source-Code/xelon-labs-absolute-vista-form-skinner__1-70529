VERSION 5.00
Object = "*\AGlasskinner.vbp"
Begin VB.Form FrmTest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Glass.Glasskinner Glasskinner1 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FrmTest.frx":0000
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmTest.frx":0018
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
