VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About ..."
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line1 
      BorderColor     =   &H00C000C0&
      X1              =   75
      X2              =   5700
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   75
      Picture         =   "frmAbout.frx":0CCA
      Stretch         =   -1  'True
      Top             =   75
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   450
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "or Email: manleviet@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   225
      TabIndex        =   4
      Top             =   1950
      Width           =   3375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Class IT K24B - Science College - Hue University"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   525
      TabIndex        =   3
      Top             =   1650
      Width           =   5130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Contact with author by: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   150
      TabIndex        =   2
      Top             =   1350
      Width           =   2490
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Created by Le Viet Man"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   150
      TabIndex        =   1
      Top             =   1050
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Huffman Coding Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1200
      TabIndex        =   0
      Top             =   75
      Width           =   3540
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label6.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
