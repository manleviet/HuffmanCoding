VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHuffman 
   Caption         =   "Huffman Coding Program"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   ForeColor       =   &H00000000&
   Icon            =   "frmHuffman.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   1125
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   840
      Left            =   4950
      TabIndex        =   40
      Top             =   5400
      Width           =   6390
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4275
         TabIndex        =   142
         Top             =   225
         Width           =   1740
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2250
         TabIndex        =   42
         Top             =   225
         Width           =   1665
      End
      Begin VB.CommandButton cmdSaveHuffman 
         Caption         =   "Save to file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   225
         TabIndex        =   41
         Top             =   225
         Width           =   1665
      End
   End
   Begin VB.Frame Frame3 
      Height          =   840
      Left            =   0
      TabIndex        =   35
      Top             =   5400
      Width           =   4815
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3150
         TabIndex        =   38
         Top             =   225
         Width           =   1515
      End
      Begin VB.CommandButton cmdSavep 
         Caption         =   "Save to file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1725
         TabIndex        =   37
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load from file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   75
         TabIndex        =   36
         Top             =   225
         Width           =   1590
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "M· Huffman ®­îc lËp:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   4890
      Left            =   4950
      TabIndex        =   39
      Top             =   525
      Width           =   6390
      Begin VB.Line Line5 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   5
         X1              =   6150
         X2              =   150
         Y1              =   4425
         Y2              =   4425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "L(a):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   66
         Left            =   2025
         TabIndex        =   144
         Top             =   4500
         Width           =   570
      End
      Begin VB.Label lblhuffman1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2625
         TabIndex        =   143
         Top             =   4500
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   32
         Left            =   4800
         TabIndex        =   141
         Top             =   4050
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   31
         Left            =   4800
         TabIndex        =   140
         Top             =   3675
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   30
         Left            =   4800
         TabIndex        =   139
         Top             =   3300
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   29
         Left            =   4800
         TabIndex        =   138
         Top             =   2925
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   28
         Left            =   4800
         TabIndex        =   137
         Top             =   2550
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   27
         Left            =   4800
         TabIndex        =   136
         Top             =   2175
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   26
         Left            =   4800
         TabIndex        =   135
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   25
         Left            =   4800
         TabIndex        =   134
         Top             =   1425
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   24
         Left            =   4800
         TabIndex        =   133
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   23
         Left            =   4800
         TabIndex        =   132
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   22
         Left            =   4800
         TabIndex        =   131
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   21
         Left            =   2625
         TabIndex        =   130
         Top             =   4050
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   20
         Left            =   2625
         TabIndex        =   129
         Top             =   3675
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   19
         Left            =   2625
         TabIndex        =   128
         Top             =   3300
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   18
         Left            =   2625
         TabIndex        =   127
         Top             =   2925
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   17
         Left            =   2625
         TabIndex        =   126
         Top             =   2550
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   16
         Left            =   2625
         TabIndex        =   125
         Top             =   2175
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   15
         Left            =   2625
         TabIndex        =   124
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   14
         Left            =   2625
         TabIndex        =   123
         Top             =   1425
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   13
         Left            =   2625
         TabIndex        =   122
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   12
         Left            =   2625
         TabIndex        =   121
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   11
         Left            =   2625
         TabIndex        =   120
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   10
         Left            =   450
         TabIndex        =   119
         Top             =   4050
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   9
         Left            =   450
         TabIndex        =   118
         Top             =   3675
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   8
         Left            =   450
         TabIndex        =   117
         Top             =   3300
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   7
         Left            =   450
         TabIndex        =   116
         Top             =   2925
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   6
         Left            =   450
         TabIndex        =   115
         Top             =   2550
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   5
         Left            =   450
         TabIndex        =   114
         Top             =   2175
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   450
         TabIndex        =   113
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   450
         TabIndex        =   112
         Top             =   1425
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   450
         TabIndex        =   111
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   450
         TabIndex        =   110
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label lblhuffman 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   450
         TabIndex        =   43
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   65
         Left            =   150
         TabIndex        =   109
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "f:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   64
         Left            =   150
         TabIndex        =   108
         Top             =   3675
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ª:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   63
         Left            =   150
         TabIndex        =   107
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "e:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   62
         Left            =   150
         TabIndex        =   106
         Top             =   2925
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "®:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   61
         Left            =   150
         TabIndex        =   105
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "d:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   60
         Left            =   150
         TabIndex        =   104
         Top             =   2175
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "c:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   59
         Left            =   150
         TabIndex        =   103
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¨:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   58
         Left            =   150
         TabIndex        =   102
         Top             =   675
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "©:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   57
         Left            =   150
         TabIndex        =   101
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "b:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   56
         Left            =   150
         TabIndex        =   100
         Top             =   1425
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "g:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   55
         Left            =   150
         TabIndex        =   99
         Top             =   4050
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "h:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   54
         Left            =   2325
         TabIndex        =   98
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   53
         Left            =   2325
         TabIndex        =   97
         Top             =   3675
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¬:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   52
         Left            =   2325
         TabIndex        =   96
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "o:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   51
         Left            =   2325
         TabIndex        =   95
         Top             =   2925
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "n:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   50
         Left            =   2325
         TabIndex        =   94
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "m:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   49
         Left            =   2325
         TabIndex        =   93
         Top             =   2175
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "l:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   48
         Left            =   2325
         TabIndex        =   92
         Top             =   1800
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "i:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   47
         Left            =   2325
         TabIndex        =   91
         Top             =   675
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "j:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   46
         Left            =   2325
         TabIndex        =   90
         Top             =   1050
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "k:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   45
         Left            =   2325
         TabIndex        =   89
         Top             =   1425
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "p:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   44
         Left            =   2325
         TabIndex        =   88
         Top             =   4050
         Width           =   240
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   5
         X1              =   2100
         X2              =   2100
         Y1              =   300
         Y2              =   4350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "q:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   43
         Left            =   4500
         TabIndex        =   87
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "w:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   42
         Left            =   4500
         TabIndex        =   86
         Top             =   3675
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   41
         Left            =   4500
         TabIndex        =   85
         Top             =   3300
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "x:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   40
         Left            =   4500
         TabIndex        =   84
         Top             =   2925
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "v:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   39
         Left            =   4500
         TabIndex        =   83
         Top             =   2550
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "­:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   38
         Left            =   4500
         TabIndex        =   82
         Top             =   2175
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "u:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   37
         Left            =   4500
         TabIndex        =   81
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "r:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   36
         Left            =   4500
         TabIndex        =   80
         Top             =   675
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "s:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   35
         Left            =   4500
         TabIndex        =   79
         Top             =   1050
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "t:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   34
         Left            =   4500
         TabIndex        =   78
         Top             =   1425
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "z:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Index           =   33
         Left            =   4500
         TabIndex        =   77
         Top             =   4050
         Width           =   225
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   5
         X1              =   4275
         X2              =   4275
         Y1              =   300
         Y2              =   4350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "X¸c suÊt c¸c ch÷ c¸i xuÊt hiÖn:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4890
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   525
      Width           =   4815
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   32
         Left            =   3750
         TabIndex        =   34
         Top             =   4050
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   31
         Left            =   3750
         TabIndex        =   33
         Top             =   3675
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   30
         Left            =   3750
         TabIndex        =   32
         Top             =   3300
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   29
         Left            =   3750
         TabIndex        =   31
         Top             =   2925
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   22
         Left            =   3750
         TabIndex        =   24
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   28
         Left            =   3750
         TabIndex        =   30
         Top             =   2550
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   27
         Left            =   3750
         TabIndex        =   29
         Top             =   2175
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   26
         Left            =   3750
         TabIndex        =   28
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   25
         Left            =   3750
         TabIndex        =   27
         Top             =   1425
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   24
         Left            =   3750
         TabIndex        =   26
         Top             =   1050
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   23
         Left            =   3750
         TabIndex        =   25
         Top             =   675
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   21
         Left            =   2100
         TabIndex        =   23
         Top             =   4050
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   20
         Left            =   2100
         TabIndex        =   22
         Top             =   3675
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   19
         Left            =   2100
         TabIndex        =   21
         Top             =   3300
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   18
         Left            =   2100
         TabIndex        =   20
         Top             =   2925
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   11
         Left            =   2100
         TabIndex        =   13
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   17
         Left            =   2100
         TabIndex        =   19
         Top             =   2550
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   16
         Left            =   2100
         TabIndex        =   18
         Top             =   2175
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   15
         Left            =   2100
         TabIndex        =   17
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   14
         Left            =   2100
         TabIndex        =   16
         Top             =   1425
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   13
         Left            =   2100
         TabIndex        =   15
         Top             =   1050
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   12
         Left            =   2100
         TabIndex        =   14
         Top             =   675
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   10
         Left            =   450
         TabIndex        =   12
         Top             =   4050
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   9
         Left            =   450
         TabIndex        =   11
         Top             =   3675
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   8
         Left            =   450
         TabIndex        =   10
         Top             =   3300
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   7
         Left            =   450
         TabIndex        =   9
         Top             =   2925
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   0
         Left            =   450
         TabIndex        =   2
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   6
         Left            =   450
         TabIndex        =   8
         Top             =   2550
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   5
         Left            =   450
         TabIndex        =   7
         Top             =   2175
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   4
         Left            =   450
         TabIndex        =   6
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   3
         Left            =   450
         TabIndex        =   5
         Top             =   1425
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   2
         Left            =   450
         TabIndex        =   4
         Top             =   1050
         Width           =   915
      End
      Begin VB.TextBox txtp 
         Height          =   315
         Index           =   1
         Left            =   450
         TabIndex        =   3
         Top             =   675
         Width           =   915
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   5
         X1              =   3225
         X2              =   3225
         Y1              =   300
         Y2              =   4350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "z:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   32
         Left            =   3450
         TabIndex        =   76
         Top             =   4050
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "t:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   31
         Left            =   3450
         TabIndex        =   75
         Top             =   1425
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "s:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   30
         Left            =   3450
         TabIndex        =   74
         Top             =   1050
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "r:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   29
         Left            =   3450
         TabIndex        =   73
         Top             =   675
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "u:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   28
         Left            =   3450
         TabIndex        =   72
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "­:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   27
         Left            =   3450
         TabIndex        =   71
         Top             =   2175
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "v:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   26
         Left            =   3450
         TabIndex        =   70
         Top             =   2550
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "x:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   25
         Left            =   3450
         TabIndex        =   69
         Top             =   2925
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   24
         Left            =   3450
         TabIndex        =   68
         Top             =   3300
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "w:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   23
         Left            =   3450
         TabIndex        =   67
         Top             =   3675
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "q:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   22
         Left            =   3450
         TabIndex        =   66
         Top             =   300
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   5
         X1              =   1575
         X2              =   1575
         Y1              =   300
         Y2              =   4350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "p:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   21
         Left            =   1800
         TabIndex        =   65
         Top             =   4050
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "k:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   20
         Left            =   1800
         TabIndex        =   64
         Top             =   1425
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "j:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   19
         Left            =   1800
         TabIndex        =   63
         Top             =   1050
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "i:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   18
         Left            =   1800
         TabIndex        =   62
         Top             =   675
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "l:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   17
         Left            =   1800
         TabIndex        =   61
         Top             =   1800
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "m:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   16
         Left            =   1800
         TabIndex        =   60
         Top             =   2175
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "n:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   15
         Left            =   1800
         TabIndex        =   59
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "o:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   14
         Left            =   1800
         TabIndex        =   58
         Top             =   2925
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¬:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   13
         Left            =   1800
         TabIndex        =   57
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   12
         Left            =   1800
         TabIndex        =   56
         Top             =   3675
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "h:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   11
         Left            =   1800
         TabIndex        =   55
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "g:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   10
         Left            =   150
         TabIndex        =   54
         Top             =   4050
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "b:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   150
         TabIndex        =   53
         Top             =   1425
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "©:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   150
         TabIndex        =   52
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¨:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   150
         TabIndex        =   51
         Top             =   675
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "c:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   4
         Left            =   150
         TabIndex        =   50
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "d:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   150
         TabIndex        =   49
         Top             =   2175
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "®:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   6
         Left            =   150
         TabIndex        =   48
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "e:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   7
         Left            =   150
         TabIndex        =   47
         Top             =   2925
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ª:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   8
         Left            =   150
         TabIndex        =   46
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "f:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   9
         Left            =   150
         TabIndex        =   45
         Top             =   3675
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   150
         TabIndex        =   44
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Huffman Coding Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   435
      Left            =   3375
      TabIndex        =   0
      Top             =   0
      Width           =   4410
   End
End
Attribute VB_Name = "frmHuffman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Le Viet Man - TinK24B - DHKH Hue
'Email: manleviet@yahoo.com
Option Explicit

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdCreate_Click()
    Call ClearH(Huffman)
    Call ReadInfor(Huffman)
    Call FixSource(Huffman)
    Call HCoding(Huffman)
    Call WriteInfor(Huffman)
End Sub

Private Sub cmdLoad_Click()
On Error GoTo DialogError
   With cdgOpen
      .CancelError = True  ' Generate Error number cdlCancel if user click Cancel
      .InitDir = App.Path    ' Initial (i.e. default ) Folder
      .Filter = "Input File (*.INP) | *.inp"
      .FilterIndex = 1  ' Select ""Executables (*.exe) | *.exe" as default
      .DialogTitle = "Select a input file to get information"
      .ShowOpen   ' Lauch the Open Dialog
      Call LoadInfor(cdgOpen)
   End With
   Exit Sub
DialogError:
   If Err.Number <> cdlCancel Then
      MsgBox "Error in Dialog's use: " & Err.Description, vbOKOnly + vbCritical, "Error!"
      Exit Sub
   End If
End Sub

Private Sub cmdSaveHuffman_Click()
On Error GoTo DialogError
   With cdgOpen
      .CancelError = True  ' Generate Error number cdlCancel if user click Cancel
      .InitDir = App.Path    ' Initial (i.e. default ) Folder
      .Filter = "Input File (*.OUT) | *.out"
      .FilterIndex = 1  ' Select ""Executables (*.exe) | *.exe" as default
      .DialogTitle = "Save information to a output file"
      .ShowSave
      Call SaveOut(cdgOpen)
   End With
   Exit Sub
DialogError:
   If Err.Number <> cdlCancel Then
      MsgBox "Error in Dialog's use: " & Err.Description, vbOKOnly + vbCritical, "Error!"
      Exit Sub
   End If
End Sub

Private Sub cmdSavep_Click()
On Error GoTo DialogError
   With cdgOpen
      .CancelError = True  ' Generate Error number cdlCancel if user click Cancel
      .InitDir = App.Path    ' Initial (i.e. default ) Folder
      .Filter = "Input File (*.INP) | *.inp"
      .FilterIndex = 1  ' Select ""Executables (*.exe) | *.exe" as default
      .DialogTitle = "Save information to a input file"
      .ShowSave
      Call SaveInp(cdgOpen)
   End With
   Exit Sub
DialogError:
   If Err.Number <> cdlCancel Then
      MsgBox "Error in Dialog's use: " & Err.Description, vbOKOnly + vbCritical, "Error!"
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 0 To 32
        txtp(i).Text = "0.0"
    Next i
End Sub

