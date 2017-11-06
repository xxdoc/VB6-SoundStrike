VERSION 5.00
Begin VB.Form frmAcerca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de..."
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcerca.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Frame frmProgramador 
      Height          =   2655
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.Label lblFunk 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Funk"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   17
         ToolTipText     =   "Por templar los ánimos"
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label lblRocaket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "r0cak3t"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1425
         TabIndex        =   16
         ToolTipText     =   "Por usar siempre la Glock"
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label lblFlus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Flus++"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1470
         TabIndex        =   15
         ToolTipText     =   "Por ser el mejor aZmin"
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblPeluche 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Peluche de Combate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   825
         TabIndex        =   14
         ToolTipText     =   "Por su sinceridad"
         Top             =   1560
         Width           =   1980
      End
      Begin VB.Label lblÑ 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[ñ]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1665
         TabIndex        =   13
         ToolTipText     =   "Por su pasión por el esfuerzo"
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label lblAgradecimientos 
         AutoSize        =   -1  'True
         Caption         =   "Agradecimientos a:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label lblProgramador 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[ DVD ]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1470
         MouseIcon       =   "frmAcerca.frx":1E72
         TabIndex        =   11
         ToolTipText     =   "Porque yo lo valgo :)"
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lblProgramacion 
         AutoSize        =   -1  'True
         Caption         =   "Programación:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame frmPrograma 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.Label lblCSS 
         AutoSize        =   -1  'True
         Caption         =   "Counter Strike Source"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   8
         Top             =   2130
         Width           =   2145
      End
      Begin VB.Label lblCS 
         AutoSize        =   -1  'True
         Caption         =   "Counter Strike v.1.6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   7
         Top             =   1830
         Width           =   1950
      End
      Begin VB.Label lblVB6SP6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Visual Basic 6.0 SP6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   6
         Top             =   1080
         Width           =   1950
      End
      Begin VB.Label lblCompatible 
         AutoSize        =   -1  'True
         Caption         =   "Compatible con:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label lblProgramado 
         AutoSize        =   -1  'True
         Caption         =   "Programado en:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "v.0.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Sound-Strike"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1710
      End
      Begin VB.Image imgCSS 
         Height          =   240
         Left            =   600
         Picture         =   "frmAcerca.frx":217C
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgCS 
         Height          =   240
         Left            =   600
         Picture         =   "frmAcerca.frx":786E
         Top             =   1800
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Versión " + CStr(App.Major) + "." + CStr(App.Minor)
End Sub
