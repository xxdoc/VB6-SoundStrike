VERSION 5.00
Begin VB.Form frmGuia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guía del Usuario"
   ClientHeight    =   6780
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
   Icon            =   "frmManual.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtGuia 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmManual.frx":1E72
      Top             =   120
      Width           =   7335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   6300
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Unload Me
End Sub
