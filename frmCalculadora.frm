VERSION 5.00
Begin VB.Form frmCalculadora 
   Caption         =   "Calculadora"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "."
      Height          =   495
      Index           =   10
      Left            =   2400
      TabIndex        =   19
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   2400
      TabIndex        =   18
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Width           =   4215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Borrar todo el cálculo"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Borrar todo el número"
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Borrar un número"
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "/"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "*"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "_"
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "="
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   2055
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Text1 = Text1 + Command1(Index).Caption
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode >= 48 And KeyCode <= 57) Then Number_Click (Chr(KeyCode))
    If (KeyCode >= 96 And KeyCode <= 105) Then
        aux = KeyCode - 48
        Text1 = Text1 + Chr(aux)
    End If
End Sub

