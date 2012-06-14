VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLectorTarea 
   Caption         =   "Tarea del x/x/x"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   4830
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmLectorTarea.frx":0000
   End
End
Attribute VB_Name = "frmLectorTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'en lugar de pasar a este lector, el frmMesTareasX tiene que pasar al cuaderno

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftKey As Integer
    
    If KeyCode = vbKeyEscape Then
        frmMesTareasX.Show
        Unload Me
    End If
    
    ShiftKey = Shift And 7
    If ShiftKey = vbAltMask And KeyCode = vbKeyF4 Then End 'si presiona alt + f4 se termina el programa
End Sub

