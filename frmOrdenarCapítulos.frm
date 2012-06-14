VERSION 5.00
Begin VB.Form frmOrdenarCap�tulos 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5220
   Icon            =   "frmOrdenarCap�tulos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmOrdenarCap�tulos.frx":08CA
   ScaleHeight     =   6750
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1223
      TabIndex        =   6
      Top             =   6120
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Height          =   1215
      Left            =   4770
      Picture         =   "frmOrdenarCap�tulos.frx":2922
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar cap�tulo"
      Top             =   2360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   1575
      Index           =   1
      Left            =   4770
      Picture         =   "frmOrdenarCap�tulos.frx":297D
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Bajar un lugar"
      Top             =   3780
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   1575
      Index           =   0
      Left            =   4770
      Picture         =   "frmOrdenarCap�tulos.frx":2A12
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Subir un lugar"
      Top             =   600
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   330
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Suba o baje los cap�tulos del libro seleccionado para que se muestren en el orden que usted desee"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   330
      TabIndex        =   5
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cap�tulos ya a�adidos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   323
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmOrdenarCap�tulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public swMateria As String
Public swLibro As String


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command4_Click(Index As Integer)
    Dim aux�ndice
    Dim auxCadena
    
    If List1.ListIndex = -1 Then '(list1.ListIndex) = "" Then
        frmMsgBox.cadenaAMostrar = "Primero elija un cap�tulo de la lista y luego pulse este bot�n"
        frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If Index = 0 Then 'si se presiona el bot�n subir
        If List1.ListIndex <> 0 Then 'si la materia a subir no est� ya arriba
            aux�ndice = List1.ListIndex - 1 'se cargan en los auxiliares los datos del elemento a bajar
            auxCadena = List1.List(List1.ListIndex - 1)
            List1.List(aux�ndice) = List1.List(List1.ListIndex) 'se hace el cambio para abajo
            List1.List(List1.ListIndex) = auxCadena
            List1.ListIndex = aux�ndice
        End If
    Else 'si se presiona el bot�n bajar
        If List1.ListIndex <> List1.ListCount - 1 Then 'se chequea que no est� en el fin la materia a bajar
            aux�ndice = List1.ListIndex + 1 'se cargan en los auxiliares los datos del elemento a bajar
            auxCadena = List1.List(List1.ListIndex + 1)
            
            List1.List(aux�ndice) = List1.List(List1.ListIndex) 'se hace el cambio para abajo
            List1.List(List1.ListIndex) = auxCadena
            List1.ListIndex = aux�ndice
        End If
    End If

End Sub

Private Sub Command6_Click()
    If List1.ListCount = 0 Then
        frmMsgBox.cadenaAMostrar = "No se puede eliminar ning�n cap�tulo pues la lista est� vac�a."
        frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If List1.ListIndex = -1 Then '(list1.ListIndex) = "" Then
        frmMsgBox.cadenaAMostrar = "Antes de apretar este bot�n tiene que seleccionar alg�n cap�tulo de la lista."
        frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    Dim aux As Integer
    
    frmMsgBox.swMostrarCancelar = False
    frmMsgBox.cadenaAMostrar = "�Realmente quer�s eliminar el cap�tulo?"
    frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
    frmMsgBox.Show 1
    If frmMsgBox.swResultadoMostrado = True Then
        Call eliminarCap�tulo(swMateria, swLibro, List1.List(List1.ListIndex)) 'se elimina el archivo
        If List1.ListIndex < List1.ListCount - 1 Then 'se elimina el nombre de la lista de cap�tulos
            aux = List1.ListIndex
        Else
            aux = List1.ListIndex - 1
        End If
        List1.RemoveItem List1.ListIndex
        List1.SetFocus
        If List1.ListCount <> 0 Then List1.ListIndex = aux
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long, archivolibre As Byte, cap�tulo As String
    'si existe el archivo que ordena los cap�tulos, o sea que no es el primer cap�tulo a agregar
    
    Me.Caption = "Reordenar cap�tulos"
    Label1 = "Cap�tulos del libro " + Chr(34) + swLibro + Chr(34) + "."
    archivolibre = FreeFile
    'se carga en la lista los cap�tulos en orden
    Open App.Path + "\trabajos\" + swMateria + "\libros\" + swLibro + "\ordenCap�tulos" For Input As #archivolibre 'se abre el trabajo ya guardado
    List1.Clear
    Do While Not EOF(archivolibre)
        Input #archivolibre, cap�tulo
        If Trim(cap�tulo) <> "" Then List1.AddItem cap�tulo
    Loop
    Close #archivolibre
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call guardarOrdenCap�tulosLibro(List1, swMateria, swLibro)
    frmVerLibros.actualizar�rbol
End Sub

Sub eliminarCap�tulo(materia As String, libro As String, nombre As String)
    Kill (App.Path + "\trabajos\" + materia + "\libros\" + libro + "\" + nombre)
End Sub

