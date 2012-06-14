VERSION 5.00
Begin VB.Form frmOrdenarCapítulos 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5220
   Icon            =   "frmOrdenarCapítulos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmOrdenarCapítulos.frx":08CA
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
      Picture         =   "frmOrdenarCapítulos.frx":2922
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar capítulo"
      Top             =   2360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   1575
      Index           =   1
      Left            =   4770
      Picture         =   "frmOrdenarCapítulos.frx":297D
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
      Picture         =   "frmOrdenarCapítulos.frx":2A12
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
      Caption         =   "Suba o baje los capítulos del libro seleccionado para que se muestren en el orden que usted desee"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   330
      TabIndex        =   5
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Capítulos ya añadidos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   323
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmOrdenarCapítulos"
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
    Dim auxÍndice
    Dim auxCadena
    
    If List1.ListIndex = -1 Then '(list1.ListIndex) = "" Then
        frmMsgBox.cadenaAMostrar = "Primero elija un capítulo de la lista y luego pulse este botón"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If Index = 0 Then 'si se presiona el botón subir
        If List1.ListIndex <> 0 Then 'si la materia a subir no está ya arriba
            auxÍndice = List1.ListIndex - 1 'se cargan en los auxiliares los datos del elemento a bajar
            auxCadena = List1.List(List1.ListIndex - 1)
            List1.List(auxÍndice) = List1.List(List1.ListIndex) 'se hace el cambio para abajo
            List1.List(List1.ListIndex) = auxCadena
            List1.ListIndex = auxÍndice
        End If
    Else 'si se presiona el botón bajar
        If List1.ListIndex <> List1.ListCount - 1 Then 'se chequea que no esté en el fin la materia a bajar
            auxÍndice = List1.ListIndex + 1 'se cargan en los auxiliares los datos del elemento a bajar
            auxCadena = List1.List(List1.ListIndex + 1)
            
            List1.List(auxÍndice) = List1.List(List1.ListIndex) 'se hace el cambio para abajo
            List1.List(List1.ListIndex) = auxCadena
            List1.ListIndex = auxÍndice
        End If
    End If

End Sub

Private Sub Command6_Click()
    If List1.ListCount = 0 Then
        frmMsgBox.cadenaAMostrar = "No se puede eliminar ningún capítulo pues la lista está vacía."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If List1.ListIndex = -1 Then '(list1.ListIndex) = "" Then
        frmMsgBox.cadenaAMostrar = "Antes de apretar este botón tiene que seleccionar algún capítulo de la lista."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    Dim aux As Integer
    
    frmMsgBox.swMostrarCancelar = False
    frmMsgBox.cadenaAMostrar = "¿Realmente querés eliminar el capítulo?"
    frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
    frmMsgBox.Show 1
    If frmMsgBox.swResultadoMostrado = True Then
        Call eliminarCapítulo(swMateria, swLibro, List1.List(List1.ListIndex)) 'se elimina el archivo
        If List1.ListIndex < List1.ListCount - 1 Then 'se elimina el nombre de la lista de capítulos
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
    Dim i As Long, archivolibre As Byte, capítulo As String
    'si existe el archivo que ordena los capítulos, o sea que no es el primer capítulo a agregar
    
    Me.Caption = "Reordenar capítulos"
    Label1 = "Capítulos del libro " + Chr(34) + swLibro + Chr(34) + "."
    archivolibre = FreeFile
    'se carga en la lista los capítulos en orden
    Open App.Path + "\trabajos\" + swMateria + "\libros\" + swLibro + "\ordenCapítulos" For Input As #archivolibre 'se abre el trabajo ya guardado
    List1.Clear
    Do While Not EOF(archivolibre)
        Input #archivolibre, capítulo
        If Trim(capítulo) <> "" Then List1.AddItem capítulo
    Loop
    Close #archivolibre
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call guardarOrdenCapítulosLibro(List1, swMateria, swLibro)
    frmVerLibros.actualizarÁrbol
End Sub

Sub eliminarCapítulo(materia As String, libro As String, nombre As String)
    Kill (App.Path + "\trabajos\" + materia + "\libros\" + libro + "\" + nombre)
End Sub

