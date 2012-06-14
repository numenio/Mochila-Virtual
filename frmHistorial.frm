VERSION 5.00
Object = "{9D4116BA-0A8D-4B9C-B752-263DC689B0EE}#1.0#0"; "TransparentButton.ocx"
Begin VB.Form frmHistorial 
   Caption         =   "Historial de Materias"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   Icon            =   "frmHistorial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmHistorial.frx":08CA
   ScaleHeight     =   5640
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin TransparentButton.ButtonTransparent ButtonTransparent1 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Al eliminar una materia del historial ya no se puede recuperar otra vez"
      Top             =   4920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      Caption         =   "Borrar definitivamente la materia"
      EstiloDelBoton  =   1
      Picture         =   "frmHistorial.frx":2922
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin TransparentButton.ButtonTransparent ButtonTransparent2 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Al eliminar una materia del historial ya no se puede recuperar otra vez"
      Top             =   4200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      Caption         =   "Recuperar la materia y añadirla a la mochila"
      EstiloDelBoton  =   1
      Picture         =   "frmHistorial.frx":31FC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      XPDefaultColors =   0   'False
      ForeColor       =   16777215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija qué materia del historial quiere recuperar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonTransparent1_Click() 'borrar una materia
    If List1.List(List1.ListIndex) = "No hay materia en el historial que no esté ya cargada" Then
'        MsgBox "No se puede eliminar ninguna materia pues la lista está vacía.", , "Lista vacía"
        frmMsgBox.cadenaAMostrar = "No se puede eliminar ninguna materia pues la lista está vacía."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If List1.List(List1.ListIndex) = "" Then 'se controla que no se
'        MsgBox "Antes de apretar este botón tiene que seleccionar alguna materia de la lista.", , "No hay materia seleccionada"
        frmMsgBox.cadenaAMostrar = "Antes de apretar este botón tiene que seleccionar alguna materia de la lista."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    Dim aux As Integer 'se toma un auxiliar para que el foco del cuadro queden en el elemento inferior
    If List1.ListIndex < List1.ListCount - 1 Then
        aux = List1.ListIndex
    Else
        aux = List1.ListIndex - 1
    End If
    
    borrarCarpeta (App.path + "\trabajos\" + List1.List(List1.ListIndex)) 'se borran los directorios de la materia
    
    List1.RemoveItem List1.ListIndex 'se quita la materia de la lista
    
    Dim i As Byte
    Open App.path + "\datos\historialMaterias.txt" For Output As #1 'se abre el historial ya guardado
    If List1.ListCount > 0 Then
        For i = 0 To List1.ListCount - 1 'se le escriben las materias tal como quedaron en la lista
            Print #1, List1.List(i)
        Next
    End If
    Close #1
    
    List1.SetFocus
    If List1.ListCount <> 0 Then List1.ListIndex = aux
    swHuboCambioEnMaterias = True 'para guardar los cambios
End Sub

Private Sub ButtonTransparent2_Click() 'recuperar una materia
    If List1.List(List1.ListIndex) = "No hay materia en el historial que no esté ya cargada" Then
        'MsgBox "No se puede recuperar una materia pues la lista está vacía.", , "Lista vacía"
        frmMsgBox.cadenaAMostrar = "No se puede recuperar una materia pues la lista está vacía."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    If List1.List(List1.ListIndex) = "" Then
        'MsgBox "Antes de apretar este botón seleccione una materia del historial.", , "¡Cuidado!"
        frmMsgBox.cadenaAMostrar = "Antes de apretar este botón seleccione una materia del historial."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 0 To frmControl.List2.ListCount - 1
        If List1.List(List1.ListIndex) = frmControl.List2.List(i) Then
           ' MsgBox "Ya hay una materia en la Mochila con el mismo nombre de la que está intentando recuperar. No pueden haber dos materias con el mismo nombre", , "Imposible recuperar"
            frmMsgBox.cadenaAMostrar = "Ya hay una materia en la Mochila con el mismo nombre de la que está intentando recuperar. No pueden haber dos materias con el mismo nombre"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
            Exit Sub
        End If
    Next
    
    If frmControl.List2.ListCount = 28 Then
        'MsgBox "Ya se han agregado las 28 materias que acepta este programa. Para añadir una materia distinta, por favor borre una ya guardada", , "Información"
        frmMsgBox.cadenaAMostrar = "Ya se han agregado las 28 materias que acepta este programa. Para añadir una materia distinta, por favor borre una ya guardada"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Exit Sub
    End If
    
'    Dim i As Integer
'    For i = 0 To frmControl.List2.ListCount - 1 'se controla que no haya una materia ya añadida con el mismo nombre
'        If Trim(Text2) = List2.List(i) Then
'            MsgBox "Ya hay una materia con el nombre " + Chr(34) + Trim(Text2) + Chr(34) + ". No se puede añadir.", , "Nombre repetido"
'            Text2.SelStart = 0
'            Text2.SelLength = Len(Text2)
'            Text2.SetFocus
'            Exit Sub
'        End If
'    Next
    
    frmControl.List2.AddItem Trim(List1.List(List1.ListIndex))
'    frmControl.Combo12.AddItem Trim(List1.List(List1.ListIndex))
    
    Dim aux As Integer
    If List1.ListIndex < List1.ListCount - 1 Then
        aux = List1.ListIndex
    Else
        aux = List1.ListIndex - 1
    End If
    
    List1.RemoveItem List1.ListIndex
    
    List1.SetFocus
    If List1.ListCount <> 0 Then List1.ListIndex = aux
    
    swHuboCambioEnMaterias = True 'para guardar los cambios
End Sub

Private Sub Form_DblClick()
    ButtonTransparent2_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    
    If KeyCode = vbKeyF1 Then ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/config historial.htm", "", 1
End Sub

Private Sub Form_Load()
    Dim archivolibre As Byte, i As Integer, materiaYaAgregada As Boolean, cadena As String
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    archivolibre = FreeFile 'se abren las materias
    Open App.path + "\datos\historialMaterias.txt" For Input As archivolibre
    While Not EOF(archivolibre)
        Line Input #archivolibre, cadena
        For i = 0 To frmControl.List2.ListCount - 1
            If frmControl.List2.List(i) = cadena Then
                materiaYaAgregada = True
                Exit For
            End If
        Next
        
        If materiaYaAgregada = False Then List1.AddItem Trim(cadena) 'se añaden las materias al listado y al combo
        
        materiaYaAgregada = False
    Wend
    Close #archivolibre
    If List1.ListCount = 0 Then List1.AddItem "No hay materia en el historial que no esté ya cargada"
End Sub


Sub borrarCarpeta(ruta As String)
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.deletefolder ruta, True 'se fuerza a borrar aunque sea de sólo lectura
    Set fs = Nothing 'se libera la memoria del objeto fs
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call guardarMaterias(frmControl.List2)
    swHuboCambioEnMaterias = False
    'Call contarFormularios(False)
End Sub
