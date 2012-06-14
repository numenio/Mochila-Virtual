VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmVerLibros 
   Caption         =   "Libros guardados"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   Icon            =   "frmVerLibros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmVerLibros.frx":08CA
   ScaleHeight     =   6360
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOrdenar 
      Caption         =   "Reordenar capítulos"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   1755
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btnComando 
      Caption         =   "Modificar"
      Height          =   375
      Index           =   2
      Left            =   5250
      TabIndex        =   1
      Top             =   5760
      Width           =   1755
   End
   Begin VB.CommandButton btnComando 
      Caption         =   "Eliminar"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   0
      Top             =   5760
      Width           =   1755
   End
   Begin ComctlLib.TreeView árbol 
      Height          =   5415
      Left            =   125
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9551
      _Version        =   327682
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imágenes"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList imágenes 
      Left            =   2400
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   28
      ImageHeight     =   29
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerLibros.frx":2922
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerLibros.frx":32F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerLibros.frx":504A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVerLibros.frx":702C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmVerLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub llenarLibrosMaterias()
'    Dim materia As String, libro As String, nombreCapítulo As String ', nombreArchivo As String
'
'    materia = Combo7.Text 'se llena la materia
'    libro = Combo2.Text
''    númeroCapítulo = Int(Combo4.Text)
'    nombreCapítulo = Trim(Text6.Text)
'
'    File1.Path = App.Path + "\trabajos\" + materia + "\libros\" + libro + "\" 'se ve si no hay ya otro capítulo con el mismo nombre
    btnComando(1).Enabled = False
    btnComando(2).Enabled = False
    btnOrdenar.Enabled = False
    
    árbol.Nodes.Clear
    árbol.Nodes.Add , , "root", Trim(nombreUsuario), 1
    
    Dim j As Integer, z As Integer, i As Integer, p As Integer, q As Integer
    z = 0
    q = 0
    
    Dim archivolibre As Byte, cadena As String, cadenaAux As String, capítulo As String, archivolibre2 As Byte
    archivolibre = FreeFile 'se abren las materias
    Open App.path + "\datos\materias.txt" For Input As archivolibre
    While Not EOF(archivolibre)
        Line Input #archivolibre, cadena
        árbol.Nodes.Add "root", tvwChild, "materia" & i, Trim(cadena), 2
        Dir1.path = App.path + "\trabajos\" + Trim(cadena) + "\libros\"
        For j = 0 To Dir1.ListCount - 1
            árbol.Nodes.Add "materia" & i, tvwChild, "libro" & z, Right(Dir1.List(j), Len(Dir1.List(j)) - InStrRev(Dir1.List(j), "\")), 4
'            File1.Path = Dir1.List(j)
'            For p = 0 To File1.ListCount - 1
'                If Right(File1.List(p), 4) = ".rtf" Then 'si es un archivo rtf, o sea que es un capítulo
'                    árbol.Nodes.Add "libro" & z, tvwChild, "capítulo" & q, Left(Right(File1.List(p), Len(File1.List(p)) - InStrRev(File1.List(p), "\")), Len(Right(File1.List(p), Len(File1.List(p)) - InStrRev(File1.List(p), "\"))) - 4), 3
'                    q = q + 1
'                End If
'            Next
            'se carga en la lista los capítulos en orden
            cadenaAux = Dir1.List(j) + "\ordenCapítulos"
            If existeCarpeta(cadenaAux) Then
                archivolibre2 = FreeFile
                Open cadenaAux For Input As #archivolibre2 'se abre el trabajo ya guardado
                Do While Not EOF(archivolibre2)
                    Input #archivolibre2, capítulo
                    If Trim(capítulo) <> "" Then árbol.Nodes.Add "libro" & z, tvwChild, "capítulo" & q, capítulo, 3
                    q = q + 1
                Loop
                Close #archivolibre2
            End If
            z = z + 1
        Next
        i = i + 1
    Wend
    Close #archivolibre
    
    árbol.Nodes(1).Expanded = True
End Sub

Private Sub árbol_Collapse(ByVal Node As ComctlLib.Node)
    'si al contraerse el árbol no se selecciona ni materia ni root, se permite eliminar y modificar
    If Left(árbol.SelectedItem.Key, 1) <> "m" And Left(árbol.SelectedItem.Key, 1) <> "r" Then
        btnComando(1).Enabled = True
        btnComando(2).Enabled = True
        btnOrdenar.Enabled = True
    Else
        btnComando(1).Enabled = False
        btnComando(2).Enabled = False
        btnOrdenar.Enabled = False
    End If
End Sub

Private Sub árbol_NodeClick(ByVal Node As ComctlLib.Node)
    If Left(árbol.Nodes.Item(Node.Index).Key, 5) = "libro" Or Left(árbol.Nodes.Item(Node.Index).Key, 8) = "capítulo" Then
        If Left(árbol.Nodes.Item(Node.Index).Key, 5) = "libro" Then
            btnComando(1).Caption = "Eliminar libro"
            btnComando(2).Caption = "Modificar libro"
        End If
        
        If Left(árbol.Nodes.Item(Node.Index).Key, 8) = "capítulo" Then
            btnComando(1).Caption = "Eliminar capítulo"
            btnComando(2).Caption = "Modificar capítulo"
        End If
        
        btnComando(1).Enabled = True
        btnComando(2).Enabled = True
        btnOrdenar.Enabled = True
    Else
        btnComando(1).Enabled = False
        btnComando(2).Enabled = False
        btnOrdenar.Enabled = False
    End If
End Sub

Private Sub btnComando_Click(Index As Integer)
    Dim cadena As String, swLibro As Boolean, materia As String, libro As String, nombre As String
    'se da por hecho que si se pudo hacer click está seleccionado un libro o un capítulo, pues si no los botones no están enabled
    If Index = 1 Then 'si es el botón eliminar
        If Left(árbol.SelectedItem.Key, 5) = "libro" Then
            cadena = "¿Realmente desea eliminar el libro " + Chr(34) + árbol.SelectedItem.Text + Chr(34) + "?"
            swLibro = True
        Else
            cadena = "¿Realmente desea eliminar el capítulo " + Chr(34) + árbol.SelectedItem.Text + Chr(34) + " del libro " + Chr(34) + árbol.SelectedItem.Parent.Text + Chr(34) + "?"
            swLibro = False
        End If
        frmMsgBox.cadenaAMostrar = cadena
        frmMsgBox.swSíNoóAceptar = True 'se elige que sea cuadro sí-no
        frmMsgBox.Show 1
        If frmMsgBox.swResultadoMostrado Then
            nombre = árbol.SelectedItem.Text
            If swLibro = True Then
                materia = árbol.SelectedItem.Parent.Text
                Call eliminarLibro(materia, nombre)
            Else
                materia = árbol.SelectedItem.Parent.Parent.Text
                libro = árbol.SelectedItem.Parent.Text
                nombre = nombre + ".rtf"
                Call eliminarCapítulo(materia, libro, nombre)
            End If
            Call llenarLibrosMaterias
        End If
    Else 'si es el botón modificar
        If Left(árbol.SelectedItem.Key, 5) = "libro" Then 'si es un libro
            materia = árbol.SelectedItem.Parent.Text
            libro = árbol.SelectedItem.Text
            Call modificarLibro(materia, libro)
        Else 'si es un capítulo
            materia = árbol.SelectedItem.Parent.Parent.Text
            libro = árbol.SelectedItem.Parent.Text
            nombre = árbol.SelectedItem.Text
            Call modificarCapítuloLibro(materia, libro, nombre)
        End If
    End If
End Sub

Sub modificarCapítuloLibro(materia As String, libro As String, nombre As String)
    frmAñadirCapítuloLibro.swEditar = True
    frmAñadirCapítuloLibro.swQuéMateria = materia
    frmAñadirCapítuloLibro.swQuéLibro = libro
    frmAñadirCapítuloLibro.swQuéCapítulo = nombre
    frmAñadirCapítuloLibro.Show 1
End Sub


Sub modificarLibro(materia As String, libro As String)
    frmAñadirLibro.swEditar = True
    frmAñadirLibro.swEditarLibroQuéMateria = materia
    frmAñadirLibro.swEditarQuéLibro = libro
    frmAñadirLibro.Show 1
End Sub

Sub eliminarLibro(materia As String, nombre As String)
    Dim i As Integer
    File1.path = App.path + "\trabajos\" + materia + "\libros\" + nombre
    For i = 0 To File1.ListCount - 1
        Kill File1.path + "\" + File1.List(i)
    Next
    RmDir (App.path + "\trabajos\" + materia + "\libros\" + nombre)
End Sub

Sub eliminarCapítulo(materia As String, libro As String, nombre As String)
    'arreglar que también borre el orden de los capítulos en el txt
    Kill (App.path + "\trabajos\" + materia + "\libros\" + libro + "\" + nombre)
    
    Dim archivolibre As Byte, cadena() As String, contador As Long, i As Long ', archivolibre2 As Byte
    archivolibre = FreeFile
    'archivolibre2 = FreeFile
    contador = 0
    Open App.path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Input As archivolibre
    
    Do While Not EOF(archivolibre)
        ReDim Preserve cadena(0 To contador)
        Line Input #archivolibre, cadena(contador)
        contador = contador + 1
    Loop
    
    Close #archivolibre
    archivolibre = FreeFile
    'Kill App.path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos"
    Open App.path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Output As archivolibre
    For i = 0 To UBound(cadena)
        If cadena(i) <> Left(nombre, Len(nombre) - 4) Then
            Print #archivolibre, cadena(i)
        End If
    Next
    Close #archivolibre
End Sub

Private Sub btnOrdenar_Click()
    Dim materia As String, libro As String
    'se llenan las materia y el libro
    If Left(árbol.SelectedItem.Key, 5) = "libro" Then
        materia = árbol.SelectedItem.Parent.Text
        libro = árbol.SelectedItem.Text
    Else 'si está seleccionado un capítulo
        materia = árbol.SelectedItem.Parent.Parent.Text
        libro = árbol.SelectedItem.Parent.Text
    End If
        
    'si no es el primer capítulo
    If existeCarpeta(App.path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos") Then
        frmOrdenarCapítulos.swMateria = materia
        frmOrdenarCapítulos.swLibro = libro
        frmOrdenarCapítulos.Show 1
    Else
        frmMsgBox.swMostrarCancelar = False
        frmMsgBox.cadenaAMostrar = "El libro elegido no tiene capítulos en él, así que no se pueden reordenar."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro acepar
        frmMsgBox.Show 1
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    shiftkey = Shift And 7
    If KeyCode = vbKeyF1 Then ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/ver libros.htm", "", 1
    
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Call llenarLibrosMaterias
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    btnComando(1).Enabled = False
    btnComando(2).Enabled = False
    btnOrdenar.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call contarFormularios(False)
End Sub

Public Sub actualizarÁrbol()
    Call llenarLibrosMaterias
End Sub
