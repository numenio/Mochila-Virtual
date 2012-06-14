VERSION 5.00
Begin VB.Form frmAñadirLibro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir un libro"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmAñadirLibro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAñadirLibro.frx":08CA
   ScaleHeight     =   2925
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   4800
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   270
      TabIndex        =   1
      Top             =   600
      Width           =   6375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Guardar el libro"
      Height          =   375
      Left            =   2550
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4920
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba aquí el nombre del libro a añadir:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   270
      TabIndex        =   5
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija la materia a que pertenece el libro:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   270
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "frmAñadirLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim swEmpezando As Boolean
Dim nombreOriginalMateria As String
Dim nombreOriginalLibro As String
Dim huboCambioEnlibro As Boolean
Public swEditar As Boolean
Public swEditarQuéLibro As String
Public swEditarLibroQuéMateria As String


'Private Sub Combo3_Change()
'    huboCambioEnlibro = True
'End Sub

Private Sub Command11_Click()
    If Trim(Text5 = "") Then
'        MsgBox "Antes de apretar este botón escriba el nombre del libro a añadir.", , "Cuidado"
        frmMsgBox.cadenaAMostrar = "Antes de apretar este botón escriba el nombre del libro a añadir."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Text5.SetFocus
        Exit Sub
    End If
    
    If Combo3.Text = "" Then 'si no se ha elegido una materia
'        MsgBox "No se ha elegido la materia a que pertenece el libro. Por favor, seleccione una y vuelva a presionar el botón " + Chr(34) + "Añadir" + Chr(34), , "Cuidado"
        frmMsgBox.cadenaAMostrar = "No se ha elegido la materia a que pertenece el libro. Por favor, seleccione una y vuelva a presionar el botón " + Chr(34) + "Añadir" + Chr(34)
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Combo3.SetFocus
        Exit Sub
    End If
        
    Dim materia As String, nombre As String, i As Integer
'    If swEditar = False Then 'si no se está editando un libro ya guardado, o sea que es uno nuevo
        materia = Combo3.Text
        nombre = Trim(Text5.Text)
'    Else 'si se edita uno ya guardado
'        materia = swEditarLibroQuéMateria
'        nombre = swEditarQuéLibro
'    End If
    
    If materia <> nombreOriginalMateria Or nombre <> nombreOriginalLibro Then
        Dir1.Path = App.Path + "\trabajos\" + materia + "\libros\"
        For i = 0 To Dir1.ListCount - 1
            If Dir1.List(i) = App.Path + "\trabajos\" + materia + "\libros\" + nombre Then
'                MsgBox "Ya hay un libro en la materia " + materia + " con el nombre " + Chr(34) + Trim(Text5) + Chr(34) + ". No se puede añadir.", , "Nombre repetido"
                frmMsgBox.cadenaAMostrar = "Ya hay un libro en la materia " + materia + " con el nombre " + Chr(34) + Trim(Text5) + Chr(34) + ". No se puede añadir."
                frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
                frmMsgBox.Show 1
                Text5.SelStart = 0
                Text5.SelLength = Len(Text5)
                Text5.SetFocus
                Exit Sub
            End If
        Next
        
        'se crea la ubicación para el libro
        MkDir (App.Path + "\trabajos\" + materia + "\libros\" + nombre)
        
        If swEditar = True Then
             'se copian todos los capítulos del libro a la nueva ubicación
             File1.Path = App.Path + "\trabajos\" + nombreOriginalMateria + "\libros\" + nombreOriginalLibro
             For i = 0 To File1.ListCount - 1
                 Call FileCopy(File1.Path + "\" + File1.List(i), App.Path + "\trabajos\" + materia + "\libros\" + nombre + "\" + File1.List(i))
             Next
            
            'se elimina el libro
            Call eliminarLibro(nombreOriginalMateria, nombreOriginalLibro)
    '             MsgBox "El libro se modificó exitosamente", , "Libro modificado"
            frmMsgBox.cadenaAMostrar = "El libro se modificó exitosamente" '"Hoy es " + cadenaFecha + ". Es la hora " + cadenaTiempo
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
            frmVerLibros.actualizarÁrbol
        Else
'            MsgBox "El libro se guardó exitosamente", , "Libro guardado"
            frmMsgBox.cadenaAMostrar = "El libro se guardó exitosamente"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
        End If
            
        
'        Text5 = ""
        
        Unload Me
    Else
        If swEditar = True Then
'            MsgBox "El nombre del libro y la materia están igual que antes, no se han hecho cambios", , "Información"
            frmMsgBox.cadenaAMostrar = "El nombre del libro y la materia están igual que antes, no se han hecho cambios"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call centrarFormulario(Me)
    'Call contarFormularios(True)
    Call llenarComboMaterias(Combo3)
    If swEditar = True Then
        For i = 0 To Combo3.ListCount - 1 'se pone el foco en la materia enviada
            If Combo3.List(i) = swEditarLibroQuéMateria Then
                Combo3.ListIndex = i
                Exit For
            End If
        Next
        
        Text5 = swEditarQuéLibro
        nombreOriginalMateria = swEditarLibroQuéMateria
        nombreOriginalLibro = swEditarQuéLibro
        Command11.Caption = "Guardar modificaciones"
'        huboCambioEnlibro = False
    End If
    swEmpezando = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
    
    If shiftkey = 0 And KeyCode = vbKeyF1 Then 'leer la ayuda
         ShellExecute 0, "open", "hh.exe", App.Path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/añadir libros.htm", "", 1
         Exit Sub
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
End Sub


Private Sub Form_Paint()
    If swEmpezando = True Then
        Text5.SetFocus
        swEmpezando = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call contarFormularios(False)
    swEditar = False
    swEditarQuéLibro = ""
    swEditarLibroQuéMateria = ""
End Sub

'Private Sub Text5_Change()
'    huboCambioEnlibro = True
'End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call controlarCaracteresEspeciales(KeyCode, Text5)
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim tecla As Integer
    If Text5 = "" Then
        tecla = 0
    Else
        tecla = Asc(Right(Text5, 1))
    End If
    If tecla > 0 Then Call controlarCaracteresEspeciales(tecla, Text5)
End Sub

Sub eliminarLibro(materia As String, nombre As String)
    Dim i As Integer
    File1.Path = App.Path + "\trabajos\" + materia + "\libros\" + nombre
    For i = 0 To File1.ListCount - 1
        Kill File1.Path + "\" + File1.List(i)
    Next
    RmDir (App.Path + "\trabajos\" + materia + "\libros\" + nombre)
End Sub

