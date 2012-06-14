VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAñadirCapítuloLibro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir capítulo al libro"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   Icon            =   "frmAñadirCapítuloLibro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAñadirCapítuloLibro.frx":08CA
   ScaleHeight     =   6675
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Puede añadir un libro aquí"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   960
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog diálogo 
      Left            =   1800
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   6030
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3270
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Width           =   3345
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   270
      TabIndex        =   2
      Top             =   1815
      Width           =   6375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Buscar archivo"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   2325
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Añadir capítulo"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   6000
      Width           =   2055
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   495
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox rtfLibros 
      Height          =   3015
      Left            =   240
      TabIndex        =   4
      Top             =   2805
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmAñadirCapítuloLibro.frx":29F9
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija el libro al que le quiere añadir capítulos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3270
      TabIndex        =   9
      Top             =   255
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba aquí el título o tema del capítulo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   270
      TabIndex        =   8
      Top             =   1575
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba abajo el capítulo a agregar, o haga clic aquí ------------------------->"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2445
      Width           =   5055
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija la materia a que pertenece el libro:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   270
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmAñadirCapítuloLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public swEditar As Boolean
Public swQuéMateria As String
Public swQuéLibro As String
Public swQuéCapítulo As String
Dim nombreOriginalLibro As String
Dim nombreOriginalMateria As String
Dim nombreOriginalCapítulo As String
Dim huboCambio As Boolean

Private Sub Command1_Click()
    frmAñadirLibro.Show 1
End Sub

Private Sub Command10_Click()
    Dim i As Integer
    If Combo2.Text = "" Or Combo2.Text = "No hay libros guardados para esta materia" Then
        frmMsgBox.cadenaAMostrar = "No se ha elegido el libro al cual añadir un capítulo. Por favor, elija uno"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Combo2.SetFocus
        Exit Sub
    End If
    
    If Trim(Text6 = "") Then
        frmMsgBox.cadenaAMostrar = "Antes de apretar este botón se debe escribir un nombre al capítulo, por favor escriba uno"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        Text6.SetFocus
        Exit Sub
    End If
    
    If Trim(rtfLibros.Text = "") Then
        frmMsgBox.cadenaAMostrar = "El capítulo que está intentando agregar está en blanco. Por favor, escriba en el cuadro y vuelva a apretar el botón " + Chr(34) + "Añadir" + Chr(34) + "."
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
        rtfLibros.SetFocus
        Exit Sub
    End If
    
    Dim materia As String, libro As String, nombreCapítulo As String ', nombreArchivo As String
    
    materia = Combo7.Text 'se llena la materia
    libro = Combo2.Text
    nombreCapítulo = Trim(Text6.Text)
    
    If materia <> nombreOriginalMateria Or libro <> nombreOriginalLibro Or nombreCapítulo <> nombreOriginalCapítulo Then
        File1.Refresh
        File1.Path = App.Path + "\trabajos\" + materia + "\libros\" + libro + "\" 'se ve si no hay ya otro capítulo con el mismo nombre
        
        
        For i = 0 To File1.ListCount - 1 'se ve si ya hay un libro guardado con el mismo nombre
            If File1.List(i) = nombreCapítulo + ".rtf" Then
                frmMsgBox.cadenaAMostrar = "Ya hay un capítulo con el mismo nombre, por favor escriba otro"
                frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
                frmMsgBox.Show 1
                Exit Sub
            End If
        Next
        
        'se guarda el capítulo, ya sea original o la modificación, y el orden del mismo
        rtfLibros.SaveFile App.Path + "\trabajos\" + materia + "\libros\" + libro + "\" + nombreCapítulo + ".rtf"
'        Call chequearEspacioEnDisco(Left(App.Path, 2))
        Call guardarOrdenCap(materia, nombreOriginalMateria, libro, nombreOriginalLibro, nombreCapítulo, nombreOriginalCapítulo)
        
        If swEditar = True Then 'si se esta editando
            'se elimina el capítulo original
            Call eliminarCapítuloLibro(nombreOriginalMateria, nombreOriginalLibro, nombreOriginalCapítulo)
            frmMsgBox.cadenaAMostrar = "El capítulo se modificó exitosamente"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
            frmVerLibros.actualizarÁrbol
        Else
            frmMsgBox.cadenaAMostrar = "El capítulo se guardó exitosamente"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
        End If
        Unload Me
    Else
        If swEditar = True Then
            frmMsgBox.cadenaAMostrar = "El nombre del capítulo, del libro y la materia están igual que antes, no se han hecho cambios"
            frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
            frmMsgBox.Show 1
        End If
    End If
End Sub

Private Sub Combo7_Click() 'cuando se elije una materia para ver sus libros
    Dim materia As String, i As Integer
    materia = Combo7.Text
    If materia <> "" Then
        Dir1.Refresh
        Dir1.Path = App.Path + "\trabajos\" + materia + "\libros\"
        Combo2.Clear
        For i = 0 To Dir1.ListCount - 1
            Combo2.AddItem Right(Dir1.List(i), Len(Dir1.List(i)) - InStrRev(Dir1.List(i), "\")) 'se añaden los libros de la materia elegida para añadir capítulos
        Next
        
        If Combo2.ListCount = 0 Then Combo2.AddItem "No hay libros guardados para esta materia"
    End If
End Sub


Private Sub Command9_Click()
    ' Establecer CancelError a True
    diálogo.CancelError = True
    On Error GoTo ErrHandler
    diálogo.Filter = "Archivos de Texto (*.txt; *.rtf)|*.txt;*.rtf|"
    diálogo.ShowOpen
    rtfLibros.LoadFile diálogo.FileName
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim shiftkey As Integer
    
    shiftkey = Shift And 7
    
    If shiftkey = 0 And KeyCode = vbKeyF1 Then 'leer la ayuda
         ShellExecute 0, "open", "hh.exe", App.Path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/añadir capítulos.htm", "", 1
         Exit Sub
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'Call contarFormularios(True)
    Call llenarComboMaterias(Combo7)
    'swPrimerCapítulo = extenderForm(swQuéMateria, swQuéLibro)
    Call centrarFormulario(Me)
    If swEditar = True Then
        For i = 0 To Combo7.ListCount - 1 'se pone el foco en la materia enviada
            If Combo7.List(i) = swQuéMateria Then
                Combo7.ListIndex = i
                Exit For
            End If
        Next
        
        Me.Caption = "Modificar capítulo guardado"
        
        Text6 = swQuéCapítulo
        rtfLibros.LoadFile App.Path + "\trabajos\" + swQuéMateria + "\libros\" + swQuéLibro + "\" + swQuéCapítulo + ".rtf"
        nombreOriginalMateria = swQuéMateria
        nombreOriginalLibro = swQuéLibro
        nombreOriginalCapítulo = swQuéCapítulo
        Command10.Caption = "Guardar modificaciones"
    Else
        nombreOriginalMateria = ""
        nombreOriginalLibro = ""
        nombreOriginalCapítulo = ""
    End If
End Sub

Private Sub Form_Paint()
    Dim i As Integer
    Combo7_Click 'se actualiza el combo2 por si se añadió un libro nuevo
    If swEditar = True Then
        For i = 0 To Combo2.ListCount - 1 'se pone el foco en el libro enviado
            If Combo2.List(i) = swQuéLibro Then
                Combo2.ListIndex = i
                Exit For
            End If
        Next
    End If
    'Combo7.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    swEditar = False
    swQuéMateria = ""
    swQuéCapítulo = ""
    swQuéLibro = ""
    Me.Refresh
    'Call contarFormularios(False)
End Sub

Private Sub Text6_Change()
    Dim tecla As Integer
    If Text6 = "" Then
        tecla = 0
    Else
        tecla = Asc(Right(Text6, 1))
    End If
    If tecla > 0 Then Call controlarCaracteresEspeciales(tecla, Text6)
End Sub

Sub eliminarCapítuloLibro(materia As String, libro As String, capítulo As String)
    Kill App.Path + "\trabajos\" + materia + "\libros\" + libro + "\" + capítulo + ".rtf"
End Sub

Function guardarOrdenCap(materia As String, nombreMateriaOriginal As String, libro As String, nombreLibroOriginal As String, nombreCap As String, nombreCapOriginal As String) As Boolean
    Dim matrizCapítulos() As String, matrizAux() As String, contador As Integer, swProblema As Boolean
    Dim i As Integer, aux As String, archivolibre As Byte
    On Error GoTo manejoError
    archivolibre = FreeFile
    If nombreOriginalMateria <> "" Then 'si no se está añadiendo un capítulo nuevo
        
        If materia <> nombreOriginalMateria Then 'si cambió de un libro de una materia a otro de otra materia
            Open App.Path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Append As #archivolibre
            Print #archivolibre, nombreCap
            Close #archivolibre
            'se abre la lista original
            Open App.Path + "\trabajos\" + nombreOriginalMateria + "\libros\" + nombreLibroOriginal + "\ordenCapítulos" For Input As #archivolibre
            If nombreCapOriginal <> nombreCap Then nombreCap = nombreCapOriginal 'si se cambió el nombre, que se borre el original
            contador = 0
            While Not EOF(archivolibre) 'llenamos una matriz con los capítulos de la lista original
                Line Input #archivolibre, aux
                If aux <> nombreCap Then 'se copian todos los cap menos el que se pasó a otra lista
                    ReDim Preserve matrizAux(0 To contador)
                    matrizAux(contador) = aux
                    contador = contador + 1
                End If
            Wend
        End If
        
        If libro <> nombreLibroOriginal And materia = nombreOriginalMateria Then 'si cambió de libro dentro de la misma materia
            Open App.Path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Append As #archivolibre
            Print #archivolibre, nombreCap
            Close #archivolibre
            'se modifica la lista original
            Open App.Path + "\trabajos\" + materia + "\libros\" + nombreLibroOriginal + "\ordenCapítulos" For Input As #archivolibre
            contador = 0
            While Not EOF(archivolibre) 'llenamos una matriz con los capítulos de la lista original
                Line Input #archivolibre, aux
                If aux <> nombreCap Then 'se copian todos los cap menos el que se pasó a otra lista
                    ReDim Preserve matrizAux(0 To contador)
                    matrizAux(contador) = aux
                    contador = contador + 1
                End If
            Wend
        End If

        If nombreCap <> nombreCapOriginal And materia = nombreOriginalMateria And libro = nombreLibroOriginal Then 'si cambió el nombre del capítulo
            'sólo se modifica la lista original
            Open App.Path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Input As #archivolibre
            contador = 0
            While Not EOF(archivolibre) 'llenamos una matriz con los capítulos de la lista original
                Line Input #archivolibre, aux
                ReDim Preserve matrizAux(0 To contador)
                If aux <> nombreCapOriginal Then 'se copian todos los cap menos el que se pasó a otra lista
                    matrizAux(contador) = aux
                Else
                    matrizAux(contador) = nombreCap
                End If
                contador = contador + 1
            Wend
        End If
        
        
        Close #archivolibre
        'se guarda la lista original sin el capítulo eliminado
        If contador > 0 Then 'si en la lista original hay algún capítulo guardado
            If Not guardarOrdenCapítulosDesdeMatriz(nombreOriginalMateria, nombreLibroOriginal, matrizAux) Then swProblema = True
        Else
            Open App.Path + "\trabajos\" + nombreOriginalMateria + "\libros\" + nombreLibroOriginal + "\ordenCapítulos" For Output As #archivolibre
            Close #archivolibre
        End If
    
    Else 'si es un capítulo nuevo
        Open App.Path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Append As #archivolibre
        Print #archivolibre, nombreCap
        Close #archivolibre
    End If
    
    If swProblema = True Then 'se evalúa si hubo problemas en la función
        guardarOrdenCap = False
    Else
        guardarOrdenCap = True
    End If
    'Close #1
    'guardarOrdenCap = True
    Exit Function
manejoError:
    Dim archivolibre1 As Byte
    archivolibre1 = FreeFile
    Open App.Path + "\trabajos\" + materia + "\libros\" + libro + "\ordenCapítulos" For Output As #archivolibre1
    Close #archivolibre1
    Resume
End Function
