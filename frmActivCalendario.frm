VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmActivDefVisual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actividades guardadas"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "frmActivCalendario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmActivCalendario.frx":08CA
   ScaleHeight     =   6945
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mostrar s�lo actividades de este a�o"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   7215
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar actividad"
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   5
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar actividad"
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   4
      Top             =   6360
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Top             =   1230
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.TreeView �rbolActividades 
      Height          =   3975
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      _Version        =   327682
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija un mes:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija una materia para ver sus actividades:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmActivDefVisual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim trabajos() As String

Private Sub �rbolActividades_Click()
    'si se ha seleccionado una actividad (se quita No porque es que no hay actividad)
    If Left(�rbolActividades.SelectedItem.Key, 1) = "a" And Left(�rbolActividades.SelectedItem.Text, 2) <> "No" Then
        Command2(0).Enabled = True
        Command2(1).Enabled = True
    Else 'si se seleccion� cualquier otra cosa
        Command2(0).Enabled = False
        Command2(1).Enabled = False
    End If
End Sub

'Private Sub �rbolActividades_Collapse(ByVal Node As ComctlLib.Node)
'    'si se ha seleccionado una actividad
'    If Left(�rbolActividades.SelectedItem.Key, 1) = "a" Then
'        Command2(0).Enabled = True
'        Command2(1).Enabled = True
'    Else 'si se seleccion� cualquier otra cosa
'        Command2(0).Enabled = False
'        Command2(1).Enabled = False
'    End If
'End Sub
'
'Private Sub �rbolActividades_Expand(ByVal Node As ComctlLib.Node)
'    'si se ha seleccionado una actividad
'    If Left(�rbolActividades.SelectedItem.Key, 1) = "a" Then
'        Command2(0).Enabled = True
'        Command2(1).Enabled = True
'    Else 'si se seleccion� cualquier otra cosa
'        Command2(0).Enabled = False
'        Command2(1).Enabled = False
'    End If
'End Sub

Private Sub �rbolActividades_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        'si se ha seleccionado una actividad
        If Left(�rbolActividades.SelectedItem.Key, 1) = "a" Then
            Command2(0).Enabled = True
            Command2(1).Enabled = True
        Else 'si se seleccion� cualquier otra cosa
            Command2(0).Enabled = False
            Command2(1).Enabled = False
        End If
    End If
End Sub

Private Sub Check6_Click()
    Dim mes As Byte
    mes = decodificarMes(Combo6.Text)
    Call llenar�rbolActividades(Combo5.Text, Trim(Str(mes)), Check6.Value)
End Sub

Private Sub Combo5_Click()
    Dim mes As Byte
    mes = decodificarMes(Combo6.Text)
    Call llenar�rbolActividades(Combo5.Text, Trim(Str(mes)), Check6.Value)
End Sub

Private Sub Combo6_Click()
    Dim mes As Byte
    mes = decodificarMes(Combo6.Text)
    Call llenar�rbolActividades(Combo5.Text, Trim(Str(mes)), Check6.Value)
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim mes As Date, a�o As String, d�a As String
    
'    If Left(�rbolActividades.SelectedItem.Text, 32) <> "No hay actividades guardadas de " Then
        If Left(�rbolActividades.SelectedItem.Key, 1) = "a" And Left(�rbolActividades.SelectedItem.Text, 2) <> "No" Then
            If Index = 1 Then 'si es el bot�n modificar
                If Check6.Value = 1 Then
                    a�o = Trim(Str(2008))
                Else
                    a�o = Right(�rbolActividades.SelectedItem.Parent.Text, 4)
                End If
                
                Dim i As Integer, swEmpez�N�mero As Boolean
                For i = 1 To Len(�rbolActividades.SelectedItem.Text)
                    If IsNumeric(Mid(�rbolActividades.SelectedItem.Text, i, 1)) Then
                        d�a = d�a + Mid(�rbolActividades.SelectedItem.Text, i, 1)
                        swEmpez�N�mero = True
                    Else
                        If swEmpez�N�mero = True Then Exit For
                    End If
                Next
                
                mes = Format(d�a + "/" + Trim(Str(decodificarMes(Combo6.Text))) + "/" + a�o)
                frmA�adirActividad.d�aParaCargarActividades = mes
                frmA�adirActividad.materia = Combo5.Text
                frmA�adirActividad.swCargarFecha = True
                frmA�adirActividad.swEditarActividades = True
                frmA�adirActividad.Show 1 'a�adir una actividad
            Else 'si es el bot�n eliminar
                frmMsgBox.cadenaAMostrar = "�Realmente desea eliminar la actividad del d�a " + �rbolActividades.SelectedItem.Text + "?"
                frmMsgBox.swS�No�Aceptar = True 'se elige que sea cuadro s�-no
                frmMsgBox.Show 1
                If frmMsgBox.swResultadoMostrado Then Call eliminarActividad(Combo5.Text, Trim(Str(decodificarMes(Combo6.Text))), trabajos(Int(Right(�rbolActividades.SelectedItem.Key, 1))))
                Call llenar�rbolActividades(Combo5.Text, Trim(Str(decodificarMes(Combo6.Text))), Check6.Value)
            End If
        End If
'    Else
'        Command2(0).Enabled = False
'        Command2(1).Enabled = False
'    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And 7 = vbAltMask And KeyCode = 18 Then 'se neutraliza el men� de ventana si se aprieta alt
        Shift = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyF1 Then ShellExecute 0, "open", "hh.exe", App.path + "\Ayuda\Ayuda_Mochila_Virtual_1.0.chm::/ver actividades con jaws.htm", "", 1
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim mes As Byte
    'Call contarFormularios(True)
    Call centrarFormulario(Me)
    Check6.Value = 1
    Combo6.AddItem "Enero"
    Combo6.AddItem "Febrero"
    Combo6.AddItem "Marzo"
    Combo6.AddItem "Abril"
    Combo6.AddItem "Mayo"
    Combo6.AddItem "Junio"
    Combo6.AddItem "Julio"
    Combo6.AddItem "Agosto"
    Combo6.AddItem "Setiembre"
    Combo6.AddItem "Octubre"
    Combo6.AddItem "Noviembre"
    Combo6.AddItem "Diciembre"
    
    mes = Mid(Format(Date, "dd/mm/yyyy"), 4, 2)
    
    Combo6.ListIndex = mes - 1
    
    Call llenarComboMaterias(Combo5)
    If Combo5.ListCount <> 0 Then Combo5.ListIndex = 0
End Sub

Sub llenar�rbolActividades(materia As String, mes As String, mostrarS�loA�oActual As Boolean)
    Dim j As Integer, z As Integer, archivolibre As Byte, cadena As String, cadenaAux As String, miRegistro As DatosActividad
    Dim actividades() As String, contador As Integer, a�o1 As String, a�o2 As String, i As Integer, contadorA�os As Integer
    Dim a�osA�adidos() As String, swA�adirNodo As Boolean, p As Integer, t As Integer
    File1.Refresh
    
    If materia <> "" And mes <> 0 Then
        �rbolActividades.Nodes.Clear
        �rbolActividades.Nodes.Add , , "root", Trim(materia) ', 3, 4
        z = 0
        File1.path = App.path + "\trabajos\" + materia + "\actividades\" + mes + "\"
        If mostrarS�loA�oActual = True Then
            For t = 1 To 31 'se eval�a por d�a
                For j = 0 To File1.ListCount - 1
                    If Left(Right(File1.List(j), 8), 4) = Trim(Str(Year(Date))) And Mid(File1.List(j), cantPrefijo + 1, 2) = t Then
                        cadena = File1.List(j)
                        cadena = Left(Right(cadena, Len(cadena) - cantPrefijo), Len(Right(cadena, Len(cadena) - cantPrefijo)) - 4)
                        cadenaAux = Left(cadena, 3) + mes + "-" + Right(cadena, 4)
                        cadenaAux = Format(Left(cadenaAux, 10))
                        cadenaAux = Format(cadenaAux, "Long Date")
                        
    '                    Open App.Path + "\datos\" + mes + "\datosActividades.gui" For Random As #1 Len = Len(miRegistro)
    '                    Do While Not EOF(1)   ' Repite hasta el final del archivo.
    '                       Get #1, , miRegistro   ' Lee el registro siguiente.
    '                       If Trim(miRegistro.DirArchivo) = File1.Path + "\" + File1.List(j) Then Exit Do
    '                    Loop
    '                    Close #1   ' Cierra el archivo.
                        
                        Open App.path + "\trabajos\" + materia + "\actividades\" + mes + "\datosActividades\" + Left(File1.List(j), Len(File1.List(j)) - 4) + ".gui" For Random As #2 Len = Len(miRegistro)
                        Get #2, 1, miRegistro   ' Lee el regitro
                        Close #2   ' Cierra el archivo.
                        
                        
                        If Asc(Left(miRegistro.tema, 1)) Then
                            cadenaAux = cadenaAux + ". Tema: " + Trim(miRegistro.tema) + "."
                        Else
                            cadenaAux = cadenaAux + ". Sin tema."
                        End If
                    
                        �rbolActividades.Nodes.Add "root", tvwChild, "actividad" & z, cadenaAux ', 7
                        ReDim Preserve trabajos(0 To z)
                        trabajos(z) = File1.List(j)
                        z = z + 1
                    End If
                Next
            Next
            If �rbolActividades.Nodes.Count = 1 Then �rbolActividades.Nodes.Add "root", tvwChild, "actividad" & z, "No hay actividades guardadas de " + materia + " para el mes de " + Combo6.List(Combo6.ListIndex)
        Else 'si se muestran las actividades de todos los a�os
            contador = 0
            For j = 0 To File1.ListCount - 1
                cadena = File1.List(j)
                cadena = Left(Right(cadena, Len(cadena) - cantPrefijo), Len(Right(cadena, Len(cadena) - cantPrefijo)) - 4)
                cadenaAux = Left(cadena, 3) + mes + "-" + Right(cadena, 4)
                cadenaAux = Format(Left(cadenaAux, 10))
                cadenaAux = Format(cadenaAux, "Long Date")
                
'                Open App.Path + "\datos\" + mes + "\datosActividades.gui" For Random As #1 Len = Len(miRegistro)
'                Do While Not EOF(1)   ' Repite hasta el final del archivo.
'                   Get #1, , miRegistro   ' Lee el registro siguiente.
'                   If Trim(miRegistro.DirArchivo) = File1.Path + "\" + File1.List(j) Then Exit Do
'                Loop
'                Close #1   ' Cierra el archivo.
                
                Open App.path + "\trabajos\" + materia + "\actividades\" + mes + "\datosActividades\" + Left(File1.List(j), Len(File1.List(j)) - 4) + ".gui" For Random As #2 Len = Len(miRegistro)
                Get #2, 1, miRegistro   ' Lee el regitro
                Close #2   ' Cierra el archivo.
                
                If Asc(Left(miRegistro.tema, 1)) Then
                    cadenaAux = cadenaAux + ". Tema: " + Trim(miRegistro.tema) + "."
                Else
                    cadenaAux = cadenaAux + ". Sin tema."
                End If
                
                ReDim Preserve actividades(0 To j)
                actividades(j) = Right(cadena, 4) + "-" + cadenaAux
                contador = 1
            Next
            
            If contador <> 0 Then 'si hay alguna actividad guardada en el mes seleccionado
                a�o1 = Left(actividades(0), 4)
                contador = 0
                For j = 0 To UBound(actividades)
                    swA�adirNodo = True
                    a�o2 = Left(actividades(j), 4)
                    If j <> 0 Then
                        If a�o1 <> a�o2 Then
                            For p = 0 To UBound(a�osA�adidos) 'se controla que el a�o no est� ya a�adido
                                If a�o2 = a�osA�adidos(p) Then swA�adirNodo = False
                            Next
                            
                            If swA�adirNodo = True Then
                                �rbolActividades.Nodes.Add "root", tvwChild, "ea�o" & a�o2, "A�o " + Left(actividades(j), 4) ', 7
                                ReDim Preserve a�osA�adidos(0 To contadorA�os)
                                a�osA�adidos(contadorA�os) = Left(actividades(j), 4)
                                contadorA�os = contadorA�os + 1
                                For z = 0 To UBound(actividades)
                                    If Left(actividades(z), 4) = a�o2 Then
                                        �rbolActividades.Nodes.Add "ea�o" + a�o2, tvwChild, "actividad" & i, Right(actividades(z), Len(actividades(z)) - InStr(1, actividades(z), "-")) ', 7
                                        ReDim Preserve trabajos(0 To i)
                                        trabajos(i) = File1.List(z)
                                        i = i + 1
                                    End If
                                Next
                                a�o1 = a�o2
                            End If
                        End If
                    Else 'si es el primer archivo
                        �rbolActividades.Nodes.Add "root", tvwChild, "ea�o" & a�o1, "A�o " + Left(actividades(j), 4) ', 7
                        ReDim Preserve a�osA�adidos(0 To contadorA�os)
                        a�osA�adidos(contadorA�os) = Left(actividades(j), 4)
                        contadorA�os = contadorA�os + 1
                        For z = 0 To UBound(actividades)
                            If Left(actividades(z), 4) = a�o1 Then
                                �rbolActividades.Nodes.Add "ea�o" + a�o1, tvwChild, "actividad" & i, Right(actividades(z), Len(actividades(z)) - InStr(1, actividades(z), "-")) ', 7
                                ReDim Preserve trabajos(0 To i)
                                trabajos(i) = File1.List(z)
                                i = i + 1
                            End If
                        Next
                        a�o1 = a�o2
                    End If
                Next
            Else
                �rbolActividades.Nodes.Add "root", tvwChild, "ea�o", "No hay actividades guardadas de " + materia + " para el mes de " + Combo6.List(Combo6.ListIndex) + " en ning�n a�o"
            End If
        End If
    End If
End Sub

Private Function decodificarMes(numMes As String) As Byte
    Dim mes As Byte
    Select Case LCase(numMes)
        Case "enero"
            mes = 1
        Case "febrero"
            mes = 2
        Case "marzo"
            mes = 3
        Case "abril"
            mes = 4
        Case "mayo"
            mes = 5
        Case "junio"
            mes = 6
        Case "julio"
            mes = 7
        Case "agosto"
            mes = 8
        Case "setiembre"
            mes = 9
        Case "octubre"
            mes = 10
        Case "noviembre"
            mes = 11
        Case "diciembre"
            mes = 12
    End Select
    decodificarMes = mes
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Call contarFormularios(False)
End Sub

Sub eliminarActividad(materia, mes As String, actividad As String)
    Kill App.path + "\trabajos\" + materia + "\actividades\" + mes + "\" + actividad
    Kill App.path + "\trabajos\" + materia + "\actividades\" + mes + "\datosActividades\" + Left(actividad, Len(actividad) - 4) + ".gui"
End Sub

