Attribute VB_Name = "funcionesCalendario"
Option Explicit

Public Function cantDíasMes(quéMes As Byte, quéAño As Integer) As Byte
    Select Case quéMes
        Case 1 'enero
            cantDíasMes = 31
        Case 2 'febrero
            cantDíasMes = 28
        Case 3
            cantDíasMes = 31
        Case 4
            cantDíasMes = 30
        Case 5
            cantDíasMes = 31
        Case 6
            cantDíasMes = 30
        Case 7
            cantDíasMes = 31
        Case 8
            cantDíasMes = 31
        Case 9
            cantDíasMes = 30
        Case 10
            cantDíasMes = 31
        Case 11
            cantDíasMes = 30
        Case 12 'diciembre
            cantDíasMes = 31
    End Select
    If añoBisiesto(quéAño) And quéMes = 2 Then cantDíasMes = cantDíasMes + 1 'si el año es bisiesto le sumamos el 29 de febrero
End Function

Public Function añosBisiestosHastaDía(año As Integer) As Integer
    Dim añosPasadosDesde2008 As Integer ', cantAñosBisiestos As Integer
    
    If año > 2008 Then
        añosPasadosDesde2008 = año - 2008
        If añosPasadosDesde2008 > 4 Then 'si ya hay un bisiesto desde 2008
            añosBisiestosHastaDía = Int(añosPasadosDesde2008 / 4)
        Else 'si todavía es el siguiente año bisiesto desde 2008
            añosBisiestosHastaDía = 0
        End If
    End If
End Function

Public Function díasPasadosDesde1Feb2008(día As Byte, mes As Byte, año As Integer) As Long
    Dim result, i As Integer
    
    If año > 2008 Then
        
        For i = 2 To 12 'se suman los días desde febrero de 2008 hasta finales de 2008
            díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 + cantDíasMes(CInt(i), 2008)
        Next
        
        result = año - 2008
        díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 + ((result - 1) * 365) 'se suman los días de todos los años completos
        díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 + día 'se le suman los días pasados
        
        If mes > 1 Then
            For i = 1 To mes - 1 'se suman los días de cada mes completo pasado desde principio de año
                díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 + cantDíasMes(CInt(i), año)
            Next
        End If
        
        If result > 4 Then 'si ya han pasado años bisiestos, se los suma como un día más por año
            díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 + añosBisiestosHastaDía(año)
            If añoBisiesto(año) Then díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 - 1 'si el propio año es bisiesto, se lo resta porque el día 29 ya se sumó en el mes
        End If
        
        díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 - 1 'se le resta un día pues se empieza contando también el primero de febrero
    ElseIf año = 2008 Then
        For i = 2 To mes - 1 'se suman los días desde febrero de 2008 hasta finales de 2008
            díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 + cantDíasMes(CInt(i), 2008)
        Next
        díasPasadosDesde1Feb2008 = díasPasadosDesde1Feb2008 + día - 1
    Else
'        MsgBox "No se pueden cargar años menores a 2008"
        frmMsgBox.cadenaAMostrar = "No se pueden cargar años menores a 2008"
        frmMsgBox.swSíNoóAceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
    End If
End Function

Public Function añoBisiesto(año As Integer) As Boolean
    Dim añosDesde2008 As Integer
    If año > 2008 Then
        añosDesde2008 = año - 2008
        If 0 = añosDesde2008 Mod 4 Then
            añoBisiesto = True
        Else
            añoBisiesto = False
        End If
    End If
    
    If año = 2008 Then añoBisiesto = True
End Function

Public Function nombreDeDía(día As Byte, mes As Byte, año As Integer) As Byte
    Dim cuántosDíasPasaron As Long, númDía As Byte
    cuántosDíasPasaron = díasPasadosDesde1Feb2008(día, mes, año)
    'nombreDeDía = cuántosDíasPasaron Mod 7
    
    If 0 = cuántosDíasPasaron Mod 7 Then nombreDeDía = 5 'viernes
    If 0 = (cuántosDíasPasaron + 1) Mod 7 Then nombreDeDía = 4 'jueves
    If 0 = (cuántosDíasPasaron + 2) Mod 7 Then nombreDeDía = 3 'miércoles
    If 0 = (cuántosDíasPasaron + 3) Mod 7 Then nombreDeDía = 2 'martes
    If 0 = (cuántosDíasPasaron + 4) Mod 7 Then nombreDeDía = 1 'lunes
    If 0 = (cuántosDíasPasaron - 1) Mod 7 Then nombreDeDía = 6 'sábado
    If 0 = (cuántosDíasPasaron - 2) Mod 7 Then nombreDeDía = 7 'domingo
End Function

