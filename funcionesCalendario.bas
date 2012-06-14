Attribute VB_Name = "funcionesCalendario"
Option Explicit

Public Function cantD�asMes(qu�Mes As Byte, qu�A�o As Integer) As Byte
    Select Case qu�Mes
        Case 1 'enero
            cantD�asMes = 31
        Case 2 'febrero
            cantD�asMes = 28
        Case 3
            cantD�asMes = 31
        Case 4
            cantD�asMes = 30
        Case 5
            cantD�asMes = 31
        Case 6
            cantD�asMes = 30
        Case 7
            cantD�asMes = 31
        Case 8
            cantD�asMes = 31
        Case 9
            cantD�asMes = 30
        Case 10
            cantD�asMes = 31
        Case 11
            cantD�asMes = 30
        Case 12 'diciembre
            cantD�asMes = 31
    End Select
    If a�oBisiesto(qu�A�o) And qu�Mes = 2 Then cantD�asMes = cantD�asMes + 1 'si el a�o es bisiesto le sumamos el 29 de febrero
End Function

Public Function a�osBisiestosHastaD�a(a�o As Integer) As Integer
    Dim a�osPasadosDesde2008 As Integer ', cantA�osBisiestos As Integer
    
    If a�o > 2008 Then
        a�osPasadosDesde2008 = a�o - 2008
        If a�osPasadosDesde2008 > 4 Then 'si ya hay un bisiesto desde 2008
            a�osBisiestosHastaD�a = Int(a�osPasadosDesde2008 / 4)
        Else 'si todav�a es el siguiente a�o bisiesto desde 2008
            a�osBisiestosHastaD�a = 0
        End If
    End If
End Function

Public Function d�asPasadosDesde1Feb2008(d�a As Byte, mes As Byte, a�o As Integer) As Long
    Dim result, i As Integer
    
    If a�o > 2008 Then
        
        For i = 2 To 12 'se suman los d�as desde febrero de 2008 hasta finales de 2008
            d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 + cantD�asMes(CInt(i), 2008)
        Next
        
        result = a�o - 2008
        d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 + ((result - 1) * 365) 'se suman los d�as de todos los a�os completos
        d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 + d�a 'se le suman los d�as pasados
        
        If mes > 1 Then
            For i = 1 To mes - 1 'se suman los d�as de cada mes completo pasado desde principio de a�o
                d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 + cantD�asMes(CInt(i), a�o)
            Next
        End If
        
        If result > 4 Then 'si ya han pasado a�os bisiestos, se los suma como un d�a m�s por a�o
            d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 + a�osBisiestosHastaD�a(a�o)
            If a�oBisiesto(a�o) Then d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 - 1 'si el propio a�o es bisiesto, se lo resta porque el d�a 29 ya se sum� en el mes
        End If
        
        d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 - 1 'se le resta un d�a pues se empieza contando tambi�n el primero de febrero
    ElseIf a�o = 2008 Then
        For i = 2 To mes - 1 'se suman los d�as desde febrero de 2008 hasta finales de 2008
            d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 + cantD�asMes(CInt(i), 2008)
        Next
        d�asPasadosDesde1Feb2008 = d�asPasadosDesde1Feb2008 + d�a - 1
    Else
'        MsgBox "No se pueden cargar a�os menores a 2008"
        frmMsgBox.cadenaAMostrar = "No se pueden cargar a�os menores a 2008"
        frmMsgBox.swS�No�Aceptar = False 'se elige que sea cuadro aceptar
        frmMsgBox.Show 1
    End If
End Function

Public Function a�oBisiesto(a�o As Integer) As Boolean
    Dim a�osDesde2008 As Integer
    If a�o > 2008 Then
        a�osDesde2008 = a�o - 2008
        If 0 = a�osDesde2008 Mod 4 Then
            a�oBisiesto = True
        Else
            a�oBisiesto = False
        End If
    End If
    
    If a�o = 2008 Then a�oBisiesto = True
End Function

Public Function nombreDeD�a(d�a As Byte, mes As Byte, a�o As Integer) As Byte
    Dim cu�ntosD�asPasaron As Long, n�mD�a As Byte
    cu�ntosD�asPasaron = d�asPasadosDesde1Feb2008(d�a, mes, a�o)
    'nombreDeD�a = cu�ntosD�asPasaron Mod 7
    
    If 0 = cu�ntosD�asPasaron Mod 7 Then nombreDeD�a = 5 'viernes
    If 0 = (cu�ntosD�asPasaron + 1) Mod 7 Then nombreDeD�a = 4 'jueves
    If 0 = (cu�ntosD�asPasaron + 2) Mod 7 Then nombreDeD�a = 3 'mi�rcoles
    If 0 = (cu�ntosD�asPasaron + 3) Mod 7 Then nombreDeD�a = 2 'martes
    If 0 = (cu�ntosD�asPasaron + 4) Mod 7 Then nombreDeD�a = 1 'lunes
    If 0 = (cu�ntosD�asPasaron - 1) Mod 7 Then nombreDeD�a = 6 's�bado
    If 0 = (cu�ntosD�asPasaron - 2) Mod 7 Then nombreDeD�a = 7 'domingo
End Function

