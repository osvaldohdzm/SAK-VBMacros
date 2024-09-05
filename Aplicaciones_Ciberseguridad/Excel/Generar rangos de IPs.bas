Attribute VB_Name = "Módulo1"
Option Explicit

Sub GenerarIPs()
    Dim segment As String
    Dim startIP As String
    Dim endIP As String
    Dim i As Long, j As Long
    Dim baseIP As String
    Dim ipParts() As String
    Dim startParts() As String
    Dim endParts() As String
    Dim currentIP As String
    Dim rng As Range

    ' Solicitar al usuario que introduzca el segmento de IP
    segment = InputBox("Introduzca el segmento de IP (formato CIDR o rango):", "Generar IPs")
    If segment = "" Then Exit Sub ' Salir si el usuario cancela la entrada

    ' Usar la celda actualmente seleccionada como celda de inicio
    Set rng = Application.Selection

    If InStr(segment, "/") > 0 Then
        ' Procesar notación CIDR
        baseIP = Split(segment, "/")(0)
        Dim subnet As Integer
        subnet = CInt(Split(segment, "/")(1))
        
        ipParts = Split(baseIP, ".")
        Dim ipNum As Double
        ipNum = (CDbl(ipParts(0)) * 256 ^ 3) + (CDbl(ipParts(1)) * 256 ^ 2) + (CDbl(ipParts(2)) * 256) + CDbl(ipParts(3))
        Dim numHosts As Double
        numHosts = 2 ^ (32 - subnet)
        
        For i = 0 To numHosts - 1
            Dim newIPNum As Double
            newIPNum = ipNum + i
            Dim octet1 As Long, octet2 As Long, octet3 As Long, octet4 As Long
            octet1 = Int(newIPNum / (256 ^ 3)) Mod 256
            octet2 = Int(newIPNum / (256 ^ 2)) Mod 256
            octet3 = Int(newIPNum / 256) Mod 256
            octet4 = newIPNum Mod 256
            Dim newIP As String
            newIP = octet1 & "." & octet2 & "." & octet3 & "." & octet4
            rng.Offset(i, 0).Value = newIP
        Next i
    ElseIf InStr(segment, "-") > 0 Then
        ' Procesar notación de rango
        startIP = Split(segment, "-")(0)
        endIP = Split(segment, "-")(1)
        
        startParts = Split(startIP, ".")
        If UBound(Split(endIP, ".")) = 0 Then
            endParts = startParts
            endParts(3) = endIP
        Else
            endParts = Split(endIP, ".")
        End If
        
        Dim startRange As Integer
        Dim endRange As Integer
        
        startRange = CInt(startParts(3))
        endRange = CInt(endParts(3))
        
        baseIP = startParts(0) & "." & startParts(1) & "." & startParts(2) & "."
        
        For j = startRange To endRange
            currentIP = baseIP & j
            rng.Offset(j - startRange, 0).Value = currentIP
        Next j
    Else
        MsgBox "Formato no reconocido: " & segment, vbExclamation
    End If
End Sub

