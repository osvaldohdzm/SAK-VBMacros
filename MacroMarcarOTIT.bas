Attribute VB_Name = "Module1"
Sub Clasificar_OT_IT_Avanzado()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colPuerto As Long, colServicio As Long, colOTIT As Long
    Dim i As Long
    Dim valPuerto As Variant
    Dim valServicio As String
    Dim esOT As Boolean
    Dim lngPuerto As Long
    
    ' --- CONFIGURACIÓN DE LISTAS ---
    
    ' 1. Servicios IT (PALABRAS DE VETO):
    ' Si el servicio contiene esto, SIEMPRE será IT, sin importar el puerto.
    Dim serviciosIT As Variant
    serviciosIT = Array("ssh", "sftp", "telnet", "smtp", "pop3", "imap", _
                        "domain", "dns", "ldap", "kerberos", "kpasswd", _
                        "netbios", "msrpc", "microsoft", "ms-", "rpc", _
                        "mysql", "sql", "vnc", "rdp", "adws", "exchange", _
                        "tomcat", "java", "oracle", "weblogic", "http-proxy")
                        ' Nota: "http" y "https" se manejan con cuidado abajo
                        
    ' 2. Servicios OT (PALABRAS CLAVE):
    ' Si el servicio contiene esto, se marca como OT.
    Dim serviciosOT As Variant
    serviciosOT = Array("dnp", "s7comm", "modbus", "abb-hw", "bacnet", _
                        "flexlm", "ansoft", "ansys", "cadlock", "ups", _
                        "zabbix", "patrol", "mqtt", "redis", "kyocera", _
                        "fins", "ethernet/ip", "scada", "plc", _
                        "fox", "niagara", "knx", "omron", "fanuc")

    ' --- INICIO DEL PROCESO ---
    
    On Error Resume Next
    Set tbl = ActiveSheet.ListObjects("Tbl_puertos")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "No se encontró la tabla 'Tbl_puertos'.", vbExclamation
        Exit Sub
    End If

    ' Mapear columnas
    On Error Resume Next
    colPuerto = tbl.ListColumns("Puerto").Index
    colServicio = tbl.ListColumns("Servicio").Index
    colOTIT = tbl.ListColumns("OT/IT").Index
    On Error GoTo 0

    If colPuerto = 0 Or colServicio = 0 Or colOTIT = 0 Then
        MsgBox "Faltan columnas requeridas.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Recorrer filas
    For i = 1 To tbl.ListRows.Count
        valPuerto = tbl.DataBodyRange(i, colPuerto).Value
        valServicio = LCase(CStr(tbl.DataBodyRange(i, colServicio).Value)) ' Minúsculas para comparar
        
        esOT = False ' Empezamos asumiendo IT
        
        ' ---------------------------------------------------------
        ' PASO 1: Clasificación por NÚMERO DE PUERTO (Base OT)
        ' ---------------------------------------------------------
        If IsNumeric(valPuerto) And Not IsEmpty(valPuerto) Then
            lngPuerto = CLng(valPuerto)
            Select Case lngPuerto
                ' Puertos exactos OT
                Case 102, 502, 20000, 47808, 1883, 8883, 6379, 10050, 1977, 9100, 2404
                    esOT = True
                ' Rangos de puertos (Ej. FlexLM 27000-27009)
                Case 27000 To 27009
                    esOT = True
            End Select
        End If

        ' ---------------------------------------------------------
        ' PASO 2: Clasificación por NOMBRE DE SERVICIO (Refuerzo OT)
        ' ---------------------------------------------------------
        Dim itemOT As Variant
        For Each itemOT In serviciosOT
            If InStr(1, valServicio, itemOT, vbTextCompare) > 0 Then
                esOT = True
                Exit For
            End If
        Next itemOT

        ' ---------------------------------------------------------
        ' PASO 3: EL VETO (Si es un servicio IT claro, se revoca OT)
        ' ---------------------------------------------------------
        ' Solo entramos aquí si actualmente está marcado como OT para verificar si es un falso positivo
        ' O si el servicio es HTTP/HTTPS genérico en un puerto no-OT
        
        Dim itemIT As Variant
        For Each itemIT In serviciosIT
            If InStr(1, valServicio, itemIT, vbTextCompare) > 0 Then
                esOT = False
                Exit For
            End If
        Next itemIT
        
        ' CASO ESPECIAL HTTP/HTTPS
        ' A veces OT usa http (ej. web de configuración de PLC).
        ' Lógica: Si dice "http" y NO ha sido marcado como OT por puerto o palabra clave OT, es IT.
        ' Si dice "http" pero el puerto es 80/443/8080 habituales, forzamos IT.
        If (InStr(valServicio, "http") > 0 Or InStr(valServicio, "https") > 0) Then
             ' Si el puerto es web estándar, es IT (salvo que ya hayamos detectado palabras OT como 'scada-web')
             If lngPuerto = 80 Or lngPuerto = 443 Or lngPuerto = 8080 Or lngPuerto = 8443 Then
                ' Chequeo extra: si dice "plc" o "scada" dentro del http, lo salvamos como OT
                If InStr(valServicio, "plc") = 0 And InStr(valServicio, "scada") = 0 Then
                    esOT = False
                End If
             End If
        End If

        ' ---------------------------------------------------------
        ' PASO 4: Escribir Resultado
        ' ---------------------------------------------------------
        If esOT Then
            tbl.DataBodyRange(i, colOTIT).Value = "OT"
        Else
            tbl.DataBodyRange(i, colOTIT).Value = "IT"
        End If
        
    Next i

    Application.ScreenUpdating = True
    MsgBox "Clasificación completada (Lógica Ampliada OT + Veto IT).", vbInformation
End Sub

