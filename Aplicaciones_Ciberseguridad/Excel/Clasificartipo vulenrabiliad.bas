Attribute VB_Name = "M�dulo1"
Sub AsignarTipoVulnerabilidad()

Dim rngSeleccionado As Range
    Dim rngColumnaDerecha As Range
    Dim rngColumnaIzquierda As Range
    Dim rngColumnaInicial As Range
        Dim celda As Range
    Dim tipo As String
    Dim categorias As Object
    Set categorias = CreateObject("Scripting.Dictionary")
    
    ' Verificar si se ha seleccionado un rango de celdas
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas.", vbExclamation
        Exit Sub
    End If
    
    ' Guardar la selecci�n original
    Set rngSeleccionado = Selection
    
    ' Verificar si la selecci�n tiene solo una columna
    If rngSeleccionado.Columns.Count <> 1 Then
        MsgBox "Por favor, selecciona un rango que tenga solo una columna.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener la columna a la derecha del rango seleccionado
    Set rngColumnaDerecha = rngSeleccionado.Offset(0, 1).EntireColumn
    
    ' Insertar una columna a la izquierda de la columna a la derecha
    rngColumnaDerecha.Insert Shift:=xlToLeft
    
    ' Obtener la columna inicial del rango seleccionado
    Set rngColumnaInicial = rngSeleccionado.EntireColumn
    
    ' Seleccionar nuevamente el rango original
    rngSeleccionado.Select

    
    ' Mapea los t�rminos de b�squeda a los tipos de vulnerabilidad
categorias("antimalware desactualizado") = "Antimalware desactualizado"
categorias("apache activemq") = "Versi�n desactualizada de software"
categorias("apache subversion client") = "Versi�n desactualizada de software"
categorias("apache subversion server") = "Versi�n desactualizada de software"
categorias("apache tomcat") = "Versi�n desactualizada de software"
categorias("winrar") = "Versi�n desactualizada de software"
categorias("sqli scanner") = "Versi�n desactualizada de software"
categorias("authentication bypass") = "Configuraci�n insegura"
categorias("authentication vulnerability") = "Configuraci�n insegura"
categorias("cve") = "Configuraci�n insegura"
categorias("dos") = "Configuraci�n insegura"
categorias("cgi generic local file inclusion") = "Configuraci�n insegura"
categorias("manageengine adaudit plus") = "Versi�n desactualizada de software"
categorias("netscaler unencrypted web management interface") = "Configuraci�n insegura"
categorias("putty") = "Versi�n desactualizada de software"
categorias("rhel 5 :") = "Versi�n desactualizada de software"
categorias("rhel 6 :") = "Versi�n desactualizada de software"
categorias("rhel 7 :") = "Versi�n desactualizada de software"
categorias("generic local file inclusion") = "Configuraci�n insegura"
categorias("edge chromium") = "Ausencia de parches de seguridad"
categorias("google chrome") = "Versi�n desactualizada de software"
categorias("http request smuggling") = "Configuraci�n insegura"
categorias("http response splitting") = "Configuraci�n insegura"
categorias("information disclosure") = "Configuraci�n insegura"
categorias("kernel") = "Versi�n desactualizada de sistema operativo"
categorias("kibana") = "Versi�n desactualizada de software"
categorias("linux") = "Versi�n desactualizada de sistema operativo"
categorias("mozilla firefox") = "Versi�n desactualizada de software"
categorias("mozilla") = "Versi�n desactualizada de software"
categorias("multiple vulnerabilities") = "Versi�n desactualizada de software"
categorias("oracle coherence") = "Versi�n desactualizada de software"
categorias("oracle database server") = "Versi�n desactualizada de software"
categorias("oracle java") = "Versi�n desactualizada de software"
categorias("oracle mysql connectors") = "Versi�n desactualizada de software"
categorias("oracle weblogic server") = "Versi�n desactualizada de software"
categorias("rhel 5:") = "Versi�n desactualizada de software"
categorias("rhel 6:") = "Versi�n desactualizada de software"
categorias("privilege escalation") = "Configuraci�n insegura"
categorias("rce") = "Configuraci�n insegura"
categorias("remote code execution") = "Configuraci�n insegura"
categorias("security update") = "Ausencia de parches de seguridad"
categorias("sql injection") = "Configuraci�n insegura"
categorias("unsupported os") = "Sistema operativo sin soporte"
categorias("unsupported software") = "Versi�n sin soporte"
categorias("unsupported version") = "Versi�n sin soporte"
categorias("vmware tools") = "Versi�n desactualizada de software"
categorias("vulnerability") = "Ausencia de parches de seguridad"
categorias("windows 10") = "Versi�n desactualizada de sistema operativo"
categorias("windows server 2008 r2") = "Versi�n desactualizada de sistema operativo"
categorias("windows server 2008") = "Versi�n desactualizada de sistema operativo"
categorias("windows server 2012 r2") = "Versi�n desactualizada de sistema operativo"
categorias("windows server 2012") = "Versi�n desactualizada de sistema operativo"
categorias("windows server 2016") = "Versi�n desactualizada de sistema operativo"
categorias("windows server 2019") = "Versi�n desactualizada de sistema operativo"
categorias("windows server") = "Versi�n desactualizada de sistema operativo"
categorias("xss") = "Configuraci�n insegura"
    
    
    
    ' Recorre cada celda en las celdas seleccionadas actualmente
    For Each celda In rngSeleccionado
        tipo = "No identificado" ' Establece un valor predeterminado
        
        ' Verifica el contenido de la celda y asigna el tipo de vulnerabilidad correspondiente
        For Each Key In categorias
            If InStr(1, LCase(celda.Value), Key) > 0 Then
                tipo = categorias(Key)
                Exit For ' Una vez que se encuentra una coincidencia, sal del bucle
            End If
        Next Key
        
       
        ' Escribe el tipo de vulnerabilidad en la celda adyacente
        celda.Offset(0, 1).Value = tipo
    Next celda
End Sub
