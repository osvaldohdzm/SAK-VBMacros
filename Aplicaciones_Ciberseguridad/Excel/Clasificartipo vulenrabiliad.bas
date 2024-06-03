Attribute VB_Name = "Módulo1"
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
    
    ' Guardar la selección original
    Set rngSeleccionado = Selection
    
    ' Verificar si la selección tiene solo una columna
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

    
    ' Mapea los términos de búsqueda a los tipos de vulnerabilidad
categorias("antimalware desactualizado") = "Antimalware desactualizado"
categorias("apache activemq") = "Versión desactualizada de software"
categorias("apache subversion client") = "Versión desactualizada de software"
categorias("apache subversion server") = "Versión desactualizada de software"
categorias("apache tomcat") = "Versión desactualizada de software"
categorias("winrar") = "Versión desactualizada de software"
categorias("sqli scanner") = "Versión desactualizada de software"
categorias("authentication bypass") = "Configuración insegura"
categorias("authentication vulnerability") = "Configuración insegura"
categorias("cve") = "Configuración insegura"
categorias("dos") = "Configuración insegura"
categorias("cgi generic local file inclusion") = "Configuración insegura"
categorias("manageengine adaudit plus") = "Versión desactualizada de software"
categorias("netscaler unencrypted web management interface") = "Configuración insegura"
categorias("putty") = "Versión desactualizada de software"
categorias("rhel 5 :") = "Versión desactualizada de software"
categorias("rhel 6 :") = "Versión desactualizada de software"
categorias("rhel 7 :") = "Versión desactualizada de software"
categorias("generic local file inclusion") = "Configuración insegura"
categorias("edge chromium") = "Ausencia de parches de seguridad"
categorias("google chrome") = "Versión desactualizada de software"
categorias("http request smuggling") = "Configuración insegura"
categorias("http response splitting") = "Configuración insegura"
categorias("information disclosure") = "Configuración insegura"
categorias("kernel") = "Versión desactualizada de sistema operativo"
categorias("kibana") = "Versión desactualizada de software"
categorias("linux") = "Versión desactualizada de sistema operativo"
categorias("mozilla firefox") = "Versión desactualizada de software"
categorias("mozilla") = "Versión desactualizada de software"
categorias("multiple vulnerabilities") = "Versión desactualizada de software"
categorias("oracle coherence") = "Versión desactualizada de software"
categorias("oracle database server") = "Versión desactualizada de software"
categorias("oracle java") = "Versión desactualizada de software"
categorias("oracle mysql connectors") = "Versión desactualizada de software"
categorias("oracle weblogic server") = "Versión desactualizada de software"
categorias("rhel 5:") = "Versión desactualizada de software"
categorias("rhel 6:") = "Versión desactualizada de software"
categorias("privilege escalation") = "Configuración insegura"
categorias("rce") = "Configuración insegura"
categorias("remote code execution") = "Configuración insegura"
categorias("security update") = "Ausencia de parches de seguridad"
categorias("sql injection") = "Configuración insegura"
categorias("unsupported os") = "Sistema operativo sin soporte"
categorias("unsupported software") = "Versión sin soporte"
categorias("unsupported version") = "Versión sin soporte"
categorias("vmware tools") = "Versión desactualizada de software"
categorias("vulnerability") = "Ausencia de parches de seguridad"
categorias("windows 10") = "Versión desactualizada de sistema operativo"
categorias("windows server 2008 r2") = "Versión desactualizada de sistema operativo"
categorias("windows server 2008") = "Versión desactualizada de sistema operativo"
categorias("windows server 2012 r2") = "Versión desactualizada de sistema operativo"
categorias("windows server 2012") = "Versión desactualizada de sistema operativo"
categorias("windows server 2016") = "Versión desactualizada de sistema operativo"
categorias("windows server 2019") = "Versión desactualizada de sistema operativo"
categorias("windows server") = "Versión desactualizada de sistema operativo"
categorias("xss") = "Configuración insegura"
    
    
    
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
