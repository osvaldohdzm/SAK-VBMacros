
$wordPath = "c:\Users\osvaldohm\Desktop\SAK-VBMacros\Word\Ciberseguridad\Modulo Ciberseguridad ESP-MX-CP1252 (Default).bas"
$excelPath = "c:\Users\osvaldohm\Desktop\SAK-VBMacros\Excel\Ciberseguridad\Módulo de Ciberseguridad ESP-MX-CP1252 (Default).bas"

# Function to perform replacement
function Rename-Macros {
    param(
        [string]$filePath,
        [string[]]$subsList
    )

    Write-Host "Processing $filePath..."
    
    # Read with CP1252 to preserve formatting
    $encoding = [System.Text.Encoding]::GetEncoding(1252)
    $content = [System.IO.File]::ReadAllText($filePath, $encoding)

    $counter = 1

    foreach ($oldName in $subsList) {
        # Clean the name: remove existing prefix if any, replace Vi?etas
        # Patterns to remove: CYB000_, CYB_000_, CYB123_
        $cleanName = $oldName -replace "^CYB_?\d+_", ""
        
        # Replace Vi?etas with Vinetas (handling bad encoding chars if possible, assuming input string names are correct from file)
        # Note: The $oldName comes from the file content we read previously, so it matches exactly.
        # But we want the NEW name to have "Vinetas".
        
        $cleanName = $cleanName -replace "Vi.etas", "Vinetas" 
        $cleanName = $cleanName -replace "Viñetas", "Vinetas"

        # Format new name
        $prefix = "CYB_" + "{0:D3}" -f $counter
        $newName = $prefix + "_" + $cleanName

        Write-Host "Renaming $oldName -> $newName"

        # Escape for Regex
        $escapedOld = [regex]::Escape($oldName)
        
        # Regex replace for whole word match to avoid partial replacements
        # We use CaseInsensitive because VBA is case insensitive
        $pattern = "\b$escapedOld\b"
        
        $content = [regex]::Replace($content, $pattern, $newName, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        
        $counter++
    }

    # Final "Vi?etas" cleanup in the whole file if requested by user for other occurrences?
    # "aplica los cambios solamente sustituye la palabra Vi?etas por Vinetas"
    # This might apply to strings too.
    $content = $content -replace "Vi\?etas", "Vinetas"
    $content = $content -replace "Viñetas", "Vinetas"

    # Write back
    [System.IO.File]::WriteAllText($filePath, $content, $encoding)
    Write-Host "Done with $filePath"
}

# Word Subs List
$wordSubs = @(
"AutoOpen",
"AutoClose",
"AddSeverityContextMenu",
"RemoveSeverityContextMenu",
"ColorearCritica",
"ColorearAlta",
"ColorearMedia",
"ColorearBaja",
"QuitarColor",
"FormatearTablaVulnerabilidadesAvanzado",
"NegritaPalabrasClave_Robusta_MultiArray_Corregido_CVECompleto",
"AjustarFormatoColumnasTablaVulnes",
"NegritaPalabrasClave",
"InsertarBloqueCodigoFormateado",
"ColorearCeldaSeveridad",
"FormatearCodigoEnRango",
"FormatKeywords",
"FormatPattern",
"FormatWithWildcards",
"FormatearCodigoHTMLSeleccionado",
"PalabrasClaveVerde_Corregida",
"CensurarIPs_X_Dinamica_Segura"
)

# Excel Subs List
$excelSubs = @(
"ApplyParagraphFormattingToCell",
"CYB008_LimpiarTextoYAgregarGuion",
"CYB020_ExportarHojaConFormatoINAI",
"CYB032_ReemplazarCadenasSeveridades",
"CYB024_LimpiarCeldasYMostrarContenidoComoArray",
"CYB029_ReemplazarConURLs",
"CYB037_AplicarFormatoCondicional",
"CYB033_ConvertirATextoEnOracion",
"CYB027_QuitarEspacios",
"CYB009_ProcesadoCompletoSalidaHerramientas",
"CYB009_LimpiarSalida",
"CYB010_AgregarSaltosLineaATextoGuiones",
"CYB011_BulletsAGuiones",
"CYB011_MantererSoloURLSEnLinea",
"CYB012_PingIPs",
"CYB012_ObtenerIPs",
"CYB013_ReverseDNS",
"CYB014_CheckHTTPHTTPS",
"CYB014_EscanearIPsDesdeSeleccionSinDuplicados",
"CYB026_OrdenaSegunColorRelleno",
"EliminarUltimasFilasSiEsSalidaPruebaSeguridad",
"CYB015_CrearEstilo",
"CYB038_ExportarTablaContenidoADocumentoWord",
"CYB017_LimpiarColumnaReferencias",
"CYB018_LeerArchivoTXT",
"CYB019_WordAppAlternativeReplaceParagraph",
"CYB001_GenerarDocumentosVulnerabilidadesWord",
"CYB001_GenerarDocumentosVulnerabilidadesWordAgrupadoActivo",
"CYB007_GenerarReportesVulnsAppsINAI",
"CYB003_GenerarReportesVulns",
"CYB039_KillAllWordInstances",
"ReemplazarCampos",
"ActualizarGraficos",
"FormatearParrafosGuionesCelda",
"AplicarNegritaPalabrasClaveEnCeldaWord",
"FormatearCeldaNivelRiesgo",
"WordAppReemplazarParrafo",
"CYB013_DesglosarIPs",
"CYB034_CargarResultados_DatosDesdeCSVNessus",
"CYB035_CargarResultados_DatosDesdeCSVNexPose",
"CYB036_CargarResultados_DatosDesdeXMLOpenVAS",
"CYB037_CargarResultados_DatosDesdeCSVAcunetix",
"CheckColumnExists",
"CYB040_ResaltarFalsosPositivosEnVerde",
"CYB036_CargarResultados_DatosDesdeXMLAcunetix_v5",
"CYB041_IrACatalogoVulnerabilidad",
"CYB042_MarcarMultiplesEnCatalogoVulnerabilidad",
"CYB042_Estandarizar",
"CYB043_AplicarFormatoCondicional",
"CYB061_LLM_llama3_2_1b",
"CYB060_LLLM_deepseek_r1_1_5b",
"ObtenerRespuestasGeminiCVSS4",
"CYB068_PrepararPromptDesdeSeleccion_DescripcionVuln_General",
"CYB069_PrepararPromptDesdeSeleccion_AmenazaVuln_General",
"CYB070_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_General",
"CYB071_PrepararPromptDesdeSeleccion_DescripcionVuln_EnVPN",
"CYB072_PrepararPromptDesdeSeleccion_AmenazaVuln_EnVPN",
"CYB073_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_EnVPN",
"CYB074_PrepararPromptDesdeSeleccion_DescripcionVuln_EnRedPrivada",
"CYB075_PrepararPromptDesdeSeleccion_AmenazaVuln_EnRedPrivada",
"CYB076_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_EnRedPrivada",
"CYB0761_PrepararPromptDesdeSeleccion_ExplicacionTecnicaVuln_EnRedPrivada",
"CYB0762_PrepararPromptDesdeSeleccion_VectorCVSSVuln_EnRedPrivada",
"CYB077_PrepararPromptDesdeSeleccion_DescripcionVuln_DesdeInternet",
"CYB078_PrepararPromptDesdeSeleccion_AmenazaVuln_DesdeInternet",
"CYB079_PrepararPromptDesdeSeleccion_PropuestaRemediacionVuln_DesdeInternet",
"CYB080_PreparePromptFromSelection_DescripcionVuln_FromCode",
"CYB081_PreparePromptFromSelection_AmenazaVuln_FromCode",
"CYB082_PreparePromptFromSelection_PropuestaRemediacionVuln_FromCode",
"CYB082_PreparePromptFromSelection_MetodoDeteccion",
"CYB083_Verificar_VectorCVSS4_0",
"CopiarAlPortapapeles",
"InsertarTextoMarkdownEnWordConFormato",
"ProcessInlineFormatting",
"FusionarDocumentosInsertando",
"RawPrint",
"SustituirTextoMarkdownPorImagenes",
"AplicarFormatoCeldaEnTablaWord",
"EliminarLineasVaciasEnCeldaTablaWord",
"AjustarMarcadorCeldaEnTablaWord"
)

Try {
    Rename-Macros -filePath $wordPath -subsList $wordSubs
    Rename-Macros -filePath $excelPath -subsList $excelSubs
} Catch {
    Write-Host "Error: $_"
}
