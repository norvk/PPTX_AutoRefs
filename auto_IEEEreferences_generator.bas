Attribute VB_Name = "Module2"
Option Explicit

' --------------------------------------------------------------------
' Creado por Benjamín Moya Giachetti
' Version 1.0
'
' BREVE DESCRIPCIÓN:
'
' Esta macro VBA para PowerPoint hace lo siguiente:
' 1) Recorre todas las diapositivas y cada \cite{key} se sustituye por \cite{key}[n],
'    donde [n] es un número consecutivo según el orden de aparición.
' 2) Oculta (color y tamaño) la parte "\cite{key}" dejando visible solo [n].
' 3) Carga un archivo references.bib en la misma carpeta que el PPT.
' 4) Genera una diapositiva final llamada "Bibliografía" enumerando las referencias
'    [1], [2]... en orden de aparición.
' 5) Incluye un mecanismo de reemplazo de tildes LaTeX (sin diéresis) para que
'    en la Bibliografía no aparezca {\'i}, {\~n}, etc.
'
' USO:
' - Coloca references.bib junto a tu presentación.
' - Inserta este módulo en tu PPTM.
' - Ejecuta GenerarReferencias().
' --------------------------------------------------------------------

Public Sub GenerarReferencias()
    Dim refsMap As Object               ' Dictionary: key -> "[n]"
    Dim referencesData As Object        ' Dictionary: key -> contenido .bib
    Dim refCount As Long: refCount = 0
    
    Set refsMap = CreateObject("Scripting.Dictionary")
    
    ' (1) Primer pase: encontrar \cite{key}, asignar [n], reemplazar brackets antiguos
    Dim sld As slide, shp As shape
    For Each sld In ActivePresentation.slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    Dim originalText As String
                    originalText = shp.TextFrame.textRange.text
                    
                    ' Corrige tildes LaTeX en el texto
                    originalText = ReemplazarTildesLatex(originalText)
                    
                    ' Encontrar \cite{...}
                    Dim foundKeys As Object
                    Set foundKeys = EncontrarCites(originalText)
                    
                    If foundKeys.Count > 0 Then
                        Dim k As Variant
                        For Each k In foundKeys.Keys
                            If Not refsMap.Exists(k) Then
                                refCount = refCount + 1
                                refsMap.Add k, "[" & refCount & "]"
                            End If
                        Next k
                    End If
                    
                    ' Reemplazar (o actualizar) [m] -> [n]
                    shp.TextFrame.textRange.text = ReemplazarBrackets(originalText, refsMap)
                End If
            End If
        Next shp
    Next sld
    
    ' (2) Segundo pase: ocultar \cite{...} (color blanco, tamaño 1)
    For Each sld In ActivePresentation.slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    Call OcultarCites(shp)
                End If
            End If
        Next shp
    Next sld
    
    ' (3) Cargar references.bib
    Set referencesData = CargarBibEnDiccionario()
    
    ' (4) Generar la diapositiva final con la bibliografía
    If Not referencesData Is Nothing Then
        GenerarDiapositivaBibliografia refsMap, referencesData
    Else
        MsgBox "No se encontró references.bib. No se generará la diapositiva de Bibliografía.", vbExclamation
    End If
    
    MsgBox "Macro completada. Se asignaron " & refCount & " referencias y se generó la Bibliografía.", vbInformation
End Sub

' --------------------------------------------------------------------
' Busca secuencias \cite{key} sin bracket.
' Retorna Dictionary con cada key encontrada.
' --------------------------------------------------------------------
Private Function EncontrarCites(ByVal txt As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.pattern = "\\cite\{([^}]+)\}"
    regEx.Global = True
    
    Dim matches As Object, m As Object
    Set matches = regEx.Execute(txt)
    
    For Each m In matches
        Dim theKey As String
        theKey = m.SubMatches(0)
        If Not dict.Exists(theKey) Then
            dict.Add theKey, True
        End If
    Next m
    
    Set EncontrarCites = dict
End Function

' --------------------------------------------------------------------
' Reemplaza secuencias \cite{key}(\[\d+\])? por \cite{key}[n].
' Si había bracket [m], lo actualiza para que no se duplique.
' --------------------------------------------------------------------
Private Function ReemplazarBrackets(ByVal txt As String, ByVal refsMap As Object) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.pattern = "(\\cite\{([^}]+)\})(\[\d+\])?"
    regEx.Global = True
    
    Dim matches As Object
    Set matches = regEx.Execute(txt)
    If matches.Count = 0 Then
        ReemplazarBrackets = txt
        Exit Function
    End If
    
    Dim result As String: result = ""
    Dim lastPos As Long: lastPos = 1
    
    Dim m As Object
    For Each m In matches
        Dim matchStart As Long
        matchStart = m.FirstIndex + 1 ' +1 (base1)
        Dim matchLen As Long
        matchLen = m.Length
        
        ' Texto anterior
        result = result & Mid(txt, lastPos, matchStart - lastPos)
        
        Dim theKey As String
        theKey = m.SubMatches(1) ' subMatches(0) = "\cite{key}", subMatches(1) = "key"
        
        Dim bracketStr As String
        If refsMap.Exists(theKey) Then
            bracketStr = refsMap(theKey)  ' p.ej "[2]"
        Else
            bracketStr = "[?]"
        End If
        
        ' Reconstruct \cite{key}[n]
        Dim replacedSegment As String
        replacedSegment = m.SubMatches(0) & bracketStr
        
        result = result & replacedSegment
        
        lastPos = matchStart + matchLen
    Next m
    
    ' Resto del texto
    result = result & Mid(txt, lastPos)
    
    ReemplazarBrackets = result
End Function

' --------------------------------------------------------------------
' Colorea \cite{...} de blanco y tamaño 1, para que no se vea.
' --------------------------------------------------------------------
Private Sub OcultarCites(ByVal shp As shape)
    Dim txtRange As textRange
    Set txtRange = shp.TextFrame.textRange
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.pattern = "\\cite\{[^}]+\}"
    regEx.Global = True
    
    Dim allText As String
    allText = txtRange.text
    
    Dim matches As Object, m As Object
    Set matches = regEx.Execute(allText)
    
    For Each m In matches
        Dim startPos As Long, lengthMatch As Long
        startPos = m.FirstIndex + 1 ' base1
        lengthMatch = m.Length
        
        If startPos > 0 And lengthMatch > 0 Then
            Dim citeRange As textRange
            Set citeRange = txtRange.Characters(startPos, lengthMatch)
            
            With citeRange.Font
                .Color.RGB = vbWhite ' asumiendo fondo blanco
                .Size = 1
            End With
        End If
    Next m
End Sub

' --------------------------------------------------------------------
' Carga references.bib en un diccionario key -> contenido crudo
' --------------------------------------------------------------------
Private Function CargarBibEnDiccionario() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object, bibFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim bibPath As String
    bibPath = ActivePresentation.Path & "\references.bib"
    
    If Not fso.FileExists(bibPath) Then
        Set CargarBibEnDiccionario = Nothing
        Exit Function
    End If
    
    Set bibFile = fso.OpenTextFile(bibPath, 1, False)
    
    Dim currentKey As String: currentKey = ""
    Dim accumulator As String: accumulator = ""
    
    Do Until bibFile.AtEndOfStream
        Dim line As String
        line = bibFile.ReadLine
        
        If Left(line, 1) = "@" Then
            If currentKey <> "" Then
                dict.Add currentKey, accumulator
            End If
            
            currentKey = ExtraerBibKey(line)
            accumulator = line & vbCrLf
        Else
            If currentKey <> "" Then
                accumulator = accumulator & line & vbCrLf
            End If
        End If
    Loop
    
    If currentKey <> "" Then
        dict.Add currentKey, accumulator
    End If
    
    bibFile.Close
    
    Set CargarBibEnDiccionario = dict
End Function

' --------------------------------------------------------------------
' Extrae la clave de una línea tipo: @article{Smith2020,
' --------------------------------------------------------------------
Private Function ExtraerBibKey(ByVal line As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.pattern = "@[a-zA-Z]+{\s*([^,]+),"
    regEx.Global = False
    
    Dim matches As Object
    Set matches = regEx.Execute(line)
    
    If matches.Count > 0 Then
        ExtraerBibKey = Trim(matches(0).SubMatches(0))
    Else
        ExtraerBibKey = ""
    End If
End Function

' --------------------------------------------------------------------
' Genera la diapositiva final con las referencias enumeradas.
' Llama a FormatearBibEntry para parsear los campos BibTeX.
' --------------------------------------------------------------------
Private Sub GenerarDiapositivaBibliografia(ByVal refsMap As Object, ByVal referencesData As Object)
    Dim sld As slide
    Set sld = ActivePresentation.slides.Add(ActivePresentation.slides.Count + 1, ppLayoutText)
    sld.Shapes.Title.TextFrame.textRange.text = "Bibliografía"
    
    Dim cuerpo As shape
    Set cuerpo = sld.Shapes.Placeholders(2)
    
    Dim keysArr() As Variant
    keysArr = refsMap.Keys
    
    ' Creamos un array para (key, numero)
    Dim refNumbers() As Variant
    ReDim refNumbers(0 To UBound(keysArr))
    
    Dim i As Long, j As Long
    For i = 0 To UBound(keysArr)
        Dim theKey As String
        theKey = keysArr(i)
        Dim numeroStr As String
        numeroStr = refsMap(theKey)   ' "[1]"
        Dim numSolo As Long
        numSolo = CLng(Replace(Replace(numeroStr, "[", ""), "]", ""))
        refNumbers(i) = Array(theKey, numSolo)
    Next i
    
    ' Ordenar por numSolo
    Dim temp As Variant
    For i = LBound(refNumbers) To UBound(refNumbers) - 1
        For j = i + 1 To UBound(refNumbers)
            If refNumbers(i)(1) > refNumbers(j)(1) Then
                temp = refNumbers(i)
                refNumbers(i) = refNumbers(j)
                refNumbers(j) = temp
            End If
        Next j
    Next i
    
    Dim bibText As String: bibText = ""
    
    For i = LBound(refNumbers) To UBound(refNumbers)
        theKey = refNumbers(i)(0)
        Dim theNum As Long
        theNum = refNumbers(i)(1)
        
        Dim labelNum As String
        labelNum = "[" & theNum & "] "
        
        Dim formattedRef As String
        If referencesData.Exists(theKey) Then
            formattedRef = FormatearBibEntry(referencesData(theKey), theKey)
        Else
            formattedRef = "(No encontrado en references.bib: " & theKey & ")"
        End If
        
        bibText = bibText & labelNum & formattedRef & vbCrLf & vbCrLf
    Next i
    
    cuerpo.TextFrame.textRange.text = bibText
End Sub

' --------------------------------------------------------------------
' Formatear cada entrada BibTeX con parseo sencillo: author, title, year, journal
' El resultado se muestra en la diapositiva final.
' --------------------------------------------------------------------
Private Function FormatearBibEntry(ByVal rawBib As String, ByVal refKey As String) As String
    Dim lines() As String
    lines = Split(rawBib, vbCrLf)
    
    Dim strAuthor As String: strAuthor = ""
    Dim strTitle As String:  strTitle = ""
    Dim strYear As String:   strYear = ""
    Dim strJournal As String: strJournal = ""
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim(lines(i))
        
        If InStr(1, ln, "author = {", vbTextCompare) > 0 Then
            strAuthor = ExtraerValorBib(ln)
        ElseIf InStr(1, ln, "title = {", vbTextCompare) > 0 Then
            strTitle = ExtraerValorBib(ln)
        ElseIf InStr(1, ln, "year = {", vbTextCompare) > 0 Then
            strYear = ExtraerValorBib(ln)
        ElseIf InStr(1, ln, "journal = {", vbTextCompare) > 0 Or InStr(1, ln, "booktitle = {", vbTextCompare) > 0 Then
            strJournal = ExtraerValorBib(ln)
        End If
    Next i
    
    ' Reemplazar tildes (sin diéresis)
    strAuthor = ReemplazarTildesLatex(strAuthor)
    strTitle = ReemplazarTildesLatex(strTitle)
    strYear = ReemplazarTildesLatex(strYear)
    strJournal = ReemplazarTildesLatex(strJournal)
    
    ' Si faltan campos esenciales, advertencia
    If strAuthor = "" Then MsgBox "Advertencia: a la referencia '" & refKey & "' le falta 'author'", vbExclamation
    If strTitle = "" Then MsgBox "Advertencia: a la referencia '" & refKey & "' le falta 'title'", vbExclamation
    If strYear = "" Then MsgBox "Advertencia: a la referencia '" & refKey & "' le falta 'year'", vbExclamation
    
    ' Armar estilo IEEE básico
    Dim finalRef As String
    finalRef = strAuthor
    If finalRef <> "" Then finalRef = finalRef & ", "
    
    If strTitle <> "" Then
        finalRef = finalRef & Chr(34) & strTitle & Chr(34) & ", "
    End If
    
    If strJournal <> "" Then
        finalRef = finalRef & "*" & strJournal & "*"
    End If
    
    If strYear <> "" Then
        If strJournal <> "" Then
            finalRef = finalRef & ", " & strYear & "."
        Else
            finalRef = finalRef & " " & strYear & "."
        End If
    End If
    
    FormatearBibEntry = finalRef
End Function

' --------------------------------------------------------------------
' Extrae lo que hay en { ... } (sin anidamiento complejo).
' --------------------------------------------------------------------
Private Function ExtraerValorBib(ByVal line As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.pattern = "\{([^}]*)\}"
    regEx.Global = False
    
    Dim matches As Object
    Set matches = regEx.Execute(line)
    
    If matches.Count > 0 Then
        ExtraerValorBib = matches(0).SubMatches(0)
    Else
        ExtraerValorBib = ""
    End If
End Function

' --------------------------------------------------------------------
' Reemplaza tildes LaTeX: {\'a}, {\'e}, {\'i}, {\'o}, {\'u}, {\~n}, etc.
' Maneja dobles llaves: "{{\'a}}" -> "á".
' --------------------------------------------------------------------
Private Function ReemplazarTildesLatex(ByVal textIn As String) As String
    Dim txt As String
    txt = textIn
    
    ' Manejo de dobles llaves. Por si hay casos como {{\'a}}
    txt = Replace(txt, "{{", "{")
    txt = Replace(txt, "}}", "}")
    
    ' Vocales con tilde minúsculas
    txt = Replace(txt, "{\'a}", "á")
    txt = Replace(txt, "{\'e}", "é")
    txt = Replace(txt, "{\'i}", "í")
    txt = Replace(txt, "{\'o}", "ó")
    txt = Replace(txt, "{\'u}", "ú")

    ' Mayúsculas
    txt = Replace(txt, "{\'A}", "Á")
    txt = Replace(txt, "{\'E}", "É")
    txt = Replace(txt, "{\'I}", "Í")
    txt = Replace(txt, "{\'O}", "Ó")
    txt = Replace(txt, "{\'U}", "Ú")

    ' Ñ y ñ
    txt = Replace(txt, "{\~n}", "ñ")
    txt = Replace(txt, "{\~N}", "Ñ")

    ' Eliminación de nuevas dobles llaves remanentes
    txt = Replace(txt, "{{", "{")
    txt = Replace(txt, "}}", "}")
    
    ReemplazarTildesLatex = txt
End Function

