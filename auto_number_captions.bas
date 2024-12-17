Attribute VB_Name = "Module1"
Sub auto_number_captions()

' ---------------------------------------------------------------
' Creado por Benjamín Moya Giachetti
' Versión: 1.0
'
' Esta macro automatiza la numeración de captions en una presentación
' de PowerPoint. Busca prefijos configurados (como Figura, Fig, Tabla,
' Cuadro, Gráfico, etc.) y los numera de forma independiente y secuencial.
' Si encuentra un prefijo sin número (e.g., "Fig:"), le asigna la numeración
' correspondiente. También corrige números incorrectos en los captions.
'
' Configuración:
' - Modifica el array "prefixArray" para agregar los prefijos que desees.
' - Ajusta "boldCaption" y "italicCaption" para aplicar formato a los prefijos.
'
' Modo de uso:
' 1. Accede al editor de VBA en PowerPoint (Alt + F11).
' 2. Inserta un nuevo módulo e importa este código.
' 3. Personaliza los prefijos en el array "prefixArray".
' 4. Ejecuta la macro "auto_number_captions".
'
' Ejemplo de uso en la diapositiva:
' Antes de ejecutar la macro:
'     Fig: Imagen sin numerar
'     Cuadro 7: Texto existente
'
' Después de ejecutar la macro:
'     Fig 1: Imagen sin numerar
'     Cuadro 1: Texto existente
'
' ---------------------------------------------------------------


    ' --- CONFIGURACIÓN INICIAL ---
    Dim prefixArray As Variant
    prefixArray = Array("Figura", "Fig", "Tabla", "Cuadro", "Gráfico") ' Prefijos a buscar y numerar
    Dim boldCaption As Boolean: boldCaption = False       ' True para negrita, False para normal
    Dim italicCaption As Boolean: italicCaption = False  ' True para cursiva, False para normal
    
    ' --- VARIABLES INTERNAS ---
    Dim slide As slide
    Dim shape As shape
    Dim prefixCounters() As Integer
    Dim i As Integer, prefix As String
    Dim totalPrefixes As Integer
    
    totalPrefixes = UBound(prefixArray)
    ReDim prefixCounters(0 To totalPrefixes) ' Inicializa los contadores para cada prefijo
    
    ' Recorre todas las diapositivas
    For Each slide In ActivePresentation.slides
        ' Recorre todas las formas en la diapositiva
        For Each shape In slide.Shapes
            If shape.Type = msoTextBox Then ' Solo busca captions en cuadros de texto
                For i = LBound(prefixArray) To UBound(prefixArray)
                    prefix = prefixArray(i)
                    ' Busca si hay un prefijo con número o solo el prefijo seguido de ":"
                    If IsCaptionWithPrefix(shape.TextFrame.textRange.text, prefix) Then
                        prefixCounters(i) = prefixCounters(i) + 1
                        UpdateCaption shape, prefix, prefixCounters(i), boldCaption, italicCaption
                        Exit For ' Salir después de encontrar y procesar el prefijo correcto
                    End If
                Next i
            End If
        Next shape
    Next slide
    
    MsgBox "Numeración completada.", vbInformation
End Sub

Function IsCaptionWithPrefix(text As String, prefix As String) As Boolean
    ' Verifica si el texto contiene el prefijo seguido de un número o solo ":"
    If InStr(1, text, prefix & " ", vbTextCompare) > 0 Or _
       InStr(1, text, prefix & ":", vbTextCompare) > 0 Then
        IsCaptionWithPrefix = True
    Else
        IsCaptionWithPrefix = False
    End If
End Function

Sub UpdateCaption(targetShape As shape, prefix As String, number As Integer, bold As Boolean, italic As Boolean)
    ' Actualiza el texto del caption con el número correcto
    Dim fullCaption As String
    Dim restOfText As String
    
    ' Si el texto tiene ":" pero sin número, extrae el texto después de ":"
    If InStr(targetShape.TextFrame.textRange.text, ":") > 0 Then
        restOfText = Mid(targetShape.TextFrame.textRange.text, InStr(targetShape.TextFrame.textRange.text, ":") + 1)
    Else
        restOfText = ""
    End If
    
    ' Reconstruye el caption
    fullCaption = prefix & " " & number & ": " & Trim(restOfText)
    
    ' Aplica el texto y el formato especial al prefijo y número
    With targetShape.TextFrame.textRange
        .text = fullCaption
        
        ' Formatea solo el prefijo y el número
        With .Characters(1, Len(prefix & " " & CStr(number) & ":")).Font
            .bold = bold
            .italic = italic
        End With
        
        ' El resto del texto se deja sin formato especial
        With .Characters(Len(prefix & " " & CStr(number) & ":") + 1).Font
            .bold = False
            .italic = False
        End With
    End With
End Sub


