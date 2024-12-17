#PPTX_AutoRefs
Automatización de captions y gestión de referencias en PowerPoint con formato IEEE

PPTX_AutoRefs es una herramienta VBA que facilita la numeración automática de captions y la gestión de citas y bibliografía en presentaciones de PowerPoint. Diseñada para trabajos académicos y técnicos, automatiza el proceso de reemplazo de referencias LaTeX y genera una diapositiva final con la bibliografía ordenada en formato IEEE.

Características
Numeración automática de captions

Busca y numera prefijos configurables como Figura, Tabla, Gráfico, etc.
Aplica formato opcional (negrita o itálica) a los captions.
Gestión de referencias en formato IEEE

Sustituye citas LaTeX \cite{key} por referencias numeradas [n] en orden de aparición.
Genera una diapositiva final llamada "Bibliografía" con las referencias ordenadas automáticamente.
Reemplazo de caracteres LaTeX

Corrige caracteres especiales como \'a, \~n, etc., para mostrar una bibliografía limpia.
Requisitos
PowerPoint con soporte para macros VBA (.pptm).
Archivo references.bib en la misma carpeta que la presentación.
Instalación
Abre tu archivo PowerPoint con extensión .pptm.
Accede al editor VBA con Alt + F11.
Inserta un nuevo módulo e importa el código VBA.
Personaliza el array prefixArray si necesitas agregar más prefijos (Figura, Tabla, etc.).
Guarda el archivo y ejecuta las macros:
auto_number_captions: Numeración automática de captions.
GenerarReferencias: Generación de referencias bibliográficas en formato IEEE.
Uso
Ejemplo inicial:
plaintext
Copiar código
Fig: Esquema del proceso  
Tabla 3: Resultados experimentales  
\cite{smith2023}
Resultado después de ejecutar las macros:
plaintext
Copiar código
Fig 1: Esquema del proceso  
Tabla 1: Resultados experimentales  
[1]
Diapositiva "Bibliografía":

plaintext
Copiar código
[1] J. Smith, "Título del artículo", Revista, vol. 10, no. 2, pp. 123–130, 2023.
Personalización
Modifica el array prefixArray en el código VBA para incluir los prefijos que necesites.
Ajusta el formato (negrita o itálica) en las secciones de numeración según tus preferencias.
Notas importantes
Actualmente, la bibliografía se genera únicamente en formato IEEE.
Los archivos temporales generados por PowerPoint no deben subirse al repositorio. Usa un archivo .gitignore adecuado.
Contribuciones
Si deseas mejorar esta herramienta, abre un issue o envía un pull request. Toda mejora es bienvenida.

Créditos
Desarrollado por: Benjamín Moya Giachetti

Estado del proyecto
Versión inicial funcional. Mejoras en proceso.
