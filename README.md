# FIAS · Inteligencia de Directorio v0.5

## Alcance de esta entrega

Interfaz Enterprise nueva, independiente de los proyectos anteriores. Aplicación estática lista para **publicación en GitHub Pages** (no requiere npm ni compilar JSX/TSX). No se incluye ningún XLSX, PDF o datos institucionales dentro de los archivos públicos.

- Autenticación delegada con Microsoft Entra ID/MSAL y acceso a matriz XLSX mediante Graph y enlaces de SharePoint del FIAS.
- Búsqueda documental de matriz, exploración filtrada, estadísticas, detalle y enlace a la fuente.
- Búsqueda neuronal real opcional: embeddings multilingües calculados **en el navegador**, con `@huggingface/transformers`. Requiere descargar un modelo del proveedor cuando se activa por primera vez, compatible con recursos y políticas de la red del usuario.
- Lectura adicional de PDF a través de Microsoft Graph, recorriendo subcarpetas del enlace de repositorio. Extrae texto mediante PDF.js, agrupa pasajes en memoria y permite recuperar pasajes con nombre y páginas. Documentos escaneados que no contienen texto seleccionable requieren OCR que **no está incluido**.
- Síntesis **extractiva** con referencias sin generar afirmaciones adicionales. Las respuestas **generativas** requieren una API IA institucional autenticada, no incluida en el despliegue público.

## Instalación en GitHub Pages

1. Extraer los archivos y copiar `index.html` y `assets/` **a la raíz del repositorio** `Actas-Directorio`.
2. Abrir `https://jaircruz1304.github.io/Actas-Directorio/index.html`, limpiar caché del navegador (Ctrl+F5).
3. En Microsoft Entra ID registrar como **SPA** el URI de redirección anterior para el identificador cliente configurado. Verificar consentimiento y permisos delegados de Microsoft Graph `User.Read`, `Files.Read.All`, `Sites.Read.All`. Emplear permisos mínimos que realmente se necesiten y aprobaciones del administrador.
4. Iniciar sesión y sincronizar la matriz.
5. Para indexar PDF, entrar a **Asistente inteligente → Indexar PDF de SharePoint**. La descarga de los PDF exige acceso a esos archivos y a las carpetas; algunas políticas CORS y de descargas temporales de Graph pueden impedir operar desde GitHub Pages.
6. Activar la búsqueda semántica, que puede consumir memoria y descargar varios recursos externos. Los embeddings se calculan localmente y los documentos en RAM se pierden al recargar.
7. Para síntesis generativa, configurar una API protegida por Entra ID con `scope` válido y procesamiento respaldado por un servicio Azure autorizado. **No colocar claves de IA en HTML/JS o GitHub Pages**.

## Límites y seguridad

- No se verificó esta entrega contra el tenant del FIAS ni se ejecutó una ingesta de las actas protegidas. Las capacidades Graph requieren prueba en la cuenta corporativa.
- La carpeta compartida se usa como **origen remoto**, no como archivo PDF adjunto. No se afirma haber descargado todos los documentos ni aplicado OCR.
- En ausencia de una API corporativa configurada, se usa extracción y ranking; esto **no equivale** a respuestas generativas.
- La IA generativa **no debe asumir** que una coincidencia temática es una resolución aprobada, ni inventar números, fechas o montos. Su backend debe validar los identificadores de evidencia frente al catálogo autenticado antes de procesar documentos, registrar auditoría, filtrar por permisos y devolver citas verificables.
- La salida del servicio IA esperada es JSON `{ "answer": "...", "citations": ["id-registro", "pdf-id"] }`. Solo se muestran citas a identificadores presentes en las evidencias seleccionadas.
- El archivo matriz incluye fechas, asuntos y resoluciones; los PDF pueden tener restricciones de acceso y protección de datos. Mantener todo procesamiento y retención de información conforme a las políticas del FIAS.
- El lector PDF.js no hace OCR ni interpreta firmas, anexos o tablas complejas con garantía de exactitud.

## Próximos pasos de producción

Arquitectura recomendada: Microsoft Entra ID → API Azure Functions autenticada → Graph (acceso delegado/OBO) → extracción/OCR Azure AI Document Intelligence → índice de fragmentos y metadatos en Azure AI Search → respuestas con Azure OpenAI/Azure AI Foundry y citas verificadas. Agregar auditoría, ACL por documento, sincronización incremental mediante delta queries, monitoreo, retención y pruebas de calidad antes de habilitar IA generativa.

## Pruebas

`node --check assets/app.js` se ejecutó sin errores sintácticos. En este entorno el navegador de pruebas bloqueó la navegación a localhost y file://, por lo que **no se han completado las pruebas visuales ni la validación del tenant**.
