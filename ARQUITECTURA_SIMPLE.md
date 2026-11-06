# FIAS Directorio · Experiencia Enterprise v1.1

## Flujo automático

1. Microsoft 365 identifica al usuario institucional con MSAL.
2. Microsoft Graph recupera el índice JSON privado alojado en SharePoint.
3. La matriz Excel se descarga de SharePoint y su información se normaliza.
4. Se recorre el repositorio documental para comprobar actas PDF nuevas.
5. Si hay novedades, el navegador procesa texto/OCR, actualiza el índice local y
   *intenta* publicarlo con control ETag en SharePoint (si hay autorización de edición).
6. El usuario puede consultar el explorador, la inteligencia documental y
   la trazabilidad de antecedentes, siempre con referencias cuando están disponibles.

No se utiliza Azure OpenAI, Azure AI Search, procesos Windows de vigilancia ni
copias de los documentos privados en GitHub Pages.

## UX

- Navegación principal: Panorama, Explorador, Inteligencia, Trazabilidad.
- Los controles técnicos manuales se retiraron; los mensajes de estado se concentran
  en un diálogo de sincronización con etapas, indicadores y errores visibles.
- La opción discreta «Actualizar datos» se mantiene como recuperación.
- «Trazabilidad» presenta registros relacionados por tema en orden cronológico;
  no prueba que entre ellos exista una dependencia jurídica formal.

## Seguridad y límites

- `config.js` es público; su enlace debe ser un identificador, no un token de acceso.
- Los permisos de SharePoint aplican a cada usuario; revise el uso de enlaces públicos.
- Si Files.ReadWrite no está consentido, la publicación automática puede quedar pendiente.
- La semántica neuronal del navegador es opcional según compatibilidad/red, con
  búsqueda contextual como alternativa; no hay generación de lenguaje avanzada.
