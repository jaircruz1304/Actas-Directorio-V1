# Arquitectura v1.0 — Sin backend de IA ni vigilancia permanente

GitHub Pages (HTML/CSS/JS estáticos) -> MSAL M365 (login) -> Microsoft Graph -> matriz Excel / carpeta actas / índice JSON protegido en SharePoint.

Al sincronizar se recupera la versión actual del índice y matriz; el navegador comprueba el repositorio y ejecuta extracción de PDF/OCR para actas nuevas, incrementa el índice y lo publica si está autorizado. El índice no se aloja en GitHub.

**Límite de seguridad:** `config.js` ES público. Un GitHub Actions Secret que se incruste en un archivo JavaScript público no conservará el secreto. Este ZIP incluye un *localizador* SharePoint en config.js, no una contraseña. El usuario administrador debe restringir el enlace a usuarios con acceso autorizado y renovar el vínculo si actualmente está configurado para cualquiera con enlace.

**Límite operativo:** no hay ejecución en segundo plano; la detección solo ocurre cuando usuarios ingresan o sincronizan. Para análisis generativo real sería necesario un modelo local o API privada. Las respuestas extractivas no deben presentarse como juicios de IA.
