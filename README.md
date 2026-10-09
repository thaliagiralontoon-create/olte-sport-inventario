# OLTE SPORT — Inventario

Sitio de inventario de chalecos reductores de hombre y mujer.

## Base de datos Google Sheets

El sitio incluye conexión a Google Sheets mediante Google Apps Script. El usuario debe:

1. Abrir **Extensiones → Apps Script** desde la hoja de OLTE SPORT.
2. Reemplazar `Code.gs` con `apps-script/Code.gs` y guardar.
3. Ejecutar `setupOlteSportDatabase` una vez y autorizar el acceso.
4. Implementar como **Aplicación web**, ejecutar como **Yo** y dar acceso a **Solo yo**.
5. Copiar la URL terminada en `/exec` y pegarla en el sitio, en **Respaldos → Base de datos Google Sheets**.
6. Elegir si carga desde Sheets o sube los datos de este navegador. La segunda opción reemplaza las filas de datos de las pestañas; descarga un respaldo JSON primero.

La app mantiene una copia local en cada navegador y sincroniza con la hoja después de conectarla. Para empezar desde otro equipo, usa **Cargar datos desde Sheets**. Mantén la implementación de Apps Script restringida a **Solo yo**.

## GitHub Pages

En **Settings → Pages**, elegir **Deploy from a branch**, rama `main`, carpeta `/ (root)`.
