# Copa Integración DNB 2026

Proyecto web de inscripción conectado con Google Apps Script.

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Sube el archivo `index.html` en la raíz del repositorio.
3. Entra en `Settings > Pages`.
4. Selecciona `Deploy from a branch`.
5. Selecciona la rama `main` y la carpeta `/ (root)`.
6. Guarda y abre la dirección que te entrega GitHub Pages.

## Google Apps Script

El archivo `apps-script/Code.gs` es una copia de respaldo. Debes pegarlo dentro del editor de Google Apps Script asociado a tu Google Sheets.

La web ya está conectada a esta implementación:

`https://script.google.com/macros/s/AKfycbzqvLtiyyijTlrb5hLqQpQe-HARp521otZeVXoVb0w0Ji7FQF7INb6Bt2KSiAqMISqopg/exec`

Antes de usar la web públicamente, confirma en Apps Script:

- Ejecutar como: **Yo**.
- Acceso: **Cualquier persona**.
- La implementación corresponde a la versión más reciente del código.
- Se autorizaron los permisos de Sheets, Drive y Gmail.

## Contenido

- `index.html`: página que debe publicar GitHub Pages.
- `apps-script/Code.gs`: backend para Google Apps Script.
- `documentos-ejemplo/`: modelos de la ficha Word y del Excel administrativo.
