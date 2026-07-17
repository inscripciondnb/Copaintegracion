COPA INTEGRACIÓN DNB 2026 — INSTALACIÓN

ARCHIVOS PARA GITHUB
- index.html

ARCHIVO PARA GOOGLE APPS SCRIPT
- Code-final.gs

PASOS
1. Crea un Google Sheets vacío.
2. Ve a Extensiones > Apps Script.
3. Borra el código existente y pega Code-final.gs.
4. Completa MAIL_ORGANIZACION con tu correo.
5. Guarda el proyecto.
6. Implementar > Nueva implementación > Aplicación web.
7. Ejecutar como: Yo.
8. Quién tiene acceso: Cualquier persona.
9. Autoriza Drive, Documentos, Hojas de cálculo y Gmail.
10. Copia la URL que termina en /exec.
11. Abre index.html y reemplaza:
    const SCRIPT_URL = "";
    por:
    const SCRIPT_URL = "TU_URL_EXEC";
12. Sube index.html a la raíz del repositorio GitHub.
13. En GitHub: Settings > Pages > Deploy from a branch > main / root.

IMPORTANTE
- No subas Code-final.gs como secreto ni credenciales: el archivo no contiene claves.
- La URL /exec del Apps Script puede estar en el HTML público; las autorizaciones permanecen en Google.
- No cambies la implementación a "solo yo", porque la web pública no podrá enviar formularios.
