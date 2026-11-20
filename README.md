# Office Deployment Tool — configuraciones 2019–2024

Archivos para instalaciones Click-to-Run desde la CDN de Microsoft, y una herramienta para activar ediciones por volumen con los servidores KMS o claves MAK del cliente.

## Uso

1. Descarga Office Deployment Tool desde https://www.microsoft.com/en-us/download/details.aspx?id=49117 y extrae setup.exe.
2. Coloca setup.exe junto a las carpetas Retail y Volume.
3. En CMD, desde la carpeta principal, ejecuta por ejemplo:
   setup.exe /configure Retail\Professional2024Retail.xml
   setup.exe /configure Volume\ProPlus2024Volume.xml
4. En PowerShell usa: .\setup.exe /configure .\Retail\Professional2024Retail.xml

## Instalador y activador gráfico (InstaladorOffice.exe)

Un solo exe para .NET Framework 4.8 (incluido en Windows 10/11). No necesita los demás archivos del repositorio.

### Compilar

Requiere el SDK de .NET 8 o superior; funciona en Linux, macOS o Windows.

- Linux/macOS: `./build.sh`
- Windows: `build.cmd`

El exe queda en `dist/InstaladorOffice.exe`. El código está en `src/InstaladorOffice/`.

### Configurar servidores KMS y claves

Edita `src/InstaladorOffice/Configuracion.cs` y vuelve a compilar:

- `ServidoresKms`: servidores KMS de tus clientes; aparecen en el menú desplegable «Servidor KMS».
- `ClavesOffice` y `ClavesWindows`: ya incluyen las GVLK públicas de Microsoft; agrega tus claves MAK al final de cada lista.

Todo lo que pongas ahí queda dentro del exe y se puede extraer de él. No subas ese archivo con datos reales a un repositorio público.

### Pestañas

- **Instalar Office:** elige la edición, los programas, la arquitectura y el idioma. Si falta setup.exe, lo descarga de Microsoft y verifica su firma digital. «Instalar» ejecuta `setup.exe /configure`; «Guardar XML…» solo guarda el XML en `Personalizado`.
- **Activar Office:** detecta la edición Click-to-Run y las licencias instaladas. Usa `ospp.vbs` para instalar o desinstalar claves (`/inpkey`, `/unpkey`), configurar el servidor KMS (`/sethst`, `/setprt`, `/remhst`) y activar (`/act`). Solo permite activar ediciones por volumen.
- **Activar Windows:** detecta la edición y la licencia. Usa `slmgr.vbs` (`/ipk`, `/upk`, `/cpky`, `/skms`, `/ckms`, `/ato`, `/dlv`). Windows Home no admite activación por volumen.

El estado de las licencias se lee por WMI (`SoftwareLicensingProduct`); la salida de cada comando aparece en la consola de la pestaña.

## Instalador interactivo en consola (instalar.cmd)

1. Si falta setup.exe, el script lo descarga y verifica automáticamente con PowerShell.
2. Haz doble clic en instalar.cmd; se pedirá elevar permisos de administrador.
3. Elige la versión (Retail o Volume), marca o desmarca los programas escribiendo sus números (por ejemplo `4 5`) y revisa el resumen.
4. En el resumen puedes cambiar la arquitectura (64/32 bits) y el idioma (es-mx, es-es, en-us, MatchOS); después instala o guarda solo el XML.
5. Los XML generados se guardan en la carpeta Personalizado y se pueden reutilizar con `setup.exe /configure Personalizado\<archivo>.xml`.

## Consideraciones

- Retail: se omite Channel, para no forzar canales de licencias por volumen. Volume: PerpetualVL2019, PerpetualVL2021 o PerpetualVL2024.
- Los ID se basan en la lista oficial: https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run
- La instalación no garantiza activación: usa una licencia legítima que corresponda exactamente al producto.
- Home & Student / Home 2024 y Personal no incluyen Outlook; la exclusión explícita evita solicitarlo.
- La disponibilidad de apps depende de la edición. ExcludeApp no agrega apps ausentes.
- Office 2019 y 2021 están fuera de soporte desde el 14 de octubre de 2025. Evita nuevas instalaciones para equipos expuestos si puedes elegir una versión compatible.
- No se usa RemoveMSI para evitar desinstalar versiones existentes sin consentimiento.
- No se incluyen Standard Retail: los ID Standard2019/2021/2024Volume corresponden a licencias por volumen.
- Algunas combinaciones antiguas de Office y Windows pueden ser incompatibles. Actualiza ODT antes de instalar.
