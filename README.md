# Office Deployment Tool — configuraciones 2019–2024

Archivos para instalaciones Click-to-Run desde la CDN de Microsoft. No incluyen licencias ni activadores.

## Uso
1. Descarga Office Deployment Tool desde https://www.microsoft.com/en-us/download/details.aspx?id=49117 y extrae setup.exe.
2. Coloca setup.exe junto a las carpetas Retail y Volume.
3. En CMD, desde la carpeta principal, ejecuta por ejemplo:
   setup.exe /configure Retail\Professional2024Retail.xml
   setup.exe /configure Volume\ProPlus2024Volume.xml
4. En PowerShell usa: .\setup.exe /configure .\Retail\Professional2024Retail.xml

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
