using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;

namespace InstaladorOffice;

/// <summary>Descarga Office Deployment Tool, verifica la firma de Microsoft y extrae setup.exe.</summary>
static class OdtDownloader
{
    public const string PaginaUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=49117";
    const string PaginaRespaldoUrl = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117";
    // La página de Microsoft responde 403 a clientes que no parecen navegador.
    const string UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36";
    static readonly Regex EnlaceRegex = new(
        @"https://download\.microsoft\.com/download/[\w\-/\.]*officedeploymenttool[\w\-\.]*\.exe",
        RegexOptions.IgnoreCase);

    /// <summary>Bloqueante: ejecutar en segundo plano. progreso(porcentaje o -1, mensaje).</summary>
    public static void Descargar(string destinoSetup, Action<int, string> progreso)
    {
        ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

        progreso(-1, "Buscando la versión más reciente de Office Deployment Tool…");
        string url = BuscarEnlace();

        string temp = Path.Combine(Path.GetTempPath(), "odt-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        try
        {
            string odtExe = Path.Combine(temp, Path.GetFileName(new Uri(url).AbsolutePath));
            DescargarArchivo(url, odtExe, progreso);

            progreso(-1, "Verificando la firma digital de Microsoft…");
            VerificarFirmaMicrosoft(odtExe);

            progreso(-1, "Extrayendo setup.exe…");
            string carpeta = Path.Combine(temp, "extract");
            Directory.CreateDirectory(carpeta);
            var psi = new ProcessStartInfo(odtExe, $"/quiet /extract:\"{carpeta}\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            using (var p = Process.Start(psi))
            {
                p.WaitForExit();
                if (p.ExitCode != 0)
                    throw new Exception($"La extracción del ODT terminó con código {p.ExitCode}.");
            }

            string setup = Path.Combine(carpeta, "setup.exe");
            if (!File.Exists(setup))
                throw new Exception("La extracción no generó setup.exe.");
            VerificarFirmaMicrosoft(setup);

            File.Copy(setup, destinoSetup, true);
            progreso(100, "Office Deployment Tool listo.");
        }
        finally
        {
            try { Directory.Delete(temp, true); } catch { }
        }
    }

    static string BuscarEnlace()
    {
        Exception ultimo = null;
        foreach (string pagina in new[] { PaginaUrl, PaginaRespaldoUrl })
        {
            try
            {
                using var wc = new WebClient { Encoding = Encoding.UTF8 };
                wc.Headers[HttpRequestHeader.UserAgent] = UserAgent;
                wc.Headers[HttpRequestHeader.Accept] = "text/html,application/xhtml+xml";
                wc.Headers[HttpRequestHeader.AcceptLanguage] = "en-US,en;q=0.9";
                var m = EnlaceRegex.Match(wc.DownloadString(pagina));
                if (m.Success) return m.Value;
            }
            catch (Exception ex) { ultimo = ex; }
        }
        throw new Exception("No se encontró el enlace de descarga del ODT en la página de Microsoft." +
            (ultimo != null ? " " + ultimo.Message : ""));
    }

    static void DescargarArchivo(string url, string destino, Action<int, string> progreso)
    {
        var req = (HttpWebRequest)WebRequest.Create(url);
        req.UserAgent = UserAgent;
        using var resp = (HttpWebResponse)req.GetResponse();
        using var entrada = resp.GetResponseStream();
        using var salida = File.Create(destino);
        long total = resp.ContentLength, hecho = 0;
        int ultimoPct = -2, n;
        var buf = new byte[81920];
        while ((n = entrada.Read(buf, 0, buf.Length)) > 0)
        {
            salida.Write(buf, 0, n);
            hecho += n;
            int pct = total > 0 ? (int)(hecho * 100 / total) : -1;
            if (pct == ultimoPct) continue;
            ultimoPct = pct;
            progreso(pct, "Descargando Office Deployment Tool… " + (pct >= 0 ? pct + " %" : hecho / 1024 + " KB"));
        }
    }

    // ---- Verificación Authenticode (WinVerifyTrust) ----
    static readonly Guid WintrustActionGenericVerifyV2 = new("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");

    [StructLayout(LayoutKind.Sequential)]
    struct WinTrustFileInfo
    {
        public uint cbStruct;
        public IntPtr pcwszFilePath;
        public IntPtr hFile;
        public IntPtr pgKnownSubject;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct WinTrustData
    {
        public uint cbStruct;
        public IntPtr pPolicyCallbackData;
        public IntPtr pSIPClientData;
        public uint dwUIChoice;
        public uint fdwRevocationChecks;
        public uint dwUnionChoice;
        public IntPtr pFile;
        public uint dwStateAction;
        public IntPtr hWVTStateData;
        public IntPtr pwszURLReference;
        public uint dwProvFlags;
        public uint dwUIContext;
        public IntPtr pSignatureSettings;
    }

    [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
    static extern int WinVerifyTrust(IntPtr hwnd, [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID, IntPtr pWVTData);

    public static void VerificarFirmaMicrosoft(string ruta)
    {
        IntPtr pRuta = Marshal.StringToCoTaskMemUni(ruta);
        IntPtr pArchivo = IntPtr.Zero, pDatos = IntPtr.Zero;
        int resultado;
        try
        {
            var fi = new WinTrustFileInfo
            {
                cbStruct = (uint)Marshal.SizeOf(typeof(WinTrustFileInfo)),
                pcwszFilePath = pRuta,
            };
            pArchivo = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(WinTrustFileInfo)));
            Marshal.StructureToPtr(fi, pArchivo, false);

            var wd = new WinTrustData
            {
                cbStruct = (uint)Marshal.SizeOf(typeof(WinTrustData)),
                dwUIChoice = 2,          // WTD_UI_NONE
                fdwRevocationChecks = 0, // WTD_REVOKE_NONE
                dwUnionChoice = 1,       // WTD_CHOICE_FILE
                pFile = pArchivo,
            };
            pDatos = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(WinTrustData)));
            Marshal.StructureToPtr(wd, pDatos, false);

            resultado = WinVerifyTrust(new IntPtr(-1), WintrustActionGenericVerifyV2, pDatos);
        }
        finally
        {
            if (pDatos != IntPtr.Zero) Marshal.FreeCoTaskMem(pDatos);
            if (pArchivo != IntPtr.Zero) Marshal.FreeCoTaskMem(pArchivo);
            Marshal.FreeCoTaskMem(pRuta);
        }

        string nombre = Path.GetFileName(ruta);
        if (resultado != 0)
            throw new Exception($"La firma digital de {nombre} no es válida (0x{resultado:X8}). Se descartó el archivo.");

        var cert = new X509Certificate2(X509Certificate.CreateFromSignedFile(ruta));
        if (cert.Subject.IndexOf("O=Microsoft Corporation", StringComparison.OrdinalIgnoreCase) < 0)
            throw new Exception($"{nombre} no está firmado por Microsoft Corporation.");
    }
}
