using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using Microsoft.Win32;

namespace InstaladorOffice;

/// <summary>Una licencia instalada (clase WMI SoftwareLicensingProduct).</summary>
sealed class Licencia
{
    public string Nombre { get; set; }
    public string Descripcion { get; set; }
    public int Estado { get; set; }
    public string ClaveParcial { get; set; }
    public string Kms { get; set; }
    public string KmsDescubierto { get; set; }
    public long MinutosGracia { get; set; }

    /// <summary>Canal de la descripción, p. ej. «Office 24, VOLUME_KMSCLIENT channel» → VOLUME_KMSCLIENT.</summary>
    public string Canal
    {
        get
        {
            string d = Descripcion ?? "";
            int i = d.LastIndexOf(", ", StringComparison.Ordinal);
            string c = i >= 0 ? d.Substring(i + 2) : d;
            return c.Replace(" channel", "").Trim();
        }
    }

    public bool EsVolumen => Canal.StartsWith("VOLUME", StringComparison.OrdinalIgnoreCase);

    public string EstadoTexto => Estado switch
    {
        0 => "Sin licencia",
        1 => "Activado",
        2 => "Periodo de gracia inicial",
        3 => "Gracia (hardware cambiado)",
        4 => "Gracia (no genuino)",
        5 => "Notificación (sin activar)",
        6 => "Gracia extendida",
        _ => "Desconocido (" + Estado + ")",
    };

    public string GraciaTexto => MinutosGracia > 0 ? $"{MinutosGracia / 1440} días" : "";

    public string KmsTexto =>
        !string.IsNullOrEmpty(Kms) ? Kms :
        !string.IsNullOrEmpty(KmsDescubierto) ? KmsDescubierto + " (DNS)" : "";
}

/// <summary>Resultado de detectar el producto instalado.</summary>
sealed class Deteccion
{
    /// <summary>Texto de resumen (varias líneas) para mostrar.</summary>
    public string Resumen { get; set; } = "";
    /// <summary>Si es false se deshabilitan instalar clave / KMS / activar.</summary>
    public bool Permitido { get; set; }
    /// <summary>Advertencia en rojo/naranja; vacía si todo está bien.</summary>
    public string Advertencia { get; set; } = "";
    public ClaveProducto ClaveSugerida { get; set; }
    /// <summary>Líneas de diagnóstico para la consola (p. ej. rutas revisadas).</summary>
    public List<string> Detalles { get; } = new();
}

interface IActivador
{
    string Nombre { get; }
    ClaveProducto[] Claves { get; }
    /// <summary>true si cada licencia se desinstala por separado (Office); false si hay una sola clave (Windows).</summary>
    bool VariasClaves { get; }

    Deteccion Detectar();
    List<Licencia> Licencias();
    Comando Estado();
    Comando[] DesinstalarClave(Licencia licencia);
    Comando InstalarClave(string clave);
    Comando[] ConfigurarKms(string host, int puerto);
    Comando[] BorrarKms();
    Comando Activar();
}

static class Wmi
{
    public const string AppOffice = "0ff1ce15-a989-479d-af46-f275c6370663";
    public const string AppWindows = "55c92734-d682-4d71-983e-d6ec3f16059f";

    public static List<Licencia> Consultar(string applicationId)
    {
        const string campos = "Name, Description, LicenseStatus, PartialProductKey, GracePeriodRemaining, " +
                              "KeyManagementServiceMachine, KeyManagementServicePort, DiscoveredKeyManagementServiceMachineName";
        string consulta = $"SELECT {campos} FROM SoftwareLicensingProduct " +
                          $"WHERE ApplicationID='{applicationId}' AND PartialProductKey IS NOT NULL";
        var lista = new List<Licencia>();
        using var buscador = new ManagementObjectSearcher(@"root\CIMV2", consulta);
        foreach (ManagementObject o in buscador.Get())
        {
            using (o)
            {
                string kms = o["KeyManagementServiceMachine"] as string;
                long puerto = Convert.ToInt64(o["KeyManagementServicePort"] ?? 0L);
                if (!string.IsNullOrEmpty(kms) && puerto > 0 && puerto != 1688) kms += ":" + puerto;
                lista.Add(new Licencia
                {
                    Nombre = o["Name"] as string,
                    Descripcion = o["Description"] as string,
                    Estado = Convert.ToInt32(o["LicenseStatus"] ?? 0),
                    ClaveParcial = o["PartialProductKey"] as string,
                    Kms = kms,
                    KmsDescubierto = o["DiscoveredKeyManagementServiceMachineName"] as string,
                    MinutosGracia = Convert.ToInt64(o["GracePeriodRemaining"] ?? 0L),
                });
            }
        }
        return lista;
    }

    public static RegistryKey AbrirHklm(string ruta) =>
        RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(ruta);
}

// ---------------------------------------------------------------------------
// Office (ospp.vbs)
// ---------------------------------------------------------------------------
sealed class ActivadorOffice : IActivador
{
    public string Nombre => "Office";
    public ClaveProducto[] Claves => Configuracion.ClavesOffice;
    public bool VariasClaves => true;

    /// <summary>Ruta de OSPP.VBS o null si no se encontró.</summary>
    public string Ospp { get; private set; }

    /// <summary>Ruta elegida a mano por el usuario; tiene prioridad sobre la búsqueda automática.</summary>
    public string OsppManual { get; set; }

    /// <summary>
    /// Busca OSPP.VBS en este orden y anota en <paramref name="revisadas"/> cada carpeta revisada:
    ///  1. Carpetas que indica el registro (InstallPath de Click-to-Run e InstallRoot de Office 16.0).
    ///  2. Program Files de 64 y 32 bits: Microsoft Office\Office16 y Microsoft Office\root\Office16.
    ///  3. Búsqueda recursiva (máx. 4 niveles) dentro de las carpetas «Microsoft Office» anteriores.
    /// </summary>
    static string BuscarOspp(List<string> revisadas)
    {
        var raicesOffice = new List<string>();
        var carpetas = new List<string>();

        void AgregarRaiz(string ruta)
        {
            if (string.IsNullOrWhiteSpace(ruta)) return;
            ruta = Environment.ExpandEnvironmentVariables(ruta.Trim().Trim('"')).TrimEnd('\\');
            if (raicesOffice.Contains(ruta, StringComparer.OrdinalIgnoreCase)) return;
            raicesOffice.Add(ruta);
            carpetas.Add(Path.Combine(ruta, "Office16"));
            carpetas.Add(Path.Combine(ruta, @"root\Office16"));
        }

        // 1. Registro: C2R guarda la carpeta base (p. ej. C:\Program Files\Microsoft Office)
        foreach (var vista in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, vista);
            using (var k = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun"))
                AgregarRaiz(k?.GetValue("InstallPath") as string);
            // InstallRoot apunta a ...\root\Office16\ en C2R o a ...\Office16\ en MSI
            using (var k = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot"))
            {
                if (k?.GetValue("Path") is string p && !string.IsNullOrWhiteSpace(p))
                {
                    p = p.Trim().TrimEnd('\\');
                    if (!carpetas.Contains(p, StringComparer.OrdinalIgnoreCase)) carpetas.Add(p);
                    // ...\Microsoft Office\root\Office16 → ...\Microsoft Office
                    var padre = Directory.GetParent(p)?.Parent;
                    if (padre != null) AgregarRaiz(padre.Name.Equals("root", StringComparison.OrdinalIgnoreCase)
                        ? padre.Parent?.FullName : padre.FullName);
                }
            }
        }

        // 2. Program Files de 64 y 32 bits (ProgramW6432 da el de 64 bits aunque el proceso sea de 32)
        foreach (string pf in new[]
                 {
                     Environment.GetEnvironmentVariable("ProgramW6432"),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                 })
        {
            if (!string.IsNullOrEmpty(pf)) AgregarRaiz(Path.Combine(pf, "Microsoft Office"));
        }

        foreach (string carpeta in carpetas)
        {
            string ruta = Path.Combine(carpeta, "OSPP.VBS");
            revisadas.Add(ruta);
            if (File.Exists(ruta)) return ruta;
        }

        // 3. Búsqueda recursiva acotada
        foreach (string raiz in raicesOffice.Where(Directory.Exists))
        {
            revisadas.Add(raiz + @"\… (búsqueda recursiva)");
            string encontrado = BuscarRecursivo(raiz, "OSPP.VBS", 4);
            if (encontrado != null) return encontrado;
        }
        return null;
    }

    static string BuscarRecursivo(string carpeta, string archivo, int profundidad)
    {
        try
        {
            string ruta = Path.Combine(carpeta, archivo);
            if (File.Exists(ruta)) return ruta;
            if (profundidad == 0) return null;
            foreach (string sub in Directory.EnumerateDirectories(carpeta))
            {
                string r = BuscarRecursivo(sub, archivo, profundidad - 1);
                if (r != null) return r;
            }
        }
        catch (UnauthorizedAccessException) { }
        catch (IOException) { }
        return null;
    }

    public Deteccion Detectar()
    {
        var d = new Deteccion();
        var revisadas = new List<string>();
        Ospp = !string.IsNullOrEmpty(OsppManual) && File.Exists(OsppManual) ? OsppManual : BuscarOspp(revisadas);
        if (Ospp == null)
        {
            d.Detalles.Add("Rutas revisadas para OSPP.VBS:");
            d.Detalles.AddRange(revisadas.Select(r => "  - " + r));
        }

        string ids = "", plataforma = "", version = "";
        using (var k = Wmi.AbrirHklm(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
        {
            if (k != null)
            {
                ids = k.GetValue("ProductReleaseIds") as string ?? "";
                plataforma = k.GetValue("Platform") as string ?? "";
                version = k.GetValue("VersionToReport") as string ?? "";
            }
        }
        var productos = ids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                           .Select(s => s.Trim()).ToList();

        if (productos.Count > 0)
            d.Resumen = $"Office Click-to-Run: {string.Join(", ", productos)}\n" +
                        $"Plataforma: {plataforma}   ·   Versión: {version}";
        else
            d.Resumen = "No se detectó Office Click-to-Run.";
        d.Resumen += "\nospp.vbs: " + (Ospp ?? "no encontrado");

        if (Ospp == null)
        {
            d.Advertencia = "No se encontró OSPP.VBS. Si Office está instalado, usa «Elegir ospp.vbs…» para indicar su ubicación.";
            return d;
        }

        var volumen = productos.Where(p => p.EndsWith("Volume", StringComparison.OrdinalIgnoreCase)).ToList();
        if (productos.Count > 0 && volumen.Count == 0)
        {
            d.Advertencia = "Office instalado es Retail. La activación KMS/MAK solo aplica a ediciones por volumen; " +
                            "instala una edición Volume desde la pestaña «Instalar Office».";
            return d;
        }
        if (productos.Count == 0)
            d.Advertencia = "No se pudo identificar la edición (¿instalación MSI?). Verifica que sea por volumen.";

        d.Permitido = true;
        d.ClaveSugerida = Claves.FirstOrDefault(c =>
            volumen.Any(v => string.Equals(v, c.Coincide, StringComparison.OrdinalIgnoreCase)));
        return d;
    }

    public List<Licencia> Licencias() => Wmi.Consultar(Wmi.AppOffice);

    string RequiereOspp() => Ospp ?? throw new InvalidOperationException("No se encontró OSPP.VBS. Pulsa «Detectar».");

    public Comando Estado() =>
        Comando.Script("Estado de las licencias de Office", RequiereOspp(), "/dstatus");

    public Comando[] DesinstalarClave(Licencia l) => new[]
    {
        Comando.Script($"Desinstalar clave …{l.ClaveParcial} ({l.Nombre})", RequiereOspp(), "/unpkey:" + l.ClaveParcial),
    };

    public Comando InstalarClave(string clave) =>
        Comando.Script("Instalar clave de producto", RequiereOspp(), "/inpkey:" + clave);

    public Comando[] ConfigurarKms(string host, int puerto) => new[]
    {
        Comando.Script("Configurar servidor KMS", RequiereOspp(), "/sethst:" + host),
        Comando.Script("Configurar puerto KMS", RequiereOspp(), "/setprt:" + puerto),
    };

    public Comando[] BorrarKms() => new[]
    {
        Comando.Script("Quitar servidor KMS (usar descubrimiento DNS)", RequiereOspp(), "/remhst"),
    };

    public Comando Activar() => Comando.Script("Activar Office", RequiereOspp(), "/act");
}

// ---------------------------------------------------------------------------
// Windows (slmgr.vbs)
// ---------------------------------------------------------------------------
sealed class ActivadorWindows : IActivador
{
    public string Nombre => "Windows";
    public ClaveProducto[] Claves => Configuracion.ClavesWindows;
    public bool VariasClaves => false;

    static string Slmgr => Path.Combine(Environment.SystemDirectory, "slmgr.vbs");

    public Deteccion Detectar()
    {
        var d = new Deteccion();
        string nombre = "", edicion = "", version = "";
        int build = 0;
        using (var k = Wmi.AbrirHklm(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
        {
            if (k != null)
            {
                nombre = k.GetValue("ProductName") as string ?? "";
                edicion = k.GetValue("EditionID") as string ?? "";
                version = k.GetValue("DisplayVersion") as string ?? k.GetValue("ReleaseId") as string ?? "";
                int.TryParse(k.GetValue("CurrentBuild") as string, out build);
            }
        }
        // ProductName sigue diciendo «Windows 10» en Windows 11
        if (build >= 22000 && nombre.StartsWith("Windows 10", StringComparison.Ordinal))
            nombre = "Windows 11" + nombre.Substring("Windows 10".Length);

        d.Resumen = $"{nombre}   ·   Edición: {edicion}\nVersión: {version}   ·   Compilación: {build}";

        if (edicion.StartsWith("Core", StringComparison.OrdinalIgnoreCase))
        {
            d.Advertencia = "Windows Home no admite activación por volumen (KMS/MAK).";
            return d;
        }

        d.Permitido = true;
        var candidatas = Claves.Where(c => string.Equals(c.Coincide, edicion, StringComparison.OrdinalIgnoreCase)).ToList();
        d.ClaveSugerida = candidatas.FirstOrDefault(c => c.Build == build)
                          ?? candidatas.FirstOrDefault(c => c.Build == 0);
        if (d.ClaveSugerida == null && edicion.StartsWith("Server", StringComparison.OrdinalIgnoreCase))
            d.Advertencia = "No hay una clave configurada para esta versión de Windows Server.";
        return d;
    }

    public List<Licencia> Licencias() => Wmi.Consultar(Wmi.AppWindows);

    public Comando Estado() => Comando.Script("Estado detallado de la licencia de Windows", Slmgr, "/dlv");

    public Comando[] DesinstalarClave(Licencia l) => new[]
    {
        Comando.Script("Desinstalar la clave de Windows", Slmgr, "/upk"),
        Comando.Script("Borrar la clave del registro", Slmgr, "/cpky"),
    };

    public Comando InstalarClave(string clave) => Comando.Script("Instalar clave de producto", Slmgr, "/ipk " + clave);

    public Comando[] ConfigurarKms(string host, int puerto) => new[]
    {
        Comando.Script("Configurar servidor KMS", Slmgr, $"/skms {host}:{puerto}"),
    };

    public Comando[] BorrarKms() => new[]
    {
        Comando.Script("Quitar servidor KMS (usar descubrimiento DNS)", Slmgr, "/ckms"),
    };

    public Comando Activar() => Comando.Script("Activar Windows", Slmgr, "/ato");
}
