using System.Windows.Forms;

namespace InstaladorOffice;

sealed class ServidorKms
{
    public string Nombre { get; }
    public string Host { get; }
    public int Puerto { get; }

    public ServidorKms(string nombre, string host, int puerto = 1688)
    {
        Nombre = nombre; Host = host; Puerto = puerto;
    }

    public override string ToString() => $"{Nombre}  —  {Host}:{Puerto}";
}

sealed class ClaveProducto
{
    public string Nombre { get; }
    public string Clave { get; }
    /// <summary>ProductReleaseId (Office) o EditionID (Windows) para preseleccionar.</summary>
    public string Coincide { get; }
    /// <summary>Build de Windows para distinguir versiones de Server; 0 = cualquiera.</summary>
    public int Build { get; }

    public ClaveProducto(string nombre, string clave, string coincide = null, int build = 0)
    {
        Nombre = nombre; Clave = clave; Coincide = coincide; Build = build;
    }

    public override string ToString() => $"{Nombre}  —  {Clave}";
}

// ---------------------------------------------------------------------------
// Catálogo del instalador (mismo contenido que instalar.cmd y los XML)
// ---------------------------------------------------------------------------
sealed class Producto
{
    public string Id { get; }
    public string Nombre { get; }
    public bool Volumen { get; }
    public int Anio { get; }
    public bool Outlook { get; } // Outlook marcado por defecto

    public Producto(string id, string nombre, bool volumen, int anio, bool outlook)
    {
        Id = id; Nombre = nombre; Volumen = volumen; Anio = anio; Outlook = outlook;
    }

    public string Canal => Volumen ? "PerpetualVL" + Anio : null;

    public override string ToString() => Anio < 2024 ? Nombre + "   (sin soporte)" : Nombre;
}

sealed class Programa
{
    public string Id { get; }
    public string Etiqueta { get; }
    public CheckBox Casilla { get; set; }

    public Programa(string id, string etiqueta)
    {
        Id = id; Etiqueta = etiqueta;
    }
}

sealed class Idioma
{
    public string Codigo { get; }
    public string Nombre { get; }

    public Idioma(string codigo, string nombre)
    {
        Codigo = codigo; Nombre = nombre;
    }

    public override string ToString() => $"{Nombre}  ({Codigo})";
}

static class Catalogo
{
    public static readonly Producto[] Productos =
    {
        new("ProPlus2024Retail",      "Office Profesional Plus 2024",      false, 2024, true),
        new("Professional2024Retail", "Office Profesional 2024",           false, 2024, true),
        new("HomeBusiness2024Retail", "Office Hogar y Empresas 2024",      false, 2024, true),
        new("Home2024Retail",         "Office Hogar 2024",                 false, 2024, false),
        new("ProPlus2021Retail",      "Office Profesional Plus 2021",      false, 2021, true),
        new("Professional2021Retail", "Office Profesional 2021",           false, 2021, true),
        new("HomeBusiness2021Retail", "Office Hogar y Empresas 2021",      false, 2021, true),
        new("HomeStudent2021Retail",  "Office Hogar y Estudiantes 2021",   false, 2021, false),
        new("Personal2021Retail",     "Office Personal 2021",              false, 2021, false),
        new("ProPlus2019Retail",      "Office Profesional Plus 2019",      false, 2019, true),
        new("Professional2019Retail", "Office Profesional 2019",           false, 2019, true),
        new("HomeBusiness2019Retail", "Office Hogar y Empresas 2019",      false, 2019, true),
        new("HomeStudent2019Retail",  "Office Hogar y Estudiantes 2019",   false, 2019, false),
        new("Personal2019Retail",     "Office Personal 2019",              false, 2019, false),
        new("ProPlus2024Volume",      "Office LTSC Profesional Plus 2024", true,  2024, true),
        new("Standard2024Volume",     "Office LTSC Standard 2024",         true,  2024, true),
        new("ProPlus2021Volume",      "Office LTSC Profesional Plus 2021", true,  2021, true),
        new("Standard2021Volume",     "Office LTSC Standard 2021",         true,  2021, true),
        new("ProPlus2019Volume",      "Office Profesional Plus 2019",      true,  2019, true),
        new("Standard2019Volume",     "Office Standard 2019",              true,  2019, true),
    };

    // Primera columna: programas principales. Mismos ID de ExcludeApp que instalar.cmd.
    public static Programa[] CrearProgramas() => new Programa[]
    {
        new("Word",       "Word"),
        new("Excel",      "Excel"),
        new("PowerPoint", "PowerPoint"),
        new("Outlook",    "Outlook"),
        new("OneDrive",   "OneDrive"),
        new("Access",     "Access"),
        new("Publisher",  "Publisher"),
        new("OneNote",    "OneNote"),
        new("Lync",       "Skype Empresarial"),
        new("Groove",     "OneDrive Empresa (Groove)"),
    };

    public static readonly Idioma[] Idiomas =
    {
        new("es-mx",   "Español (México)"),
        new("es-es",   "Español (España)"),
        new("en-us",   "Inglés (EE. UU.)"),
        new("pt-br",   "Portugués (Brasil)"),
        new("MatchOS", "Igual que Windows"),
    };
}
