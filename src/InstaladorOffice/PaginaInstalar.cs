using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace InstaladorOffice;

/// <summary>Pestaña «Instalar Office»: elige edición y programas, genera el XML y ejecuta setup.exe.</summary>
sealed class PaginaInstalar : UserControl
{
    readonly string carpetaApp = AppDomain.CurrentDomain.BaseDirectory;
    readonly string carpetaXml;
    readonly string rutaSetup;
    readonly Programa[] programas = Catalogo.CrearProgramas();
    readonly ToolTip tips = new();

    RadioButton rbRetail, rbVolumen, rb64, rb32;
    ListBox lstProductos;
    Label lblProducto, lblSoporte, lblOdt, lblEstado;
    ComboBox cmbIdioma;
    TextBox txtXml;
    LinkLabel lnkManual;
    ProgressBar progreso;
    Button btnOdt, btnGuardar, btnInstalar;
    readonly List<GroupBox> grupos = new();

    bool odtFallo, odtIntentado;
    public bool Ocupado { get; private set; }
    public bool Instalando { get; private set; }

    public PaginaInstalar()
    {
        carpetaXml = Path.Combine(carpetaApp, "Personalizado");
        rutaSetup = Path.Combine(carpetaApp, "setup.exe");
        AutoScaleMode = AutoScaleMode.Inherit;
        Size = new Size(796, 600);

        ConstruirVersion();
        ConstruirOpciones();
        ConstruirProgramas();
        ConstruirVistaPrevia();
        ConstruirInferior();
        LlenarProductos();
        ActualizarOdt();

        VisibleChanged += async (_, _) =>
        {
            if (Visible && !odtIntentado && !File.Exists(rutaSetup))
            {
                odtIntentado = true;
                await DescargarOdtAsync();
            }
        };
    }

    // ------------------------------------------------------------------ UI

    GroupBox Grupo(string texto, int x, int y, int w, int h)
    {
        var g = new GroupBox { Text = texto, Bounds = new Rectangle(x, y, w, h) };
        Controls.Add(g);
        grupos.Add(g);
        return g;
    }

    void ConstruirVersion()
    {
        var g = Grupo("1. Versión de Office", 8, 4, 376, 300);

        rbRetail = new RadioButton { Text = "Retail (clave individual)", Bounds = new Rectangle(14, 24, 175, 24), Checked = true };
        rbRetail.CheckedChanged += (_, _) => { if (rbRetail.Checked) LlenarProductos(); };
        rbVolumen = new RadioButton { Text = "Volumen (KMS / MAK)", Bounds = new Rectangle(196, 24, 170, 24) };
        rbVolumen.CheckedChanged += (_, _) => { if (rbVolumen.Checked) LlenarProductos(); };

        lstProductos = new ListBox { Bounds = new Rectangle(14, 52, 348, 184), IntegralHeight = false };
        lstProductos.SelectedIndexChanged += (_, _) => ProductoCambiado();

        lblProducto = new Label { Bounds = new Rectangle(14, 240, 348, 18), ForeColor = SystemColors.GrayText };
        lblSoporte = new Label { Bounds = new Rectangle(14, 260, 348, 34), ForeColor = Color.FromArgb(176, 90, 0) };

        g.Controls.AddRange(new Control[] { rbRetail, rbVolumen, lstProductos, lblProducto, lblSoporte });
    }

    void ConstruirOpciones()
    {
        var g = Grupo("3. Opciones", 8, 312, 376, 92);

        var lArq = new Label { Text = "Arquitectura:", Bounds = new Rectangle(14, 28, 90, 20) };
        var panel = new Panel { Bounds = new Rectangle(104, 22, 250, 28) };
        rb64 = new RadioButton { Text = "64 bits", Bounds = new Rectangle(0, 3, 80, 22), Checked = true };
        rb64.CheckedChanged += (_, _) => ActualizarVistaPrevia();
        rb32 = new RadioButton { Text = "32 bits", Bounds = new Rectangle(90, 3, 80, 22) };
        panel.Controls.AddRange(new Control[] { rb64, rb32 });

        var lIdioma = new Label { Text = "Idioma:", Bounds = new Rectangle(14, 60, 90, 20) };
        cmbIdioma = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Bounds = new Rectangle(104, 56, 250, 24) };
        cmbIdioma.Items.AddRange(Catalogo.Idiomas);
        cmbIdioma.SelectedIndex = 0;
        cmbIdioma.SelectedIndexChanged += (_, _) => ActualizarVistaPrevia();

        g.Controls.AddRange(new Control[] { lArq, panel, lIdioma, cmbIdioma });
    }

    void ConstruirProgramas()
    {
        var g = Grupo("2. Programas a instalar", 394, 4, 394, 214);

        for (int i = 0; i < programas.Length; i++)
        {
            int col = i / 5, fila = i % 5;
            var cb = new CheckBox { Text = programas[i].Etiqueta, Bounds = new Rectangle(16 + col * 186, 24 + fila * 26, 180, 24) };
            cb.CheckedChanged += (_, _) => ActualizarVistaPrevia();
            programas[i].Casilla = cb;
            g.Controls.Add(cb);
        }

        g.Controls.Add(new Label
        {
            Text = "Solo se instalan los programas que incluye la edición elegida.",
            ForeColor = SystemColors.GrayText,
            Bounds = new Rectangle(16, 154, 366, 18),
        });

        var bTodos = new Button { Text = "Todos", Bounds = new Rectangle(16, 176, 90, 26) };
        bTodos.Click += (_, _) => MarcarTodos(true);
        var bNinguno = new Button { Text = "Ninguno", Bounds = new Rectangle(112, 176, 90, 26) };
        bNinguno.Click += (_, _) => MarcarTodos(false);
        var bPredet = new Button { Text = "Predeterminados", Bounds = new Rectangle(208, 176, 120, 26) };
        bPredet.Click += (_, _) => AplicarPredeterminados();
        g.Controls.AddRange(new Control[] { bTodos, bNinguno, bPredet });
    }

    void ConstruirVistaPrevia()
    {
        var g = Grupo("Vista previa de la configuración", 394, 226, 394, 178);
        txtXml = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            BackColor = SystemColors.Window,
            Font = new Font("Consolas", 8.5F),
            Dock = DockStyle.Fill,
        };
        g.Controls.Add(txtXml);
    }

    void ConstruirInferior()
    {
        lblOdt = new Label { Bounds = new Rectangle(8, 420, 600, 20) };
        lnkManual = new LinkLabel { Text = "Descargar manualmente desde Microsoft", Bounds = new Rectangle(8, 442, 300, 20), Visible = false };
        lnkManual.LinkClicked += (_, _) => { try { Process.Start(OdtDownloader.PaginaUrl); } catch { } };

        btnOdt = new Button { Bounds = new Rectangle(628, 416, 160, 28) };
        btnOdt.Click += async (_, _) => await DescargarOdtAsync();
        tips.SetToolTip(btnOdt, "Descarga la versión más reciente de Office Deployment Tool (setup.exe).");

        progreso = new ProgressBar { Bounds = new Rectangle(8, 470, 780, 16) };
        lblEstado = new Label { Bounds = new Rectangle(8, 492, 780, 36) };

        btnGuardar = new Button { Text = "Guardar XML…", Bounds = new Rectangle(560, 540, 110, 34) };
        btnGuardar.Click += (_, _) => GuardarXml();
        btnInstalar = new Button { Text = "Instalar", Bounds = new Rectangle(678, 540, 110, 34) };
        btnInstalar.Font = new Font(Font, FontStyle.Bold);
        btnInstalar.Click += async (_, _) => await InstalarAsync();

        Controls.AddRange(new Control[] { lblOdt, lnkManual, btnOdt, progreso, lblEstado, btnGuardar, btnInstalar });
    }

    // ------------------------------------------------------- Selección

    Producto ProductoElegido => lstProductos.SelectedItem as Producto;
    string IdiomaElegido => ((Idioma)cmbIdioma.SelectedItem).Codigo;
    string Arquitectura => rb32.Checked ? "32" : "64";
    List<string> ProgramasElegidos => programas.Where(p => p.Casilla.Checked).Select(p => p.Casilla.Text).ToList();

    void LlenarProductos()
    {
        bool volumen = rbVolumen.Checked;
        lstProductos.BeginUpdate();
        lstProductos.Items.Clear();
        foreach (var p in Catalogo.Productos.Where(p => p.Volumen == volumen))
            lstProductos.Items.Add(p);
        lstProductos.EndUpdate();
        lstProductos.SelectedIndex = 0;
    }

    void ProductoCambiado()
    {
        var p = ProductoElegido;
        if (p == null) return;

        lblProducto.Text = "ID: " + p.Id + (p.Canal != null ? "   ·   Canal: " + p.Canal : "");
        lblSoporte.Text = p.Anio < 2024
            ? $"Aviso: Office {p.Anio} está fuera de soporte desde el 14/10/2025. Considera una versión 2024."
            : "";

        var publisher = programas.First(a => a.Id == "Publisher").Casilla;
        publisher.Enabled = p.Anio < 2024;
        tips.SetToolTip(publisher, p.Anio >= 2024 ? "Publisher no se incluye en Office 2024." : null);
        AplicarPredeterminados();
    }

    // Mismos valores que los XML de Retail\ y Volume\
    void AplicarPredeterminados()
    {
        var p = ProductoElegido;
        foreach (var a in programas)
        {
            bool marcado = a.Id is "Word" or "Excel" or "PowerPoint" or "OneDrive"
                           || (a.Id == "Outlook" && p?.Outlook == true);
            a.Casilla.Checked = marcado && a.Casilla.Enabled;
        }
        ActualizarVistaPrevia();
    }

    void MarcarTodos(bool valor)
    {
        foreach (var a in programas)
            a.Casilla.Checked = valor && a.Casilla.Enabled;
        ActualizarVistaPrevia();
    }

    string GenerarXml()
    {
        var p = ProductoElegido;
        if (p == null) return "";

        var producto = new XElement("Product", new XAttribute("ID", p.Id),
            new XElement("Language", new XAttribute("ID", IdiomaElegido)));
        foreach (var a in programas.Where(a => !a.Casilla.Checked))
            producto.Add(new XElement("ExcludeApp", new XAttribute("ID", a.Id)));

        var add = new XElement("Add", new XAttribute("OfficeClientEdition", Arquitectura));
        if (p.Canal != null) add.Add(new XAttribute("Channel", p.Canal));
        add.Add(producto);

        var raiz = new XElement("Configuration",
            add,
            new XElement("Updates", new XAttribute("Enabled", "TRUE")),
            new XElement("Display", new XAttribute("Level", "Full"), new XAttribute("AcceptEULA", "TRUE")));

        var sb = new StringBuilder();
        using (var w = XmlWriter.Create(sb, new XmlWriterSettings { OmitXmlDeclaration = true, Indent = true, IndentChars = "  " }))
            raiz.Save(w);
        return sb + "\r\n";
    }

    void ActualizarVistaPrevia()
    {
        if (txtXml == null) return;
        txtXml.Text = GenerarXml();
        ActualizarBotones();
    }

    string NombreXml() => $"{ProductoElegido.Id}-x{Arquitectura}-{IdiomaElegido}.xml";

    void EscribirXml(string ruta)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(ruta));
        File.WriteAllText(ruta, GenerarXml(), new UTF8Encoding(false));
    }

    bool ValidarSeleccion()
    {
        if (ProgramasElegidos.Count > 0) return true;
        MessageBox.Show(FindForm(), "Selecciona al menos un programa.", "Instalar Office",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
        return false;
    }

    // ------------------------------------------------------- Acciones

    void GuardarXml()
    {
        if (!ValidarSeleccion()) return;
        Directory.CreateDirectory(carpetaXml);
        using var dlg = new SaveFileDialog
        {
            InitialDirectory = carpetaXml,
            FileName = NombreXml(),
            Filter = "Configuración XML (*.xml)|*.xml",
        };
        if (dlg.ShowDialog(FindForm()) != DialogResult.OK) return;
        try
        {
            EscribirXml(dlg.FileName);
            lblEstado.Text = "XML guardado en " + dlg.FileName;
        }
        catch (Exception ex)
        {
            MessageBox.Show(FindForm(), "No se pudo guardar el XML:\n" + ex.Message, "Instalar Office",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    async Task InstalarAsync()
    {
        if (Ocupado || !ValidarSeleccion()) return;
        if (!File.Exists(rutaSetup))
        {
            ActualizarOdt();
            MessageBox.Show(FindForm(), "Falta setup.exe. Usa el botón «Descargar ODT».", "Instalar Office",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var p = ProductoElegido;
        string msg = "Se instalará:\n\n" +
                     $"{p.Nombre} ({p.Id})\n{Arquitectura} bits · {IdiomaElegido}\n\n" +
                     $"Programas: {string.Join(", ", ProgramasElegidos)}\n\n" +
                     "Office se descargará desde la CDN de Microsoft y puede tardar varios minutos.\n¿Continuar?";
        if (MessageBox.Show(FindForm(), msg, "Instalar Office", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        string xml = Path.Combine(carpetaXml, NombreXml());
        try { EscribirXml(xml); }
        catch (Exception ex)
        {
            MessageBox.Show(FindForm(), "No se pudo guardar el XML:\n" + ex.Message, "Instalar Office",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        Instalando = true;
        SetOcupado(true, $"Instalando {p.Nombre}… sigue el progreso en la ventana de Office.");
        int codigo;
        try
        {
            codigo = await Task.Run(() =>
            {
                var psi = new ProcessStartInfo(rutaSetup, $"/configure \"{xml}\"")
                {
                    WorkingDirectory = carpetaApp,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                using var proc = Process.Start(psi);
                proc.WaitForExit();
                return proc.ExitCode;
            });
        }
        catch (Exception ex)
        {
            Instalando = false;
            SetOcupado(false, "Error al ejecutar setup.exe.");
            MessageBox.Show(FindForm(), "No se pudo ejecutar setup.exe:\n" + ex.Message, "Instalar Office",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        Instalando = false;

        if (codigo == 0)
        {
            SetOcupado(false, "Instalación terminada. Para activarla, usa la pestaña «Activar Office».");
            MessageBox.Show(FindForm(), "Instalación terminada.", "Instalar Office", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            SetOcupado(false, $"setup.exe terminó con código {codigo}.");
            MessageBox.Show(FindForm(), $"setup.exe terminó con código {codigo}.\n\nRevisa los registros de ODT en:\n{Path.GetTempPath()}",
                "Instalar Office", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    async Task DescargarOdtAsync()
    {
        if (Ocupado) return;
        odtFallo = false;
        lnkManual.Visible = false;
        SetOcupado(true, "Preparando descarga de Office Deployment Tool…");

        IProgress<(int pct, string texto)> reporte = new Progress<(int pct, string texto)>(r =>
        {
            if (r.pct >= 0)
            {
                progreso.Style = ProgressBarStyle.Continuous;
                progreso.Value = Math.Min(100, r.pct);
            }
            else
            {
                progreso.Style = ProgressBarStyle.Marquee;
            }
            lblEstado.Text = r.texto;
        });

        try
        {
            await Task.Run(() => OdtDownloader.Descargar(rutaSetup, (pct, texto) => reporte.Report((pct, texto))));
            SetOcupado(false, "Office Deployment Tool descargado y verificado.");
        }
        catch (Exception ex)
        {
            odtFallo = true;
            SetOcupado(false, "No se pudo descargar Office Deployment Tool: " + ex.Message.Replace("\n", " "));
        }
        ActualizarOdt();
    }

    // ------------------------------------------------------- Estado

    void SetOcupado(bool ocupado, string estado)
    {
        Ocupado = ocupado;
        foreach (var g in grupos) g.Enabled = !ocupado;
        progreso.Style = ocupado ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
        progreso.MarqueeAnimationSpeed = ocupado ? 30 : 0;
        if (!ocupado) progreso.Value = 0;
        if (!string.IsNullOrEmpty(estado)) lblEstado.Text = estado;
        btnOdt.Enabled = !ocupado;
        ActualizarBotones();
    }

    void ActualizarOdt()
    {
        if (File.Exists(rutaSetup))
        {
            string ver = "";
            try { ver = FileVersionInfo.GetVersionInfo(rutaSetup).FileVersion; } catch { }
            lblOdt.Text = "Office Deployment Tool listo" + (string.IsNullOrEmpty(ver) ? "" : $" (setup.exe {ver})");
            lblOdt.ForeColor = Color.FromArgb(16, 124, 16);
            btnOdt.Text = "Actualizar ODT";
        }
        else
        {
            lblOdt.Text = odtFallo
                ? "Falta setup.exe y la descarga automática falló."
                : "Falta setup.exe (Office Deployment Tool).";
            lblOdt.ForeColor = Color.FromArgb(196, 43, 28);
            btnOdt.Text = odtFallo ? "Reintentar descarga" : "Descargar ODT";
        }
        lnkManual.Visible = odtFallo;
        ActualizarBotones();
    }

    void ActualizarBotones()
    {
        if (btnInstalar == null) return;
        bool hayProgramas = ProgramasElegidos.Count > 0;
        btnInstalar.Enabled = !Ocupado && hayProgramas && File.Exists(rutaSetup);
        btnGuardar.Enabled = !Ocupado && hayProgramas;
    }
}
