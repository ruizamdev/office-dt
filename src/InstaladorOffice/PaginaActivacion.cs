using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InstaladorOffice;

/// <summary>Pestaña de activación por volumen (KMS/MAK). La misma interfaz sirve para Office y Windows.</summary>
sealed class PaginaActivacion : UserControl
{
    static readonly Regex ClaveRegex = new(@"[A-Z0-9]{5}(-[A-Z0-9]{5}){4}");
    static readonly Regex HostRegex = new(@"^\s*([A-Za-z0-9][A-Za-z0-9.\-]{0,252})(?::(\d{1,5}))?\s*$");
    // ospp.vbs: «ERROR CODE: 0x…»; slmgr.vbs: «Error: 0x…»
    static readonly Regex ErrorRegex = new(@"(ERROR CODE|Error)\s*:\s*0x[0-9A-F]{8}", RegexOptions.IgnoreCase);

    readonly IActivador activador;
    readonly ToolTip tips = new();

    Label lblResumen, lblAviso;
    Button btnDetectar, btnEstado, btnInstalarClave, btnKms, btnBorrarKms, btnDesinstalar, btnDesinstalarTodas, btnActivar, btnTodo;
    ComboBox cmbClave, cmbKms;
    ListView lvLicencias;
    GroupBox grpAcciones;
    PanelConsola consola;

    readonly List<Button> botonesExtra = new();
    Deteccion deteccion;
    List<Licencia> licencias = new();
    bool detectadoUnaVez;

    public bool Ocupado { get; private set; }

    public PaginaActivacion(IActivador activador)
    {
        this.activador = activador;
        AutoScaleMode = AutoScaleMode.Inherit;
        Size = new Size(796, 600);
        Construir();
        ActualizarBotones();

        VisibleChanged += async (_, _) =>
        {
            if (Visible && !detectadoUnaVez && !Ocupado)
            {
                detectadoUnaVez = true;
                await DetectarAsync();
            }
        };
    }

    // ------------------------------------------------------------------ UI

    void Construir()
    {
        string n = activador.Nombre;

        var grpDetect = new GroupBox { Text = $"{n} detectado", Bounds = new Rectangle(8, 4, 780, 108) };
        lblResumen = new Label { Bounds = new Rectangle(12, 22, 600, 50), Text = "Pulsa «Detectar»." };
        lblAviso = new Label { Bounds = new Rectangle(12, 74, 600, 30) };
        btnDetectar = new Button { Text = "Detectar", Bounds = new Rectangle(626, 20, 142, 30) };
        btnDetectar.Click += async (_, _) => await DetectarAsync();
        btnEstado = new Button { Text = "Ver estado completo", Bounds = new Rectangle(626, 58, 142, 30) };
        btnEstado.Click += async (_, _) => await EjecutarAsync(new[] { activador.Estado() }, refrescar: false);
        tips.SetToolTip(btnEstado, n == "Office" ? "ospp.vbs /dstatus" : "slmgr.vbs /dlv");
        grpDetect.Controls.AddRange(new Control[] { lblResumen, lblAviso, btnDetectar, btnEstado });

        if (activador is ActivadorOffice office)
        {
            // Para instalaciones en rutas no estándar: elegir OSPP.VBS a mano
            lblResumen.Width = lblAviso.Width = 456;
            var btnOspp = new Button { Text = "Elegir ospp.vbs…", Bounds = new Rectangle(474, 20, 142, 30) };
            btnOspp.Click += async (_, _) =>
            {
                using var dlg = new OpenFileDialog
                {
                    Title = "Ubicación de OSPP.VBS",
                    Filter = "OSPP.VBS|ospp.vbs|Scripts VBS (*.vbs)|*.vbs",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                };
                if (dlg.ShowDialog(FindForm()) != DialogResult.OK) return;
                office.OsppManual = dlg.FileName;
                await DetectarAsync();
            };
            tips.SetToolTip(btnOspp, "Úsalo si Office está instalado en una ruta que no se detecta automáticamente.");
            grpDetect.Controls.Add(btnOspp);
            botonesExtra.Add(btnOspp);
        }

        lvLicencias = new ListView
        {
            Bounds = new Rectangle(8, 118, 780, 128),
            View = View.Details,
            FullRowSelect = true,
            MultiSelect = false,
            HideSelection = false,
            GridLines = true,
        };
        lvLicencias.Columns.Add("Licencia instalada", 270);
        lvLicencias.Columns.Add("Canal", 130);
        lvLicencias.Columns.Add("Estado", 150);
        lvLicencias.Columns.Add("Clave", 60);
        lvLicencias.Columns.Add("Servidor KMS", 100);
        lvLicencias.Columns.Add("Gracia", 60);
        lvLicencias.SelectedIndexChanged += (_, _) => ActualizarBotones();

        grpAcciones = new GroupBox { Text = "Activación por volumen", Bounds = new Rectangle(8, 252, 780, 132) };

        var lblClave = new Label { Text = "Clave:", Bounds = new Rectangle(12, 28, 88, 20) };
        cmbClave = new ComboBox { Bounds = new Rectangle(100, 24, 482, 24), DropDownStyle = ComboBoxStyle.DropDown };
        cmbClave.Items.AddRange(activador.Claves);
        cmbClave.TextChanged += (_, _) => ActualizarBotones();
        tips.SetToolTip(cmbClave, "Elige una clave de la lista o escribe/pega otra (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX).");
        btnInstalarClave = new Button { Text = "Instalar clave", Bounds = new Rectangle(590, 22, 178, 28) };
        btnInstalarClave.Click += async (_, _) => await InstalarClaveAsync();

        var lblKms = new Label { Text = "Servidor KMS:", Bounds = new Rectangle(12, 62, 88, 20) };
        cmbKms = new ComboBox { Bounds = new Rectangle(100, 58, 482, 24), DropDownStyle = ComboBoxStyle.DropDown };
        cmbKms.Items.AddRange(Configuracion.ServidoresKms);
        if (cmbKms.Items.Count > 0) cmbKms.SelectedIndex = 0;
        cmbKms.TextChanged += (_, _) => ActualizarBotones();
        tips.SetToolTip(cmbKms, "Elige un servidor de la lista o escribe host o host:puerto.");
        btnKms = new Button { Text = "Configurar", Bounds = new Rectangle(590, 56, 86, 28) };
        btnKms.Click += async (_, _) => await ConfigurarKmsAsync();
        btnBorrarKms = new Button { Text = "Quitar", Bounds = new Rectangle(682, 56, 86, 28) };
        btnBorrarKms.Click += async (_, _) => await EjecutarAsync(activador.BorrarKms());
        tips.SetToolTip(btnBorrarKms, "Quita el servidor configurado; Windows lo buscará por DNS.");

        btnDesinstalar = new Button
        {
            Text = activador.VariasClaves ? "Desinstalar seleccionada" : "Desinstalar clave",
            Bounds = new Rectangle(12, 92, 170, 30),
        };
        btnDesinstalar.Click += async (_, _) => await DesinstalarAsync(todas: false);
        btnDesinstalarTodas = new Button
        {
            Text = "Desinstalar todas",
            Bounds = new Rectangle(188, 92, 130, 30),
            Visible = activador.VariasClaves,
        };
        btnDesinstalarTodas.Click += async (_, _) => await DesinstalarAsync(todas: true);
        btnActivar = new Button { Text = "Activar", Bounds = new Rectangle(452, 92, 130, 30) };
        btnActivar.Click += async (_, _) => await EjecutarAsync(new[] { activador.Activar() });
        btnTodo = new Button { Text = "Clave + KMS + Activar", Bounds = new Rectangle(590, 92, 178, 30) };
        btnTodo.Font = new Font(Font, FontStyle.Bold);
        btnTodo.Click += async (_, _) => await TodoEnUnoAsync();
        tips.SetToolTip(btnTodo, "Instala la clave elegida, configura el servidor KMS elegido y activa.");

        grpAcciones.Controls.AddRange(new Control[]
        {
            lblClave, cmbClave, btnInstalarClave, lblKms, cmbKms, btnKms, btnBorrarKms,
            btnDesinstalar, btnDesinstalarTodas, btnActivar, btnTodo,
        });

        consola = new PanelConsola { Bounds = new Rectangle(8, 392, 780, 200) };

        Controls.AddRange(new Control[] { grpDetect, lvLicencias, grpAcciones, consola });
    }

    void ActualizarBotones()
    {
        bool libre = !Ocupado;
        bool permitido = deteccion?.Permitido == true;
        btnDetectar.Enabled = libre;
        foreach (var b in botonesExtra) b.Enabled = libre;
        btnEstado.Enabled = libre && deteccion != null && (permitido || activador is ActivadorWindows);
        lvLicencias.Enabled = libre;

        btnInstalarClave.Enabled = libre && permitido && LeerClave() != null;
        btnKms.Enabled = libre && permitido && LeerKms() != null;
        btnBorrarKms.Enabled = libre && permitido;
        btnActivar.Enabled = libre && permitido;
        btnTodo.Enabled = btnInstalarClave.Enabled && btnKms.Enabled;
        cmbClave.Enabled = cmbKms.Enabled = libre;

        btnDesinstalar.Enabled = libre && deteccion != null &&
            (activador.VariasClaves ? Seleccionada != null : licencias.Count > 0);
        btnDesinstalarTodas.Enabled = libre && deteccion != null && licencias.Count > 0;
    }

    Licencia Seleccionada => lvLicencias.SelectedItems.Count > 0 ? lvLicencias.SelectedItems[0].Tag as Licencia : null;

    string LeerClave()
    {
        if (cmbClave.SelectedItem is ClaveProducto c && cmbClave.Text == c.ToString()) return c.Clave;
        var m = ClaveRegex.Match(cmbClave.Text.ToUpperInvariant());
        return m.Success ? m.Value : null;
    }

    (string host, int puerto)? LeerKms()
    {
        if (cmbKms.SelectedItem is ServidorKms s && cmbKms.Text == s.ToString()) return (s.Host, s.Puerto);
        string texto = cmbKms.Text;
        int i = texto.LastIndexOf('—');
        if (i >= 0) texto = texto.Substring(i + 1);
        var m = HostRegex.Match(texto);
        if (!m.Success) return null;
        int puerto = 1688;
        if (m.Groups[2].Success && (!int.TryParse(m.Groups[2].Value, out puerto) || puerto < 1 || puerto > 65535))
            return null;
        return (m.Groups[1].Value, puerto);
    }

    bool Confirmar(string mensaje) =>
        MessageBox.Show(FindForm(), mensaje, $"Activar {activador.Nombre}",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

    /// <summary>Si la licencia actual no es por volumen, avisa que instalar la clave cambia el canal.</summary>
    bool ConfirmarCambioDeCanal()
    {
        if (licencias.Count == 0 || licencias.Any(l => l.EsVolumen)) return true;
        string canal = licencias[0].Canal;
        return Confirmar($"{activador.Nombre} usa actualmente el canal {canal}.\n\n" +
                         "Instalar una clave por volumen cambia el canal de licencia. Hazlo solo si el cliente " +
                         "cuenta con licencias por volumen.\n\n¿Continuar?");
    }

    void SetOcupado(bool valor)
    {
        Ocupado = valor;
        UseWaitCursor = valor;
        ActualizarBotones();
    }

    // ------------------------------------------------------------ Acciones

    public async Task DetectarAsync()
    {
        SetOcupado(true);
        consola.Escribir($">> Detectando {activador.Nombre}…", PanelConsola.ColorComando);
        try
        {
            deteccion = await Task.Run(() => activador.Detectar());
            lblResumen.Text = deteccion.Resumen;
            tips.SetToolTip(lblResumen, deteccion.Resumen);
            lblAviso.Text = deteccion.Advertencia;
            lblAviso.ForeColor = deteccion.Permitido ? Color.FromArgb(176, 90, 0) : Color.FromArgb(196, 43, 28);
            foreach (string linea in deteccion.Resumen.Split('\n'))
                consola.Escribir("  " + linea);
            foreach (string linea in deteccion.Detalles)
                consola.Escribir("  " + linea, Color.Gray);
            if (deteccion.Advertencia.Length > 0)
                consola.Escribir("  AVISO: " + deteccion.Advertencia, PanelConsola.ColorAviso);

            if (deteccion.ClaveSugerida != null)
            {
                cmbClave.SelectedItem = deteccion.ClaveSugerida;
                consola.Escribir("  Clave sugerida: " + deteccion.ClaveSugerida.Nombre);
            }
            await RefrescarLicenciasAsync();
        }
        catch (Exception ex)
        {
            consola.Escribir("  ERROR: " + ex.Message, PanelConsola.ColorError);
        }
        finally
        {
            SetOcupado(false);
        }
    }

    async Task RefrescarLicenciasAsync()
    {
        try
        {
            licencias = await Task.Run(() => activador.Licencias());
        }
        catch (Exception ex)
        {
            licencias = new List<Licencia>();
            consola.Escribir("  ERROR: No se pudieron leer las licencias (WMI): " + ex.Message, PanelConsola.ColorError);
        }

        lvLicencias.BeginUpdate();
        lvLicencias.Items.Clear();
        foreach (var l in licencias)
        {
            var item = new ListViewItem(new[]
            {
                l.Nombre, l.Canal, l.EstadoTexto, l.ClaveParcial, l.KmsTexto, l.GraciaTexto,
            }) { Tag = l };
            item.ForeColor = l.Estado == 1 ? Color.FromArgb(16, 124, 16) : Color.FromArgb(176, 90, 0);
            lvLicencias.Items.Add(item);
        }
        if (lvLicencias.Items.Count > 0) lvLicencias.Items[0].Selected = true;
        lvLicencias.EndUpdate();

        consola.Escribir(licencias.Count == 0
            ? "  No hay claves de producto instaladas."
            : $"  {licencias.Count} licencia(s) con clave instalada.");
        ActualizarBotones();
    }

    /// <summary>Ejecuta los comandos en orden y muestra su salida. Devuelve false si alguno falló.</summary>
    async Task<bool> EjecutarAsync(IEnumerable<Comando> comandos, bool refrescar = true, bool detenerSiFalla = true)
    {
        SetOcupado(true);
        bool todoBien = true;
        try
        {
            foreach (var cmd in comandos)
            {
                consola.Escribir(">> " + cmd.Descripcion, PanelConsola.ColorComando);
                consola.Escribir("  > " + cmd, Color.Gray);
                bool errorEnSalida = false;
                int codigo = await Ejecutor.EjecutarAsync(cmd, linea =>
                {
                    if (ErrorRegex.IsMatch(linea)) errorEnSalida = true;
                    if (linea.Trim().Length > 0) consola.Escribir("  " + linea);
                });
                if (codigo != 0 || errorEnSalida)
                {
                    todoBien = false;
                    consola.Escribir(codigo != 0
                        ? $"  ERROR: terminó con código {codigo} (0x{codigo:X8})."
                        : "  ERROR: El comando reportó un error.", PanelConsola.ColorError);
                    if (detenerSiFalla) break;
                }
                else
                {
                    consola.Escribir("  OK", PanelConsola.ColorOk);
                }
            }
            if (refrescar) await RefrescarLicenciasAsync();
        }
        catch (Exception ex)
        {
            todoBien = false;
            consola.Escribir("  ERROR: " + ex.Message, PanelConsola.ColorError);
        }
        finally
        {
            SetOcupado(false);
        }
        return todoBien;
    }

    async Task InstalarClaveAsync()
    {
        string clave = LeerClave();
        if (clave == null || !ConfirmarCambioDeCanal()) return;
        await EjecutarAsync(new[] { activador.InstalarClave(clave) });
    }

    async Task ConfigurarKmsAsync()
    {
        var kms = LeerKms();
        if (kms == null) return;
        await EjecutarAsync(activador.ConfigurarKms(kms.Value.host, kms.Value.puerto));
    }

    async Task DesinstalarAsync(bool todas)
    {
        List<Licencia> objetivo;
        string mensaje;
        if (!activador.VariasClaves)
        {
            objetivo = licencias.Take(1).ToList();
            mensaje = "¿Desinstalar la clave de producto de Windows?\n\nWindows quedará sin activar hasta instalar otra clave.";
        }
        else if (todas)
        {
            objetivo = licencias.ToList();
            mensaje = $"¿Desinstalar las {objetivo.Count} claves de Office instaladas?\n\nOffice quedará sin activar hasta instalar otra clave.";
        }
        else
        {
            var sel = Seleccionada;
            if (sel == null) return;
            objetivo = new List<Licencia> { sel };
            mensaje = $"¿Desinstalar la clave …{sel.ClaveParcial} de\n{sel.Nombre}?";
        }
        if (objetivo.Count == 0 || !Confirmar(mensaje)) return;
        await EjecutarAsync(objetivo.SelectMany(l => activador.DesinstalarClave(l)), detenerSiFalla: false);
    }

    async Task TodoEnUnoAsync()
    {
        string clave = LeerClave();
        var kms = LeerKms();
        if (clave == null || kms == null || !ConfirmarCambioDeCanal()) return;
        if (!Confirmar($"Se instalará la clave {clave},\nse usará el servidor KMS {kms.Value.host}:{kms.Value.puerto}\n" +
                       $"y se activará {activador.Nombre}.\n\n¿Continuar?"))
            return;

        var comandos = new List<Comando> { activador.InstalarClave(clave) };
        comandos.AddRange(activador.ConfigurarKms(kms.Value.host, kms.Value.puerto));
        comandos.Add(activador.Activar());
        bool ok = await EjecutarAsync(comandos);
        consola.Escribir(ok ? $"OK: {activador.Nombre} activado." : $"ERROR: No se completó la activación de {activador.Nombre}.",
            ok ? PanelConsola.ColorOk : PanelConsola.ColorError);
    }
}
