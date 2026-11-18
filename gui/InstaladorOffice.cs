// Instalador grafico de Office 2019 / 2021 / 2024 con Office Deployment Tool.
// Compilar con build.cmd (csc.exe de .NET Framework 4.x, C# 5).
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

[assembly: System.Reflection.AssemblyTitle("Instalador de Office")]
[assembly: System.Reflection.AssemblyProduct("Instalador de Office (ODT)")]
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace InstaladorOffice
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    // ------------------------------------------------------------------
    // Catalogos
    // ------------------------------------------------------------------
    class Product
    {
        public string Id;
        public string Name;
        public bool Volume;
        public int Year;
        public bool Outlook; // Outlook marcado por defecto

        public Product(string id, string name, bool volume, int year, bool outlook)
        {
            Id = id; Name = name; Volume = volume; Year = year; Outlook = outlook;
        }

        public string Channel
        {
            get { return Volume ? "PerpetualVL" + Year : null; }
        }

        public override string ToString()
        {
            return Year < 2024 ? Name + "   (sin soporte)" : Name;
        }
    }

    class AppItem
    {
        public string Id;
        public string Label;
        public CheckBox Box;

        public AppItem(string id, string label)
        {
            Id = id; Label = label;
        }
    }

    class LangItem
    {
        public string Code;
        public string Label;

        public LangItem(string code, string label)
        {
            Code = code; Label = label;
        }

        public override string ToString()
        {
            return Label + "  (" + Code + ")";
        }
    }

    static class Catalog
    {
        public static readonly Product[] Products = new Product[]
        {
            new Product("ProPlus2024Retail",      "Office Profesional Plus 2024",      false, 2024, true),
            new Product("Professional2024Retail", "Office Profesional 2024",           false, 2024, true),
            new Product("HomeBusiness2024Retail", "Office Hogar y Empresas 2024",      false, 2024, true),
            new Product("Home2024Retail",         "Office Hogar 2024",                 false, 2024, false),
            new Product("ProPlus2021Retail",      "Office Profesional Plus 2021",      false, 2021, true),
            new Product("Professional2021Retail", "Office Profesional 2021",           false, 2021, true),
            new Product("HomeBusiness2021Retail", "Office Hogar y Empresas 2021",      false, 2021, true),
            new Product("HomeStudent2021Retail",  "Office Hogar y Estudiantes 2021",   false, 2021, false),
            new Product("Personal2021Retail",     "Office Personal 2021",              false, 2021, false),
            new Product("ProPlus2019Retail",      "Office Profesional Plus 2019",      false, 2019, true),
            new Product("Professional2019Retail", "Office Profesional 2019",           false, 2019, true),
            new Product("HomeBusiness2019Retail", "Office Hogar y Empresas 2019",      false, 2019, true),
            new Product("HomeStudent2019Retail",  "Office Hogar y Estudiantes 2019",   false, 2019, false),
            new Product("Personal2019Retail",     "Office Personal 2019",              false, 2019, false),
            new Product("ProPlus2024Volume",      "Office LTSC Profesional Plus 2024", true,  2024, true),
            new Product("Standard2024Volume",     "Office LTSC Standard 2024",         true,  2024, true),
            new Product("ProPlus2021Volume",      "Office LTSC Profesional Plus 2021", true,  2021, true),
            new Product("Standard2021Volume",     "Office LTSC Standard 2021",         true,  2021, true),
            new Product("ProPlus2019Volume",      "Office Profesional Plus 2019",      true,  2019, true),
            new Product("Standard2019Volume",     "Office Standard 2019",              true,  2019, true),
        };

        // Mismo orden que instalar.cmd. Primera columna: los programas principales.
        public static AppItem[] CreateApps()
        {
            return new AppItem[]
            {
                new AppItem("Word",       "Word"),
                new AppItem("Excel",      "Excel"),
                new AppItem("PowerPoint", "PowerPoint"),
                new AppItem("Outlook",    "Outlook"),
                new AppItem("OneDrive",   "OneDrive"),
                new AppItem("Access",     "Access"),
                new AppItem("Publisher",  "Publisher"),
                new AppItem("OneNote",    "OneNote"),
                new AppItem("Lync",       "Skype Empresarial"),
                new AppItem("Groove",     "OneDrive Empresa (Groove)"),
            };
        }

        public static readonly LangItem[] Languages = new LangItem[]
        {
            new LangItem("es-mx",   "Español (México)"),
            new LangItem("es-es",   "Español (España)"),
            new LangItem("en-us",   "Inglés (EE. UU.)"),
            new LangItem("pt-br",   "Portugués (Brasil)"),
            new LangItem("MatchOS", "Igual que Windows"),
        };
    }

    // ------------------------------------------------------------------
    // Descarga y extraccion de Office Deployment Tool
    // ------------------------------------------------------------------
    static class OdtDownloader
    {
        public const string PageUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=49117";
        const string FallbackUrl = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117";
        const string UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36";
        static readonly Regex LinkRegex = new Regex(
            @"https://download\.microsoft\.com/download/[\w\-/\.]*officedeploymenttool[\w\-\.]*\.exe",
            RegexOptions.IgnoreCase);

        // Descarga el ODT, verifica firmas, extrae setup.exe y lo copia a destSetup.
        // report(porcentaje o -1, mensaje)
        public static void Run(string destSetup, Action<int, string> report)
        {
            ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; // TLS 1.2

            report(-1, "Buscando la versión más reciente de Office Deployment Tool...");
            string url = FindDownloadUrl();

            string temp = Path.Combine(Path.GetTempPath(), "odt-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            try
            {
                string odtExe = Path.Combine(temp, Path.GetFileName(new Uri(url).AbsolutePath));
                DownloadFile(url, odtExe, report);

                report(-1, "Verificando la firma digital de Microsoft...");
                VerifyMicrosoftSignature(odtExe);

                report(-1, "Extrayendo setup.exe...");
                string extractDir = Path.Combine(temp, "extract");
                Directory.CreateDirectory(extractDir);
                ProcessStartInfo psi = new ProcessStartInfo(odtExe, "/quiet /extract:\"" + extractDir + "\"");
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                using (Process p = Process.Start(psi))
                {
                    p.WaitForExit();
                    if (p.ExitCode != 0)
                        throw new Exception("La extracción del ODT terminó con código " + p.ExitCode + ".");
                }

                string setup = Path.Combine(extractDir, "setup.exe");
                if (!File.Exists(setup))
                    throw new Exception("La extracción no generó setup.exe.");
                VerifyMicrosoftSignature(setup);

                File.Copy(setup, destSetup, true);
                report(100, "Office Deployment Tool listo.");
            }
            finally
            {
                try { Directory.Delete(temp, true); } catch { }
            }
        }

        static string FindDownloadUrl()
        {
            Exception last = null;
            foreach (string page in new string[] { PageUrl, FallbackUrl })
            {
                try
                {
                    using (WebClient wc = new WebClient())
                    {
                        wc.Encoding = Encoding.UTF8;
                        wc.Headers[HttpRequestHeader.UserAgent] = UserAgent;
                        wc.Headers[HttpRequestHeader.Accept] = "text/html,application/xhtml+xml";
                        wc.Headers[HttpRequestHeader.AcceptLanguage] = "en-US,en;q=0.9";
                        Match m = LinkRegex.Match(wc.DownloadString(page));
                        if (m.Success) return m.Value;
                    }
                }
                catch (Exception ex) { last = ex; }
            }
            throw new Exception("No se encontró el enlace de descarga del ODT en la página de Microsoft." +
                (last != null ? "\n" + last.Message : ""));
        }

        static void DownloadFile(string url, string dest, Action<int, string> report)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.UserAgent = UserAgent;
            using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
            using (Stream input = resp.GetResponseStream())
            using (FileStream output = File.Create(dest))
            {
                long total = resp.ContentLength;
                long done = 0;
                int lastPct = -2;
                byte[] buf = new byte[81920];
                int n;
                while ((n = input.Read(buf, 0, buf.Length)) > 0)
                {
                    output.Write(buf, 0, n);
                    done += n;
                    int pct = total > 0 ? (int)(done * 100 / total) : -1;
                    if (pct != lastPct)
                    {
                        lastPct = pct;
                        report(pct, "Descargando Office Deployment Tool... " +
                            (pct >= 0 ? pct + " %" : (done / 1024) + " KB"));
                    }
                }
            }
        }

        // ---- Verificacion Authenticode (WinVerifyTrust) ----
        static readonly Guid WintrustActionGenericVerifyV2 = new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");

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

        public static void VerifyMicrosoftSignature(string path)
        {
            IntPtr pPath = Marshal.StringToCoTaskMemUni(path);
            IntPtr pFile = IntPtr.Zero;
            IntPtr pData = IntPtr.Zero;
            int result;
            try
            {
                WinTrustFileInfo fi = new WinTrustFileInfo();
                fi.cbStruct = (uint)Marshal.SizeOf(typeof(WinTrustFileInfo));
                fi.pcwszFilePath = pPath;
                pFile = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(WinTrustFileInfo)));
                Marshal.StructureToPtr(fi, pFile, false);

                WinTrustData wd = new WinTrustData();
                wd.cbStruct = (uint)Marshal.SizeOf(typeof(WinTrustData));
                wd.dwUIChoice = 2;          // WTD_UI_NONE
                wd.fdwRevocationChecks = 0; // WTD_REVOKE_NONE
                wd.dwUnionChoice = 1;       // WTD_CHOICE_FILE
                wd.pFile = pFile;
                pData = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(WinTrustData)));
                Marshal.StructureToPtr(wd, pData, false);

                result = WinVerifyTrust(new IntPtr(-1), WintrustActionGenericVerifyV2, pData);
            }
            finally
            {
                if (pData != IntPtr.Zero) Marshal.FreeCoTaskMem(pData);
                if (pFile != IntPtr.Zero) Marshal.FreeCoTaskMem(pFile);
                Marshal.FreeCoTaskMem(pPath);
            }

            if (result != 0)
                throw new Exception("La firma digital de " + Path.GetFileName(path) +
                    " no es válida (0x" + result.ToString("X8") + "). Se descartó el archivo.");

            X509Certificate2 cert = new X509Certificate2(X509Certificate.CreateFromSignedFile(path));
            if (cert.Subject.IndexOf("O=Microsoft Corporation", StringComparison.OrdinalIgnoreCase) < 0)
                throw new Exception(Path.GetFileName(path) + " no está firmado por Microsoft Corporation.");
        }
    }

    // ------------------------------------------------------------------
    // Ventana principal
    // ------------------------------------------------------------------
    class MainForm : Form
    {
        readonly string appDir = AppDomain.CurrentDomain.BaseDirectory;
        readonly string outDir;
        readonly string setupPath;
        readonly AppItem[] apps = Catalog.CreateApps();

        RadioButton rbRetail, rbVolume, rb64, rb32;
        ListBox lstProducts;
        Label lblProduct, lblSupport, lblOdt, lblStatus;
        ComboBox cmbLang;
        TextBox txtXml;
        LinkLabel lnkManual;
        ProgressBar progress;
        Button btnOdt, btnSave, btnInstall, btnExit;
        ToolTip tips = new ToolTip();

        BackgroundWorker odtWorker, installWorker;
        bool odtFailed;

        public MainForm()
        {
            outDir = Path.Combine(appDir, "Personalizado");
            setupPath = Path.Combine(appDir, "setup.exe");

            Text = "Instalador de Office 2019 – 2024";
            Font = new Font("Segoe UI", 9F);
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(780, 560);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            SuspendLayout();
            BuildVersionGroup();
            BuildOptionsGroup();
            BuildAppsGroup();
            BuildXmlGroup();
            BuildBottom();
            ResumeLayout(false);

            odtWorker = new BackgroundWorker();
            odtWorker.WorkerReportsProgress = true;
            odtWorker.DoWork += OdtWorker_DoWork;
            odtWorker.ProgressChanged += OdtWorker_ProgressChanged;
            odtWorker.RunWorkerCompleted += OdtWorker_Completed;

            installWorker = new BackgroundWorker();
            installWorker.DoWork += InstallWorker_DoWork;
            installWorker.RunWorkerCompleted += InstallWorker_Completed;

            FillProducts();
            Shown += delegate
            {
                if (!File.Exists(setupPath)) StartOdtDownload();
            };
            FormClosing += MainForm_FormClosing;
        }

        // ---------------- Construccion de la interfaz ----------------

        void BuildVersionGroup()
        {
            GroupBox g = Group("1. Versión de Office", 12, 10, 364, 300);

            rbRetail = new RadioButton();
            rbRetail.Text = "Retail (clave individual)";
            rbRetail.SetBounds(14, 24, 170, 24);
            rbRetail.Checked = true;
            rbRetail.CheckedChanged += delegate { if (rbRetail.Checked) FillProducts(); };

            rbVolume = new RadioButton();
            rbVolume.Text = "Volumen (KMS / MAK)";
            rbVolume.SetBounds(190, 24, 160, 24);
            rbVolume.CheckedChanged += delegate { if (rbVolume.Checked) FillProducts(); };

            lstProducts = new ListBox();
            lstProducts.SetBounds(14, 52, 336, 184);
            lstProducts.IntegralHeight = false;
            lstProducts.SelectedIndexChanged += delegate { ProductChanged(); };

            lblProduct = new Label();
            lblProduct.SetBounds(14, 240, 336, 18);
            lblProduct.ForeColor = SystemColors.GrayText;

            lblSupport = new Label();
            lblSupport.SetBounds(14, 260, 336, 34);
            lblSupport.ForeColor = Color.FromArgb(176, 90, 0);

            g.Controls.AddRange(new Control[] { rbRetail, rbVolume, lstProducts, lblProduct, lblSupport });
        }

        void BuildOptionsGroup()
        {
            GroupBox g = Group("3. Opciones", 12, 318, 364, 92);

            Label l1 = new Label();
            l1.Text = "Arquitectura:";
            l1.SetBounds(14, 28, 90, 20);

            // Panel propio para que los radio de arquitectura no se mezclen con otros
            Panel archPanel = new Panel();
            archPanel.SetBounds(104, 22, 240, 28);
            rb64 = new RadioButton();
            rb64.Text = "64 bits";
            rb64.SetBounds(0, 3, 80, 22);
            rb64.Checked = true;
            rb64.CheckedChanged += delegate { UpdatePreview(); };
            rb32 = new RadioButton();
            rb32.Text = "32 bits";
            rb32.SetBounds(90, 3, 80, 22);
            archPanel.Controls.AddRange(new Control[] { rb64, rb32 });

            Label l2 = new Label();
            l2.Text = "Idioma:";
            l2.SetBounds(14, 60, 90, 20);

            cmbLang = new ComboBox();
            cmbLang.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLang.SetBounds(104, 56, 240, 24);
            cmbLang.Items.AddRange(Catalog.Languages);
            cmbLang.SelectedIndex = 0;
            cmbLang.SelectedIndexChanged += delegate { UpdatePreview(); };

            g.Controls.AddRange(new Control[] { l1, archPanel, l2, cmbLang });
        }

        void BuildAppsGroup()
        {
            GroupBox g = Group("2. Programas a instalar", 388, 10, 380, 214);

            for (int i = 0; i < apps.Length; i++)
            {
                CheckBox cb = new CheckBox();
                cb.Text = apps[i].Label;
                int col = i / 5, row = i % 5;
                cb.SetBounds(16 + col * 180, 24 + row * 26, 175, 24);
                cb.CheckedChanged += delegate { UpdatePreview(); };
                apps[i].Box = cb;
                g.Controls.Add(cb);
            }

            Label note = new Label();
            note.Text = "Solo se instalan los programas que incluye la edición elegida.";
            note.ForeColor = SystemColors.GrayText;
            note.SetBounds(16, 154, 350, 18);
            g.Controls.Add(note);

            Button bAll = SmallButton("Todos", 16, 176);
            bAll.Click += delegate { SetAll(true); };
            Button bNone = SmallButton("Ninguno", 112, 176);
            bNone.Click += delegate { SetAll(false); };
            Button bDef = SmallButton("Predeterminados", 208, 176);
            bDef.Width = 120;
            bDef.Click += delegate { ApplyDefaults(); };
            g.Controls.AddRange(new Control[] { bAll, bNone, bDef });
        }

        void BuildXmlGroup()
        {
            GroupBox g = Group("Vista previa de la configuración", 388, 232, 380, 178);
            txtXml = new TextBox();
            txtXml.Multiline = true;
            txtXml.ReadOnly = true;
            txtXml.ScrollBars = ScrollBars.Both;
            txtXml.WordWrap = false;
            txtXml.BackColor = SystemColors.Window;
            txtXml.Font = new Font("Consolas", 8.5F);
            txtXml.Dock = DockStyle.Fill;
            g.Controls.Add(txtXml);
        }

        void BuildBottom()
        {
            lblOdt = new Label();
            lblOdt.SetBounds(12, 424, 590, 20);

            lnkManual = new LinkLabel();
            lnkManual.Text = "Descargar manualmente desde Microsoft";
            lnkManual.SetBounds(12, 446, 300, 20);
            lnkManual.Visible = false;
            lnkManual.LinkClicked += delegate
            {
                try { Process.Start(OdtDownloader.PageUrl); } catch { }
            };

            btnOdt = new Button();
            btnOdt.SetBounds(618, 420, 150, 28);
            btnOdt.Click += delegate { StartOdtDownload(); };
            tips.SetToolTip(btnOdt, "Descarga la versión más reciente de Office Deployment Tool (setup.exe).");

            progress = new ProgressBar();
            progress.SetBounds(12, 472, 756, 16);

            lblStatus = new Label();
            lblStatus.SetBounds(12, 492, 756, 20);

            btnSave = new Button();
            btnSave.Text = "Guardar XML…";
            btnSave.SetBounds(442, 518, 110, 32);
            btnSave.Click += delegate { SaveXml(); };

            btnInstall = new Button();
            btnInstall.Text = "Instalar";
            btnInstall.SetBounds(560, 518, 100, 32);
            btnInstall.Font = new Font(Font, FontStyle.Bold);
            btnInstall.Click += delegate { Install(); };

            btnExit = new Button();
            btnExit.Text = "Salir";
            btnExit.SetBounds(668, 518, 100, 32);
            btnExit.Click += delegate { Close(); };

            AcceptButton = btnInstall;
            Controls.AddRange(new Control[] { lblOdt, lnkManual, btnOdt, progress, lblStatus, btnSave, btnInstall, btnExit });
            RefreshOdtState();
        }

        GroupBox Group(string text, int x, int y, int w, int h)
        {
            GroupBox g = new GroupBox();
            g.Text = text;
            g.SetBounds(x, y, w, h);
            Controls.Add(g);
            return g;
        }

        static Button SmallButton(string text, int x, int y)
        {
            Button b = new Button();
            b.Text = text;
            b.SetBounds(x, y, 90, 26);
            return b;
        }

        // ---------------- Logica de seleccion ----------------

        Product SelectedProduct
        {
            get { return lstProducts.SelectedItem as Product; }
        }

        string SelectedLang
        {
            get { return ((LangItem)cmbLang.SelectedItem).Code; }
        }

        string Arch
        {
            get { return rb32.Checked ? "32" : "64"; }
        }

        void FillProducts()
        {
            bool volume = rbVolume.Checked;
            lstProducts.BeginUpdate();
            lstProducts.Items.Clear();
            foreach (Product p in Catalog.Products)
                if (p.Volume == volume) lstProducts.Items.Add(p);
            lstProducts.EndUpdate();
            lstProducts.SelectedIndex = 0;
        }

        void ProductChanged()
        {
            Product p = SelectedProduct;
            if (p == null) return;

            lblProduct.Text = "ID: " + p.Id + (p.Channel != null ? "   ·   Canal: " + p.Channel : "");
            lblSupport.Text = p.Year < 2024
                ? "⚠ Office " + p.Year + " está fuera de soporte desde el 14/10/2025. Considera una versión 2024."
                : "";

            AppItem pub = FindApp("Publisher");
            if (p.Year >= 2024)
            {
                pub.Box.Enabled = false;
                tips.SetToolTip(pub.Box, "Publisher no se incluye en Office 2024.");
            }
            else
            {
                pub.Box.Enabled = true;
                tips.SetToolTip(pub.Box, null);
            }
            ApplyDefaults();
        }

        // Mismos valores que los XML de Retail\ y Volume\
        void ApplyDefaults()
        {
            Product p = SelectedProduct;
            foreach (AppItem a in apps)
            {
                bool on = a.Id == "Word" || a.Id == "Excel" || a.Id == "PowerPoint" || a.Id == "OneDrive"
                    || (a.Id == "Outlook" && p != null && p.Outlook);
                a.Box.Checked = on && a.Box.Enabled;
            }
            UpdatePreview();
        }

        void SetAll(bool value)
        {
            foreach (AppItem a in apps)
                a.Box.Checked = value && a.Box.Enabled;
            UpdatePreview();
        }

        AppItem FindApp(string id)
        {
            foreach (AppItem a in apps)
                if (a.Id == id) return a;
            return null;
        }

        List<string> SelectedApps()
        {
            List<string> list = new List<string>();
            foreach (AppItem a in apps)
                if (a.Box.Checked) list.Add(a.Box.Text);
            return list;
        }

        string BuildXml()
        {
            Product p = SelectedProduct;
            if (p == null) return "";

            XElement product = new XElement("Product", new XAttribute("ID", p.Id),
                new XElement("Language", new XAttribute("ID", SelectedLang)));
            foreach (AppItem a in apps)
                if (!a.Box.Checked)
                    product.Add(new XElement("ExcludeApp", new XAttribute("ID", a.Id)));

            XElement add = new XElement("Add", new XAttribute("OfficeClientEdition", Arch));
            if (p.Channel != null) add.Add(new XAttribute("Channel", p.Channel));
            add.Add(product);

            XElement root = new XElement("Configuration",
                add,
                new XElement("Updates", new XAttribute("Enabled", "TRUE")),
                new XElement("Display", new XAttribute("Level", "Full"), new XAttribute("AcceptEULA", "TRUE")));

            XmlWriterSettings s = new XmlWriterSettings();
            s.OmitXmlDeclaration = true;
            s.Indent = true;
            s.IndentChars = "  ";
            StringBuilder sb = new StringBuilder();
            using (XmlWriter w = XmlWriter.Create(sb, s))
                root.Save(w);
            return sb.ToString() + "\r\n";
        }

        void UpdatePreview()
        {
            if (txtXml == null) return;
            txtXml.Text = BuildXml();
            UpdateButtons();
        }

        string DefaultXmlName()
        {
            return SelectedProduct.Id + "-x" + Arch + "-" + SelectedLang + ".xml";
        }

        void WriteXml(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, BuildXml(), new UTF8Encoding(false));
        }

        bool ValidateSelection()
        {
            if (SelectedApps().Count == 0)
            {
                MessageBox.Show(this, "Selecciona al menos un programa.", Text,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        // ---------------- Acciones ----------------

        void SaveXml()
        {
            if (!ValidateSelection()) return;
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                Directory.CreateDirectory(outDir);
                dlg.InitialDirectory = outDir;
                dlg.FileName = DefaultXmlName();
                dlg.Filter = "Configuración XML (*.xml)|*.xml";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    WriteXml(dlg.FileName);
                    lblStatus.Text = "XML guardado en " + dlg.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No se pudo guardar el XML:\n" + ex.Message, Text,
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        void Install()
        {
            if (Busy || !ValidateSelection()) return;
            if (!File.Exists(setupPath))
            {
                RefreshOdtState();
                MessageBox.Show(this, "Falta setup.exe. Usa el botón «Descargar ODT».", Text,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Product p = SelectedProduct;
            string msg = "Se instalará:\n\n" +
                p.Name + " (" + p.Id + ")\n" +
                Arch + " bits · " + SelectedLang + "\n\n" +
                "Programas: " + string.Join(", ", SelectedApps().ToArray()) + "\n\n" +
                "Office se descargará desde la CDN de Microsoft y puede tardar varios minutos.\n¿Continuar?";
            if (MessageBox.Show(this, msg, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            string xml = Path.Combine(outDir, DefaultXmlName());
            try { WriteXml(xml); }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo guardar el XML:\n" + ex.Message, Text,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SetBusy(true, "Instalando " + p.Name + "… sigue el progreso en la ventana de Office.");
            installWorker.RunWorkerAsync(xml);
        }

        void InstallWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            ProcessStartInfo psi = new ProcessStartInfo(setupPath, "/configure \"" + (string)e.Argument + "\"");
            psi.WorkingDirectory = appDir;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            using (Process proc = Process.Start(psi))
            {
                proc.WaitForExit();
                e.Result = proc.ExitCode;
            }
        }

        void InstallWorker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            SetBusy(false, "");
            if (e.Error != null)
            {
                lblStatus.Text = "Error al ejecutar setup.exe.";
                MessageBox.Show(this, "No se pudo ejecutar setup.exe:\n" + e.Error.Message, Text,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int code = (int)e.Result;
            if (code == 0)
            {
                lblStatus.Text = "Instalación terminada.";
                MessageBox.Show(this, "Instalación terminada.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                lblStatus.Text = "setup.exe terminó con código " + code + ".";
                MessageBox.Show(this, "setup.exe terminó con código " + code + ".\n\n" +
                    "Revisa los registros de ODT en:\n" + Path.GetTempPath(), Text,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        void StartOdtDownload()
        {
            if (Busy) return;
            odtFailed = false;
            lnkManual.Visible = false;
            SetBusy(true, "Preparando descarga de Office Deployment Tool…");
            odtWorker.RunWorkerAsync();
        }

        void OdtWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker w = (BackgroundWorker)sender;
            OdtDownloader.Run(setupPath, delegate(int pct, string text) { w.ReportProgress(pct, text); });
        }

        void OdtWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage >= 0)
            {
                progress.Style = ProgressBarStyle.Continuous;
                progress.Value = Math.Min(100, e.ProgressPercentage);
            }
            else
            {
                progress.Style = ProgressBarStyle.Marquee;
            }
            lblStatus.Text = (string)e.UserState;
        }

        void OdtWorker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            SetBusy(false, "");
            if (e.Error != null)
            {
                odtFailed = true;
                lblStatus.Text = "No se pudo descargar Office Deployment Tool: " + e.Error.Message.Replace("\n", " ");
                RefreshOdtState();
                return;
            }
            lblStatus.Text = "Office Deployment Tool descargado y verificado.";
            RefreshOdtState();
        }

        // ---------------- Estado ----------------

        bool Busy
        {
            get { return odtWorker.IsBusy || installWorker.IsBusy; }
        }

        void SetBusy(bool busy, string status)
        {
            foreach (Control c in Controls)
                if (c is GroupBox) c.Enabled = !busy;
            progress.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
            progress.MarqueeAnimationSpeed = busy ? 30 : 0;
            if (!busy) progress.Value = 0;
            if (status.Length > 0) lblStatus.Text = status;
            btnOdt.Enabled = !busy;
            btnSave.Enabled = !busy;
            btnInstall.Enabled = !busy && File.Exists(setupPath);
        }

        void RefreshOdtState()
        {
            bool exists = File.Exists(setupPath);
            if (exists)
            {
                string ver = "";
                try { ver = FileVersionInfo.GetVersionInfo(setupPath).FileVersion; } catch { }
                lblOdt.Text = "✔ Office Deployment Tool listo" + (string.IsNullOrEmpty(ver) ? "" : " (setup.exe " + ver + ")");
                lblOdt.ForeColor = Color.FromArgb(16, 124, 16);
                btnOdt.Text = "Actualizar ODT";
            }
            else
            {
                lblOdt.Text = odtFailed
                    ? "✖ Falta setup.exe y la descarga automática falló."
                    : "✖ Falta setup.exe (Office Deployment Tool).";
                lblOdt.ForeColor = Color.FromArgb(196, 43, 28);
                btnOdt.Text = odtFailed ? "Reintentar descarga" : "Descargar ODT";
            }
            lnkManual.Visible = odtFailed;
            UpdateButtons();
        }

        void UpdateButtons()
        {
            if (btnInstall == null || odtWorker == null) return;
            if (Busy) return;
            bool hasApps = SelectedApps().Count > 0;
            btnInstall.Enabled = hasApps && File.Exists(setupPath);
            btnSave.Enabled = hasApps;
        }

        void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (installWorker.IsBusy)
            {
                if (MessageBox.Show(this, "La instalación de Office sigue en curso y continuará aunque cierres esta ventana.\n¿Salir?",
                        Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                    e.Cancel = true;
            }
            else if (odtWorker.IsBusy)
            {
                if (MessageBox.Show(this, "Se está descargando Office Deployment Tool. ¿Salir de todos modos?",
                        Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                    e.Cancel = true;
            }
        }
    }
}
