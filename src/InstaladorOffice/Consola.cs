using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InstaladorOffice;

/// <summary>Un comando de consola (cscript ospp.vbs, slmgr.vbs, …) con su descripción.</summary>
sealed class Comando
{
    public string Descripcion { get; }
    public string Archivo { get; }
    public string Argumentos { get; }

    public Comando(string descripcion, string archivo, string argumentos)
    {
        Descripcion = descripcion; Archivo = archivo; Argumentos = argumentos;
    }

    /// <summary>cscript //nologo "script.vbs" argumentos</summary>
    public static Comando Script(string descripcion, string rutaVbs, string argumentos) =>
        new(descripcion, System.IO.Path.Combine(Environment.SystemDirectory, "cscript.exe"),
            $"//nologo \"{rutaVbs}\" {argumentos}");

    public override string ToString() => $"{System.IO.Path.GetFileName(Archivo)} {Argumentos}";
}

static class Ejecutor
{
    // cscript escribe en la página de códigos OEM de la consola (850 en español).
    static readonly Encoding CodificacionOem =
        Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);

    /// <summary>Ejecuta un proceso oculto y entrega cada línea de salida. Devuelve el código de salida.</summary>
    public static Task<int> EjecutarAsync(Comando cmd, Action<string> linea)
    {
        var tcs = new TaskCompletionSource<int>();
        var p = new Process
        {
            StartInfo = new ProcessStartInfo(cmd.Archivo, cmd.Argumentos)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = CodificacionOem,
                StandardErrorEncoding = CodificacionOem,
                WorkingDirectory = Environment.SystemDirectory,
            },
            EnableRaisingEvents = true,
        };
        p.OutputDataReceived += (_, e) => { if (e.Data != null) linea(e.Data); };
        p.ErrorDataReceived += (_, e) => { if (e.Data != null) linea(e.Data); };
        p.Exited += (_, _) =>
        {
            p.WaitForExit(); // espera a que termine de llegar la salida asíncrona
            tcs.TrySetResult(p.ExitCode);
            p.Dispose();
        };
        try
        {
            p.Start();
            p.BeginOutputReadLine();
            p.BeginErrorReadLine();
        }
        catch (Exception ex)
        {
            p.Dispose();
            tcs.TrySetException(ex);
        }
        return tcs.Task;
    }
}

/// <summary>Consola de salida con colores, para mostrar lo que hacen ospp/slmgr.</summary>
sealed class PanelConsola : RichTextBox
{
    public static readonly Color ColorComando = Color.FromArgb(97, 175, 239);
    public static readonly Color ColorOk = Color.FromArgb(152, 195, 121);
    public static readonly Color ColorError = Color.FromArgb(224, 108, 117);
    public static readonly Color ColorAviso = Color.FromArgb(229, 192, 123);

    public PanelConsola()
    {
        ReadOnly = true;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.Gainsboro;
        Font = new Font("Consolas", 9F);
        BorderStyle = BorderStyle.None;
        DetectUrls = false;
        WordWrap = false;
        ScrollBars = RichTextBoxScrollBars.Both;

        var menu = new ContextMenuStrip();
        menu.Items.Add("Copiar todo", null, (_, _) => { if (TextLength > 0) Clipboard.SetText(Text); });
        menu.Items.Add("Limpiar", null, (_, _) => Clear());
        ContextMenuStrip = menu;
    }

    public void Escribir(string texto, Color? color = null)
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => Escribir(texto, color)));
            return;
        }
        SelectionStart = TextLength;
        SelectionLength = 0;
        SelectionColor = color ?? ForeColor;
        AppendText(texto + "\n");
        SelectionColor = ForeColor;
        ScrollToCaret();
    }
}
