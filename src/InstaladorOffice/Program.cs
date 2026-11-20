using System;
using System.Drawing;
using System.Windows.Forms;

namespace InstaladorOffice;

static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new VentanaPrincipal());
    }
}

sealed class VentanaPrincipal : Form
{
    readonly PaginaInstalar instalar;
    readonly PaginaActivacion activarOffice;
    readonly PaginaActivacion activarWindows;

    public VentanaPrincipal()
    {
        Text = "Instalador y activador de Office";
        Font = new Font("Segoe UI", 9F);
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(820, 640);
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

        SuspendLayout();
        instalar = new PaginaInstalar { Location = new Point(4, 4) };
        activarOffice = new PaginaActivacion(new ActivadorOffice()) { Location = new Point(4, 4) };
        activarWindows = new PaginaActivacion(new ActivadorWindows()) { Location = new Point(4, 4) };

        var pestañas = new TabControl { Dock = DockStyle.Fill, Padding = new Point(14, 5) };
        pestañas.TabPages.Add(Pestaña("Instalar Office", instalar));
        pestañas.TabPages.Add(Pestaña("Activar Office", activarOffice));
        pestañas.TabPages.Add(Pestaña("Activar Windows", activarWindows));
        Controls.Add(pestañas);
        ResumeLayout(false);

        FormClosing += AlCerrar;
    }

    static TabPage Pestaña(string titulo, Control contenido)
    {
        var p = new TabPage(titulo) { UseVisualStyleBackColor = true };
        p.Controls.Add(contenido);
        return p;
    }

    void AlCerrar(object sender, FormClosingEventArgs e)
    {
        string mensaje = null;
        if (instalar.Instalando)
            mensaje = "La instalación de Office sigue en curso y continuará aunque cierres esta ventana.\n¿Salir?";
        else if (instalar.Ocupado)
            mensaje = "Se está descargando Office Deployment Tool. ¿Salir de todos modos?";
        else if (activarOffice.Ocupado || activarWindows.Ocupado)
            mensaje = "Hay un comando de activación en curso. ¿Salir de todos modos?";

        if (mensaje != null &&
            MessageBox.Show(this, mensaje, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            e.Cancel = true;
    }
}
