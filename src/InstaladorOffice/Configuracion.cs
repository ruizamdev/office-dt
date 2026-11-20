namespace InstaladorOffice;

// ============================================================================
//  CONFIGURACIÓN QUE SE COMPILA DENTRO DEL EXE
//
//  Edita este archivo y vuelve a compilar (build.sh o build.cmd).
//
//  ⚠ Todo lo que pongas aquí queda incrustado en InstaladorOffice.exe y
//    cualquiera que tenga el exe puede extraerlo. Las GVLK son públicas, pero
//    los nombres de servidores de clientes y las claves MAK no lo son:
//    no subas este archivo con datos reales a un repositorio público y no
//    distribuyas un exe con MAK fuera de tu equipo de trabajo.
// ============================================================================
static class Configuracion
{
    // ------------------------------------------------------------------------
    //  1. SERVIDORES KMS
    //  Aparecen en el menú desplegable «Servidor KMS» de las pestañas de
    //  activación. Formato: nombre visible, host (DNS o IP), puerto (1688 por
    //  defecto). También se puede escribir host:puerto a mano en el programa.
    // ------------------------------------------------------------------------
    public static readonly ServidorKms[] ServidoresKms =
    {
        new("Cliente 1", "kms.digiboy.ir", 1688),
        new("Cliente 2", "kms.sixyin.com", 1688),
        new("Cliente 3", "kms.sdit163.com", 1688),
        new("Cliente 4", "kms.catqu.com", 1688),
        new("Cliente 5", "kms.ghpym.com", 1688),
        new("Cliente 6", "kms8.MSGuides.com", 1688),
        new("Cliente 7", "kms9.MSGuides.com", 1688),
        new("Cliente 8", "kms.loli.best", 1688),
        new("Cliente 9", "kms.000606.xyz", 1688),
        new("Cliente 10", "kms.cgtsoft.com", 1688),
        new("Cliente 11", "kms.jm33.me", 1688),
        new("Cliente 12", "kms.kmzs123.cn", 1688),
        new("Cliente 13", "xincheng213618.cn", 1688),
        new("Cliente 14", "kms.myds.cloud", 1688),
        new("Cliente 15", "kms.mc06.net", 1688),
        new("Cliente 16", "kms.moeyuuko.top", 1688),
        new("Cliente 17", "kms.xykz.f3322.org", 1688),
        new("Cliente 18", "kms.ddns.net", 1688),
        // new("Cliente 3", "kms.cliente-c.com", 1688),
    };

    // ------------------------------------------------------------------------
    //  2. CLAVES DE OFFICE
    //  Formato: nombre visible, clave, ProductReleaseId de Click-to-Run.
    //  El ProductReleaseId sirve para preseleccionar la clave correcta cuando
    //  se detecta la edición instalada (déjalo en null si no aplica).
    //
    //  Las GVLK de abajo son las claves públicas de cliente KMS publicadas por
    //  Microsoft: https://learn.microsoft.com/office/volume-license-activation/gvlks
    //  Agrega tus claves MAK al final de la lista.
    // ------------------------------------------------------------------------
    public static readonly ClaveProducto[] ClavesOffice =
    {
        new("Office LTSC Profesional Plus 2024 (GVLK)", "XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB", "ProPlus2024Volume"),
        new("Office LTSC Standard 2024 (GVLK)",         "V28N4-JG22K-W66P8-VTMGK-H6HGR", "Standard2024Volume"),
        new("Project Professional 2024 (GVLK)",         "FQQ23-N4YCY-73HQ3-FM9WC-76HF4", "ProjectPro2024Volume"),
        new("Project Standard 2024 (GVLK)",             "PD3TT-NTHQQ-VC7CY-MFXK3-G87F8", "ProjectStd2024Volume"),
        new("Visio LTSC Professional 2024 (GVLK)",      "B7TN8-FJ8V3-7QYCP-HQPMV-YY89G", "VisioPro2024Volume"),
        new("Visio LTSC Standard 2024 (GVLK)",          "JMMVY-XFNQC-KK4HK-9H7R3-WQQTV", "VisioStd2024Volume"),

        new("Office LTSC Profesional Plus 2021 (GVLK)", "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH", "ProPlus2021Volume"),
        new("Office LTSC Standard 2021 (GVLK)",         "KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3", "Standard2021Volume"),
        new("Project Professional 2021 (GVLK)",         "FTNWT-C6WBT-8HMGF-K9PRX-QV9H8", "ProjectPro2021Volume"),
        new("Project Standard 2021 (GVLK)",             "J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T", "ProjectStd2021Volume"),
        new("Visio LTSC Professional 2021 (GVLK)",      "KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4", "VisioPro2021Volume"),
        new("Visio LTSC Standard 2021 (GVLK)",          "MJVNY-BYWPY-CWV6J-2RKRT-4M8QG", "VisioStd2021Volume"),

        new("Office Profesional Plus 2019 (GVLK)",      "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP", "ProPlus2019Volume"),
        new("Office Standard 2019 (GVLK)",              "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK", "Standard2019Volume"),
        new("Project Professional 2019 (GVLK)",         "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B", "ProjectPro2019Volume"),
        new("Project Standard 2019 (GVLK)",             "C4F7P-NCP8C-6CQPT-MQHV9-JXD2M", "ProjectStd2019Volume"),
        new("Visio Professional 2019 (GVLK)",           "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB", "VisioPro2019Volume"),
        new("Visio Standard 2019 (GVLK)",               "7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2", "VisioStd2019Volume"),

        // --- Tus claves MAK u otras claves por volumen ---
        // new("Cliente A - Office 2024 MAK", "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX", "ProPlus2024Volume"),
    };

    // ------------------------------------------------------------------------
    //  3. CLAVES DE WINDOWS
    //  Formato: nombre visible, clave, EditionID de Windows y, opcionalmente,
    //  el número de compilación (build) para distinguir versiones de Server.
    //  EditionID y build se usan para preseleccionar la clave al detectar.
    //
    //  GVLK públicas de Microsoft:
    //  https://learn.microsoft.com/windows-server/get-started/kms-client-activation-keys
    //  Agrega tus claves MAK al final de la lista.
    // ------------------------------------------------------------------------
    public static readonly ClaveProducto[] ClavesWindows =
    {
        new("Windows 10/11 Pro (GVLK)",                      "W269N-WFGWX-YVC9B-4J6C9-T83GX", "Professional"),
        new("Windows 10/11 Pro N (GVLK)",                    "MH37W-N47XK-V7XM9-C7227-GCQG9", "ProfessionalN"),
        new("Windows 10/11 Pro for Workstations (GVLK)",     "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J", "ProfessionalWorkstation"),
        new("Windows 10/11 Pro for Workstations N (GVLK)",   "9FNHH-K3HBT-3W4TD-6383H-6XYWF", "ProfessionalWorkstationN"),
        new("Windows 10/11 Pro Education (GVLK)",            "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y", "ProfessionalEducation"),
        new("Windows 10/11 Pro Education N (GVLK)",          "YVWGF-BXNMC-HTQYQ-CPQ99-66QFC", "ProfessionalEducationN"),
        new("Windows 10/11 Education (GVLK)",                "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2", "Education"),
        new("Windows 10/11 Education N (GVLK)",              "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ", "EducationN"),
        new("Windows 10/11 Enterprise (GVLK)",               "NPPR9-FWDCX-D2C8J-H872K-2YT43", "Enterprise"),
        new("Windows 10/11 Enterprise N (GVLK)",             "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4", "EnterpriseN"),
        new("Windows 10/11 Enterprise G (GVLK)",             "YYVX9-NTFWV-6MDM3-9PT4T-4M68B", "EnterpriseG"),
        new("Windows 10/11 Enterprise G N (GVLK)",           "44RPN-FTY23-9VTTB-MP9BX-T84FV", "EnterpriseGN"),
        new("Windows Enterprise LTSC 2019/2021/2024 (GVLK)", "M7XTQ-FN8P6-TTKYV-9D4CC-J462D", "EnterpriseS"),
        new("Windows Enterprise N LTSC 2019/2021/2024 (GVLK)", "92NFX-8DJQP-P6BBQ-THF9C-7CG2H", "EnterpriseSN"),
        new("Windows IoT Enterprise LTSC 2021/2024 (GVLK)",  "KBN8V-HFGQ4-MGXVD-347P6-PDQGT", "IoTEnterpriseS"),

        new("Windows Server 2025 Standard (GVLK)",           "TVRH6-WHNXV-R9WG3-9XRFY-MY832", "ServerStandard",   26100),
        new("Windows Server 2025 Datacenter (GVLK)",         "D764K-2NDRG-47T6Q-P8T8W-YP6DF", "ServerDatacenter", 26100),
        new("Windows Server 2022 Standard (GVLK)",           "VDYBN-27WPP-V4HQT-9VMD4-VMK7H", "ServerStandard",   20348),
        new("Windows Server 2022 Datacenter (GVLK)",         "WX4NM-KYWYW-QJJR4-XV3QB-6VM33", "ServerDatacenter", 20348),
        new("Windows Server 2019 Standard (GVLK)",           "N69G4-B89J2-4G8F4-WWYCC-J464C", "ServerStandard",   17763),
        new("Windows Server 2019 Datacenter (GVLK)",         "WMDGN-G9PQG-XVVXX-R3X43-63DFG", "ServerDatacenter", 17763),
        new("Windows Server 2016 Standard (GVLK)",           "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY", "ServerStandard",   14393),
        new("Windows Server 2016 Datacenter (GVLK)",         "CB7KF-BWN84-R7R2Y-793K2-8XDDG", "ServerDatacenter", 14393),

        // --- Tus claves MAK u otras claves por volumen ---
        // new("Cliente A - Windows 11 Pro MAK", "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX", "Professional"),
    };
}
