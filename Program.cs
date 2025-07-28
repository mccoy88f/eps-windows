using System;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit;
using MimeKit;
using System.Linq;
using System.Threading;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Microsoft.Win32;
using System.Text.Json;
using System.Collections.Generic;

namespace EmailPrintService
{
    // Informazioni applicazione
    public static class AppInfo
    {
        public const string Name = "Email Print Service";
        public const string Version = "1.0.0";
        public const string Author = "Antonello Migliorelli";
        public const string GitHubUsername = "mccoy88f";
        public const string GitHubUrl = "https://github.com/mccoy88f/eps-windows";
        public const string License = "MIT License";
        public const string Copyright = "© 2025 Antonello Migliorelli";
        public const string Description = "Professional Windows service for automatic PDF printing from email attachments";
    }

    // Classe per informazioni sui file PDF
    public class PdfFileInfo
    {
        public string FileName { get; set; } = "";
        public long FileSizeBytes { get; set; }
        public MimePart Attachment { get; set; }
    }

    // Classe per le impostazioni usando System.Text.Json
    public class AppSettings
    {
        public string EmailServer { get; set; } = "imap.gmail.com";
        public int EmailPort { get; set; } = 993;
        public string EmailUsername { get; set; } = "";
        public string EmailPassword { get; set; } = "";
        public string PrinterName { get; set; } = "";
        public bool SendConfirmation { get; set; } = false;
        public bool DeleteAfterPrint { get; set; } = false;
        public bool AutoStart { get; set; } = false;
        public int CheckInterval { get; set; } = 10;

        private static readonly string _settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailPrintService",
            "settings.json"
        );

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var json = File.ReadAllText(_settingsPath);
                    return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch
            {
                // Se c'è un errore, restituisce impostazioni di default
            }
            return new AppSettings();
        }

        public void Save()
        {
            try
            {
                var directory = Path.GetDirectoryName(_settingsPath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_settingsPath, json);
            }
            catch
            {
                // Ignora errori di salvataggio
            }
        }
    }

    // Enum per le modalità di stampa
    public enum PrintConfirmationMode
    {
        Automatic = 0,          // Stampa diretta automatica
        TimedConfirmation = 1,  // Popup con timer
        ManualConfirmation = 2  // Popup obbligatorio
    }

    // Enum per i metodi di stampa
    public enum PrintMethod
    {
        NetPrintDocument = 0,    // .NET PrintDocument (nativo)
        SumatraPDF = 1,         // SumatraPDF (esterno)
        WindowsShell = 2,       // Windows Shell (associazione file)
        Auto = 3                // Automatico (prova tutti)
    }

    // Classe per salvare le impostazioni di stampa
    public class PrintSettings
    {
        public string PrinterName { get; set; } = "";
        public bool Duplex { get; set; } = true; // Fronte/retro
        public string PaperSize { get; set; } = "A4";
        public bool Landscape { get; set; } = false;
        public int Copies { get; set; } = 1;
        public string Quality { get; set; } = "Normal";
        public bool ColorPrinting { get; set; } = false;
        public bool FitToPage { get; set; } = true;
        public PrintMethod PreferredMethod { get; set; } = PrintMethod.Auto;
        public PrintConfirmationMode ConfirmationMode { get; set; } = PrintConfirmationMode.Automatic;
        public int ConfirmationTimeout { get; set; } = 15; // Secondi per il timeout
        public string PrinterSettingsData { get; set; } = ""; // Dati serializzati completi

        private static readonly string _printSettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailPrintService",
            "print_settings.json"
        );

        public static PrintSettings Load()
        {
            try
            {
                if (File.Exists(_printSettingsPath))
                {
                    var json = File.ReadAllText(_printSettingsPath);
                    return JsonSerializer.Deserialize<PrintSettings>(json) ?? new PrintSettings();
                }
            }
            catch { }
            return new PrintSettings();
        }

        public void Save()
        {
            try
            {
                var directory = Path.GetDirectoryName(_printSettingsPath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_printSettingsPath, json);
            }
            catch { }
        }
    }

    // Dialog placeholder classes (da implementare)
    public partial class PdfDetailDialog : Form
    {
        public List<PdfFileInfo> SelectedFiles { get; private set; } = new List<PdfFileInfo>();

        public PdfDetailDialog(List<PdfFileInfo> fileInfos, string sender, string senderEmail, string subject, bool hasTimeout, int timeoutSeconds)
        {
            InitializeComponent();
            // TODO: Implementare il dialog
            SelectedFiles = fileInfos; // Per ora seleziona tutti
        }

        private void InitializeComponent()
        {
            this.Text = "Selezione File PDF";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            
            var btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new Point(300, 320) };
            var btnCancel = new Button { Text = "Annulla", DialogResult = DialogResult.Cancel, Location = new Point(380, 320) };
            
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);
        }
    }

    public partial class AboutDialog : Form
    {
        public AboutDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = $"About {AppInfo.Name}";
            this.Size = new Size(400, 300);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lblInfo = new Label
            {
                Text = $@"{AppInfo.Name} v{AppInfo.Version}
{AppInfo.Description}

© {AppInfo.Copyright}
License: {AppInfo.License}

GitHub: {AppInfo.GitHubUrl}",
                Location = new Point(20, 20),
                Size = new Size(350, 200),
                AutoSize = false
            };

            var btnClose = new Button
            {
                Text = "Chiudi",
                DialogResult = DialogResult.OK,
                Location = new Point(160, 230),
                Size = new Size(80, 25)
            };

            this.Controls.Add(lblInfo);
            this.Controls.Add(btnClose);
        }
    }

    public partial class MainForm : Form
    {
        private EmailPrintService _service;
        private NotifyIcon _trayIcon;
        private bool _isServiceRunning;
        private AppSettings _settings;
        private PrintSettings _printSettings;

        // Controlli UI
        private TextBox txtEmailServer, txtEmailUsername, txtEmailPassword, txtLog;
        private NumericUpDown txtEmailPort, txtInterval, txtTimeout;
        private ComboBox cmbPrinter;
        private CheckBox chkSendConfirmation, chkDeleteAfterPrint, chkAutoStart;
        private Button btnSave, btnStart, btnStop, btnConfigurePrint;
        private Label lblStatus, lblPrintInfo, lblMethodInfo, lblConfirmInfo;
        private RadioButton rbPrintAuto, rbPrintNet, rbPrintSumatra, rbPrintShell;
        private RadioButton rbConfirmAuto, rbConfirmTimed, rbConfirmManual;

        public MainForm()
        {
            InitializeComponent();
            InitializeTrayIcon();
            _service = new EmailPrintService();
            _settings = AppSettings.Load();
            _printSettings = PrintSettings.Load();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.Text = $"{AppInfo.Name} v{AppInfo.Version} - by {AppInfo.Author}";
            this.Size = new Size(500, 900);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Minimized;

            CreateControls();
        }

        private void CreateControls()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 35,
                Padding = new Padding(10)
            };

            // Impostazioni Email
            panel.Controls.Add(new Label { Text = "=== IMPOSTAZIONI EMAIL ===", Font = new Font("Arial", 10, FontStyle.Bold) }, 0, 0);
            panel.SetColumnSpan(panel.Controls[panel.Controls.Count - 1], 2);

            panel.Controls.Add(new Label { Text = "Server IMAP:" }, 0, 1);
            txtEmailServer = new TextBox { Width = 200 };
            panel.Controls.Add(txtEmailServer, 1, 1);

            panel.Controls.Add(new Label { Text = "Porta:" }, 0, 2);
            txtEmailPort = new NumericUpDown { Minimum = 1, Maximum = 65535, Value = 993, Width = 200 };
            panel.Controls.Add(txtEmailPort, 1, 2);

            panel.Controls.Add(new Label { Text = "Username:" }, 0, 3);
            txtEmailUsername = new TextBox { Width = 200 };
            panel.Controls.Add(txtEmailUsername, 1, 3);

            panel.Controls.Add(new Label { Text = "Password:" }, 0, 4);
            txtEmailPassword = new TextBox { UseSystemPasswordChar = true, Width = 200 };
            panel.Controls.Add(txtEmailPassword, 1, 4);

            // Impostazioni Stampante
            panel.Controls.Add(new Label { Text = "=== IMPOSTAZIONI STAMPANTE ===", Font = new Font("Arial", 10, FontStyle.Bold) }, 0, 6);
            panel.SetColumnSpan(panel.Controls[panel.Controls.Count - 1], 2);

            panel.Controls.Add(new Label { Text = "Stampante:" }, 0, 7);
            cmbPrinter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
            LoadPrinters();
            panel.Controls.Add(cmbPrinter, 1, 7);

            // Pulsante configurazione stampa
            btnConfigurePrint = new Button { Text = "Configura Stampa", Width = 120 };
            btnConfigurePrint.Click += BtnConfigurePrint_Click;
            panel.Controls.Add(btnConfigurePrint, 1, 8);

            // Info stampa corrente
            lblPrintInfo = new Label { Text = "Clicca 'Configura Stampa' per impostare formato, fronte/retro, etc.", ForeColor = Color.Blue, Width = 400, AutoSize = true };
            panel.Controls.Add(lblPrintInfo, 0, 9);
            panel.SetColumnSpan(lblPrintInfo, 2);

            // Metodo di stampa
            panel.Controls.Add(new Label { Text = "=== METODO STAMPA ===", Font = new Font("Arial", 10, FontStyle.Bold) }, 0, 10);
            panel.SetColumnSpan(panel.Controls[panel.Controls.Count - 1], 2);

            var printMethodGroup = new GroupBox { Text = "Scegli metodo di stampa:", Width = 400, Height = 120 };
            
            rbPrintAuto = new RadioButton { Text = "🤖 Automatico (prova tutti i metodi)", Location = new Point(10, 20), Width = 350, Checked = true };
            rbPrintAuto.CheckedChanged += (s, e) => { if (rbPrintAuto.Checked) ShowMethodInfo("Prova prima .NET, poi SumatraPDF, infine Windows Shell"); };
            printMethodGroup.Controls.Add(rbPrintAuto);

            rbPrintNet = new RadioButton { Text = "🏗️ .NET PrintDocument (nativo, zero dipendenze)", Location = new Point(10, 40), Width = 350 };
            rbPrintNet.CheckedChanged += (s, e) => { if (rbPrintNet.Checked) ShowMethodInfo("Usa solo .NET nativo - più pulito ma limitato per PDF complessi"); };
            printMethodGroup.Controls.Add(rbPrintNet);

            rbPrintSumatra = new RadioButton { Text = "⚡ SumatraPDF (richiede SumatraPDF.exe nella cartella app)", Location = new Point(10, 60), Width = 350 };
            rbPrintSumatra.CheckedChanged += (s, e) => { if (rbPrintSumatra.Checked) ShowMethodInfo("Usa SumatraPDF.exe (devi copiarlo manualmente nella cartella dell'app)"); };
            printMethodGroup.Controls.Add(rbPrintSumatra);

            rbPrintShell = new RadioButton { Text = "🪟 Windows Shell (usa app predefinita)", Location = new Point(10, 80), Width = 350 };
            rbPrintShell.CheckedChanged += (s, e) => { if (rbPrintShell.Checked) ShowMethodInfo("Usa l'app predefinita di Windows - può aprire finestre"); };
            printMethodGroup.Controls.Add(rbPrintShell);

            panel.Controls.Add(printMethodGroup, 0, 11);
            panel.SetColumnSpan(printMethodGroup, 2);

            // Info metodo selezionato
            lblMethodInfo = new Label { Text = "Metodo automatico: prova .NET, SumatraPDF, Windows Shell", ForeColor = Color.DarkBlue, Width = 400, Height = 30, AutoSize = false };
            panel.Controls.Add(lblMethodInfo, 0, 12);
            panel.SetColumnSpan(lblMethodInfo, 2);

            // Modalità conferma stampa
            panel.Controls.Add(new Label { Text = "=== MODALITÀ STAMPA ===", Font = new Font("Arial", 10, FontStyle.Bold) }, 0, 13);
            panel.SetColumnSpan(panel.Controls[panel.Controls.Count - 1], 2);

            var confirmGroup = new GroupBox { Text = "Conferma prima di stampare:", Width = 400, Height = 120 };
            
            rbConfirmAuto = new RadioButton { Text = "🚀 Stampa automatica (nessuna conferma)", Location = new Point(10, 20), Width = 350, Checked = true };
            rbConfirmAuto.CheckedChanged += (s, e) => { if (rbConfirmAuto.Checked) { ShowConfirmInfo("Stampa automaticamente senza chiedere conferma"); txtTimeout.Enabled = false; } };
            confirmGroup.Controls.Add(rbConfirmAuto);

            rbConfirmTimed = new RadioButton { Text = "⏱️ Popup con timer (stampa automatica se non si risponde)", Location = new Point(10, 40), Width = 350 };
            rbConfirmTimed.CheckedChanged += (s, e) => { if (rbConfirmTimed.Checked) { ShowConfirmInfo("Mostra popup, se non rispondi entro X secondi stampa automaticamente"); txtTimeout.Enabled = true; } };
            confirmGroup.Controls.Add(rbConfirmTimed);

            rbConfirmManual = new RadioButton { Text = "✋ Conferma obbligatoria (attende sempre risposta utente)", Location = new Point(10, 60), Width = 350 };
            rbConfirmManual.CheckedChanged += (s, e) => { if (rbConfirmManual.Checked) { ShowConfirmInfo("Mostra sempre popup e attende conferma o rifiuto dell'utente"); txtTimeout.Enabled = false; } };
            confirmGroup.Controls.Add(rbConfirmManual);

            // Timeout per modalità timer
            var lblTimeout = new Label { Text = "Timeout (sec):", Location = new Point(10, 85), Width = 80 };
            confirmGroup.Controls.Add(lblTimeout);
            
            txtTimeout = new NumericUpDown { Location = new Point(95, 83), Width = 60, Minimum = 5, Maximum = 300, Value = 15, Enabled = false };
            confirmGroup.Controls.Add(txtTimeout);

            panel.Controls.Add(confirmGroup, 0, 14);
            panel.SetColumnSpan(confirmGroup, 2);

            // Info conferma selezionata
            lblConfirmInfo = new Label { Text = "Stampa automaticamente senza chiedere conferma", ForeColor = Color.DarkGreen, Width = 400, Height = 30, AutoSize = false };
            panel.Controls.Add(lblConfirmInfo, 0, 15);
            panel.SetColumnSpan(lblConfirmInfo, 2);

            // Impostazioni Comportamento
            panel.Controls.Add(new Label { Text = "=== COMPORTAMENTO ===", Font = new Font("Arial", 10, FontStyle.Bold) }, 0, 16);
            panel.SetColumnSpan(panel.Controls[panel.Controls.Count - 1], 2);

            chkSendConfirmation = new CheckBox { Text = "Invia email di conferma stampa", Width = 300 };
            panel.Controls.Add(chkSendConfirmation, 0, 17);
            panel.SetColumnSpan(chkSendConfirmation, 2);

            chkDeleteAfterPrint = new CheckBox { Text = "Elimina email dopo stampa", Width = 300 };
            panel.Controls.Add(chkDeleteAfterPrint, 0, 18);
            panel.SetColumnSpan(chkDeleteAfterPrint, 2);

            chkAutoStart = new CheckBox { Text = "Avvia automaticamente con Windows", Width = 300 };
            panel.Controls.Add(chkAutoStart, 0, 19);
            panel.SetColumnSpan(chkAutoStart, 2);

            panel.Controls.Add(new Label { Text = "Intervallo controllo (secondi):" }, 0, 20);
            txtInterval = new NumericUpDown { Minimum = 5, Maximum = 300, Value = 10, Width = 200 };
            panel.Controls.Add(txtInterval, 1, 20);

            // Pulsanti
            var buttonPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Width = 400, Height = 40 };
            
            btnSave = new Button { Text = "Salva Impostazioni", Width = 120 };
            btnSave.Click += BtnSave_Click;
            buttonPanel.Controls.Add(btnSave);

            btnStart = new Button { Text = "Avvia Servizio", Width = 120 };
            btnStart.Click += BtnStart_Click;
            buttonPanel.Controls.Add(btnStart);

            btnStop = new Button { Text = "Ferma Servizio", Width = 120, Enabled = false };
            btnStop.Click += BtnStop_Click;
            buttonPanel.Controls.Add(btnStop);

            var btnAbout = new Button { Text = "ℹ️ Info", Width = 60 };
            btnAbout.Click += BtnAbout_Click;
            buttonPanel.Controls.Add(btnAbout);

            panel.Controls.Add(buttonPanel, 0, 22);
            panel.SetColumnSpan(buttonPanel, 2);

            // Status e Log
            panel.Controls.Add(new Label { Text = "=== STATUS E LOG ===", Font = new Font("Arial", 10, FontStyle.Bold) }, 0, 23);
            panel.SetColumnSpan(panel.Controls[panel.Controls.Count - 1], 2);

            lblStatus = new Label { Text = "Servizio fermo", ForeColor = Color.Red, Width = 400 };
            panel.Controls.Add(lblStatus, 0, 24);
            panel.SetColumnSpan(lblStatus, 2);

            txtLog = new TextBox 
            { 
                Multiline = true, 
                ScrollBars = ScrollBars.Vertical, 
                ReadOnly = true, 
                Height = 100, 
                Width = 400 
            };
            panel.Controls.Add(txtLog, 0, 25);
            panel.SetColumnSpan(txtLog, 2);

            this.Controls.Add(panel);
        }

        private void InitializeTrayIcon()
        {
            _trayIcon = new NotifyIcon()
            {
                Icon = SystemIcons.Application,
                Text = $"{AppInfo.Name} v{AppInfo.Version}",
                Visible = true
            };

            var contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("📧 Controlla Email Ora", null, async (s, e) => await ForceEmailCheck());
            contextMenu.Items.Add("📄 Mostra Finestra", null, (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; });
            contextMenu.Items.Add("🔄 Riavvia Servizio", null, async (s, e) => await RestartService());
            contextMenu.Items.Add("-");
            contextMenu.Items.Add("ℹ️ Info Stato", null, (s, e) => ShowStatusInfo());
            contextMenu.Items.Add("📁 Apri Cartella Log", null, (s, e) => OpenLogFolder());
            contextMenu.Items.Add("-");
            contextMenu.Items.Add($"👨‍💻 About {AppInfo.Name}", null, (s, e) => ShowAboutDialog());
            contextMenu.Items.Add("❌ Esci", null, (s, e) => { Application.Exit(); });

            _trayIcon.ContextMenuStrip = contextMenu;
            _trayIcon.DoubleClick += (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; };
        }

        // Event handlers e metodi principali
        private void LoadPrinters()
        {
            cmbPrinter.Items.Clear();
            cmbPrinter.Items.Add("(Stampante predefinita)");
            foreach (string printerName in PrinterSettings.InstalledPrinters)
            {
                cmbPrinter.Items.Add(printerName);
            }
            cmbPrinter.SelectedIndex = 0;
        }

        private void LoadSettings()
        {
            try
            {
                txtEmailServer.Text = _settings.EmailServer;
                txtEmailPort.Value = _settings.EmailPort;
                txtEmailUsername.Text = _settings.EmailUsername;
                txtEmailPassword.Text = _settings.EmailPassword;
                
                var printerName = _settings.PrinterName;
                if (string.IsNullOrEmpty(printerName))
                    cmbPrinter.SelectedIndex = 0;
                else
                {
                    var index = cmbPrinter.Items.IndexOf(printerName);
                    cmbPrinter.SelectedIndex = index > 0 ? index : 0;
                }

                chkSendConfirmation.Checked = _settings.SendConfirmation;
                chkDeleteAfterPrint.Checked = _settings.DeleteAfterPrint;
                chkAutoStart.Checked = _settings.AutoStart;
                txtInterval.Value = _settings.CheckInterval;

                LoadPrintMethodSelection();
                LoadConfirmationModeSelection();
                UpdatePrintInfo();
            }
            catch
            {
                txtEmailServer.Text = "imap.gmail.com";
                txtEmailPort.Value = 993;
            }
        }

        private void LoadPrintMethodSelection()
        {
            switch (_printSettings.PreferredMethod)
            {
                case PrintMethod.NetPrintDocument:
                    rbPrintNet.Checked = true;
                    break;
                case PrintMethod.SumatraPDF:
                    rbPrintSumatra.Checked = true;
                    break;
                case PrintMethod.WindowsShell:
                    rbPrintShell.Checked = true;
                    break;
                case PrintMethod.Auto:
                default:
                    rbPrintAuto.Checked = true;
                    break;
            }
        }

        private void LoadConfirmationModeSelection()
        {
            txtTimeout.Value = _printSettings.ConfirmationTimeout;

            switch (_printSettings.ConfirmationMode)
            {
                case PrintConfirmationMode.TimedConfirmation:
                    rbConfirmTimed.Checked = true;
                    txtTimeout.Enabled = true;
                    break;
                case PrintConfirmationMode.ManualConfirmation:
                    rbConfirmManual.Checked = true;
                    txtTimeout.Enabled = false;
                    break;
                case PrintConfirmationMode.Automatic:
                default:
                    rbConfirmAuto.Checked = true;
                    txtTimeout.Enabled = false;
                    break;
            }
        }

        private void ShowMethodInfo(string info)
        {
            if (lblMethodInfo != null)
            {
                lblMethodInfo.Text = info;
                lblMethodInfo.ForeColor = Color.DarkBlue;
            }
        }

        private void ShowConfirmInfo(string info)
        {
            if (lblConfirmInfo != null)
            {
                lblConfirmInfo.Text = info;
                lblConfirmInfo.ForeColor = Color.DarkGreen;
            }
        }

        private void UpdatePrintInfo()
        {
            var info = $"Stampa: {(_printSettings.PaperSize)} - {(_printSettings.Duplex ? "Fronte/Retro" : "Solo Fronte")} - {(_printSettings.FitToPage ? "Adatta alla pagina" : "Dimensione originale")}";
            lblPrintInfo.Text = info;
            lblPrintInfo.ForeColor = string.IsNullOrEmpty(_printSettings.PrinterSettingsData) ? Color.Red : Color.Green;
        }

        private void BtnConfigurePrint_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbPrinter.SelectedIndex <= 0)
                {
                    MessageBox.Show("Seleziona prima una stampante!", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var selectedPrinter = cmbPrinter.Text;
                
                using (var printDoc = new PrintDocument())
                {
                    printDoc.PrinterSettings.PrinterName = selectedPrinter;
                    
                    if (!printDoc.PrinterSettings.IsValid)
                    {
                        MessageBox.Show($"Stampante non valida: {selectedPrinter}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    using (var printDialog = new PrintDialog())
                    {
                        printDialog.Document = printDoc;
                        printDialog.UseEXDialog = true;
                        printDialog.AllowPrintToFile = false;
                        printDialog.AllowSelection = false;
                        printDialog.AllowSomePages = false;
                        printDialog.ShowHelp = false;

                        if (printDialog.ShowDialog(this) == DialogResult.OK)
                        {
                            SavePrintSettings(printDoc.PrinterSettings, printDoc.DefaultPageSettings);
                            
                            MessageBox.Show("Impostazioni di stampa salvate con successo!", 
                                          "Configurazione Completata", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            
                            UpdatePrintInfo();
                            LogMessage("Impostazioni di stampa configurate");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore nella configurazione stampa: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SavePrintSettings(PrinterSettings printerSettings, PageSettings pageSettings)
        {
            _printSettings.PrinterName = printerSettings.PrinterName;
            _printSettings.Duplex = printerSettings.Duplex != Duplex.Simplex;
            _printSettings.PaperSize = pageSettings.PaperSize.PaperName;
            _printSettings.Landscape = pageSettings.Landscape;
            _printSettings.Copies = printerSettings.Copies;
            _printSettings.ColorPrinting = printerSettings.SupportsColor && !printerSettings.DefaultPageSettings.Color;

            try
            {
                var settingsData = new
                {
                    Duplex = printerSettings.Duplex,
                    PaperSizeName = pageSettings.PaperSize.PaperName,
                    PaperSizeWidth = pageSettings.PaperSize.Width,
                    PaperSizeHeight = pageSettings.PaperSize.Height,
                    Landscape = pageSettings.Landscape,
                    Copies = printerSettings.Copies,
                    Margins = new { 
                        Left = pageSettings.Margins.Left, 
                        Right = pageSettings.Margins.Right, 
                        Top = pageSettings.Margins.Top, 
                        Bottom = pageSettings.Margins.Bottom 
                    }
                };
                
                _printSettings.PrinterSettingsData = JsonSerializer.Serialize(settingsData);
            }
            catch
            {
                _printSettings.PrinterSettingsData = "configured";
            }

            _printSettings.Save();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                _settings.EmailServer = txtEmailServer.Text;
                _settings.EmailPort = (int)txtEmailPort.Value;
                _settings.EmailUsername = txtEmailUsername.Text;
                _settings.EmailPassword = txtEmailPassword.Text;
                _settings.PrinterName = cmbPrinter.SelectedIndex == 0 ? "" : cmbPrinter.Text;
                _settings.SendConfirmation = chkSendConfirmation.Checked;
                _settings.DeleteAfterPrint = chkDeleteAfterPrint.Checked;
                _settings.AutoStart = chkAutoStart.Checked;
                _settings.CheckInterval = (int)txtInterval.Value;
                
                // Salva metodo di stampa preferito
                if (rbPrintNet.Checked)
                    _printSettings.PreferredMethod = PrintMethod.NetPrintDocument;
                else if (rbPrintSumatra.Checked)
                    _printSettings.PreferredMethod = PrintMethod.SumatraPDF;
                else if (rbPrintShell.Checked)
                    _printSettings.PreferredMethod = PrintMethod.WindowsShell;
                else
                    _printSettings.PreferredMethod = PrintMethod.Auto;

                // Salva modalità di conferma stampa
                if (rbConfirmTimed.Checked)
                    _printSettings.ConfirmationMode = PrintConfirmationMode.TimedConfirmation;
                else if (rbConfirmManual.Checked)
                    _printSettings.ConfirmationMode = PrintConfirmationMode.ManualConfirmation;
                else
                    _printSettings.ConfirmationMode = PrintConfirmationMode.Automatic;

                _printSettings.ConfirmationTimeout = (int)txtTimeout.Value;

                _printSettings.Save();
                _settings.Save();

                SetAutoStart(chkAutoStart.Checked);

                MessageBox.Show("Impostazioni salvate con successo!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogMessage("Impostazioni salvate");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore nel salvare le impostazioni: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetAutoStart(bool enabled)
        {
            try
            {
                var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
                if (enabled)
                {
                    key?.SetValue("EmailPrintService", Application.ExecutablePath);
                }
                else
                {
                    key?.DeleteValue("EmailPrintService", false);
                }
                key?.Close();
            }
            catch
            {
                // Ignora errori di registro
            }
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtEmailUsername.Text) || string.IsNullOrWhiteSpace(txtEmailPassword.Text))
            {
                MessageBox.Show("Inserisci username e password email!", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                await StartServiceInternal();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore nell'avvio del servizio: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            _isServiceRunning = false;
            lblStatus.Text = "Servizio fermo";
            lblStatus.ForeColor = Color.Red;

            _service?.Stop();
            LogMessage("Servizio fermato");
        }

        private void BtnAbout_Click(object sender, EventArgs e)
        {
            ShowAboutDialog();
        }

        private void ShowAboutDialog()
        {
            using (var aboutForm = new AboutDialog())
            {
                aboutForm.ShowDialog(this);
            }
        }

        // Altri metodi helper
        private async Task ForceEmailCheck()
        {
            try
            {
                if (!_isServiceRunning)
                {
                    _trayIcon.ShowBalloonTip(3000, "Email Print Service", "Servizio non avviato!", ToolTipIcon.Warning);
                    return;
                }

                _trayIcon.ShowBalloonTip(2000, "Controllo Email", "Controllo manuale email in corso...", ToolTipIcon.Info);
                LogMessage("🔍 Controllo email forzato dall'utente");
                
                await _service.ForceEmailCheck();
                
                _trayIcon.ShowBalloonTip(2000, "Controllo Completato", "Controllo email completato.", ToolTipIcon.Info);
            }
            catch (Exception ex)
            {
                _trayIcon.ShowBalloonTip(3000, "Errore", $"Errore nel controllo email: {ex.Message}", ToolTipIcon.Error);
                LogMessage($"❌ Errore controllo forzato: {ex.Message}");
            }
        }

        private async Task RestartService()
        {
            try
            {
                if (_isServiceRunning)
                {
                    BtnStop_Click(null, null);
                    await Task.Delay(2000);
                }
                
                if (!string.IsNullOrWhiteSpace(txtEmailUsername.Text) && !string.IsNullOrWhiteSpace(txtEmailPassword.Text))
                {
                    await StartServiceInternal();
                    _trayIcon.ShowBalloonTip(2000, "Servizio Riavviato", "Servizio riavviato con successo!", ToolTipIcon.Info);
                }
                else
                {
                    _trayIcon.ShowBalloonTip(3000, "Configurazione Mancante", "Configura email prima di avviare il servizio!", ToolTipIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                _trayIcon.ShowBalloonTip(3000, "Errore Riavvio", $"Errore: {ex.Message}", ToolTipIcon.Error);
            }
        }

        private async Task StartServiceInternal()
        {
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            _isServiceRunning = true;
            lblStatus.Text = "Servizio in avvio...";
            lblStatus.ForeColor = Color.Orange;

            try
            {
                _service.Configure(
                    txtEmailServer.Text,
                    (int)txtEmailPort.Value,
                    txtEmailUsername.Text,
                    txtEmailPassword.Text,
                    cmbPrinter.SelectedIndex == 0 ? "" : cmbPrinter.Text,
                    chkSendConfirmation.Checked,
                    chkDeleteAfterPrint.Checked,
                    (int)txtInterval.Value
                );

                _service.OnStatusChanged += Service_OnStatusChanged;
                _service.OnLogMessage += Service_OnLogMessage;
                _service.UpdatePrintSettings();

                await _service.StartAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore nell'avvio del servizio: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                BtnStop_Click(null, null);
                throw;
            }
        }

        private void ShowStatusInfo()
        {
            var status = _isServiceRunning ? "🟢 ATTIVO" : "🔴 FERMO";
            var nextCheck = _isServiceRunning ? $"Prossimo controllo tra: {_settings.CheckInterval} secondi" : "Servizio non attivo";
            var method = GetCurrentPrintMethod();
            
            var info = $"Stato Servizio: {status}\n" +
                      $"Email: {_settings.EmailUsername}\n" +
                      $"Intervallo: {_settings.CheckInterval} secondi\n" +
                      $"Metodo Stampa: {method}\n" +
                      $"{nextCheck}";
            
            MessageBox.Show(info, "Stato Email Print Service", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string GetCurrentPrintMethod()
        {
            return _printSettings.PreferredMethod switch
            {
                PrintMethod.NetPrintDocument => ".NET PrintDocument",
                PrintMethod.SumatraPDF => "SumatraPDF",
                PrintMethod.WindowsShell => "Windows Shell",
                PrintMethod.Auto => "Automatico",
                _ => "Non configurato"
            };
        }

        private void OpenLogFolder()
        {
            try
            {
                var logFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailPrintService");
                if (Directory.Exists(logFolder))
                {
                    Process.Start("explorer.exe", logFolder);
                }
                else
                {
                    MessageBox.Show("Cartella log non ancora creata.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Impossibile aprire cartella: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Service_OnStatusChanged(string status, bool isConnected)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string, bool>(Service_OnStatusChanged), status, isConnected);
                return;
            }

            lblStatus.Text = status;
            lblStatus.ForeColor = isConnected ? Color.Green : Color.Red;
            _trayIcon.Text = $"{AppInfo.Name} v{AppInfo.Version} - {status}";
        }

        private void Service_OnLogMessage(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(Service_OnLogMessage), message);
                return;
            }

            LogMessage(message);
        }

        private void LogMessage(string message)
        {
            var logEntry = $"{DateTime.Now:HH:mm:ss} - {message}";
            txtLog.AppendText(logEntry + Environment.NewLine);
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(value);
            if (!value) this.Hide();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
                _trayIcon.ShowBalloonTip(2000, "Email Print Service", "L'applicazione continua a funzionare in background", ToolTipIcon.Info);
            }
            else
            {
                _trayIcon.Dispose();
                _service?.Stop();
                base.OnFormClosing(e);
            }
        }
    }

    // Classe del servizio email (versione semplificata)
    public class EmailPrintService
    {
        private ImapClient _imapClient;
        private SmtpClient _smtpClient;
        private string _emailServer, _emailUsername, _emailPassword, _printerName;
        private int _emailPort, _checkInterval;
        private bool _sendConfirmation, _deleteAfterPrint, _isRunning;
        private string _tempFolder;
        private PrintSettings _printSettings;

        public event Action<string, bool> OnStatusChanged;
        public event Action<string> OnLogMessage;

        public EmailPrintService()
        {
            _tempFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailPrintService");
            Directory.CreateDirectory(_tempFolder);
            _printSettings = PrintSettings.Load();
        }

        public void Configure(string server, int port, string username, string password, 
                            string printer, bool sendConfirmation, bool deleteAfterPrint, int interval)
        {
            _emailServer = server;
            _emailPort = port;
            _emailUsername = username;
            _emailPassword = password;
            _printerName = printer;
            _sendConfirmation = sendConfirmation;
            _deleteAfterPrint = deleteAfterPrint;
            _checkInterval = interval * 1000;
            
            _printSettings = PrintSettings.Load();
        }

        public void UpdatePrintSettings()
        {
            _printSettings = PrintSettings.Load();
        }

        public async Task ForceEmailCheck()
        {
            if (!_isRunning || _imapClient?.IsConnected != true)
            {
                throw new Exception("Servizio non avviato o non connesso");
            }

            try
            {
                OnLogMessage?.Invoke("🔍 Controllo email forzato dall'utente");
                var inbox = _imapClient.Inbox;
                
                if (!inbox.IsOpen)
                {
                    await inbox.OpenAsync(FolderAccess.ReadWrite);
                }
                
                await ProcessNewEmails(inbox);
                OnLogMessage?.Invoke("✅ Controllo email forzato completato");
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"❌ Errore nel controllo forzato: {ex.Message}");
                throw;
            }
        }

        public async Task StartAsync()
        {
            _isRunning = true;
            OnStatusChanged?.Invoke("Connessione in corso...", false);

            try
            {
                _imapClient = new ImapClient();
                await _imapClient.ConnectAsync(_emailServer, _emailPort, true);
                await _imapClient.AuthenticateAsync(_emailUsername, _emailPassword);

                if (_sendConfirmation)
                {
                    _smtpClient = new SmtpClient();
                    var smtpServer = _emailServer.Replace("imap", "smtp");
                    var smtpPort = _emailPort == 993 ? 587 : 25;
                    await _smtpClient.ConnectAsync(smtpServer, smtpPort, false);
                    await _smtpClient.AuthenticateAsync(_emailUsername, _emailPassword);
                }

                OnStatusChanged?.Invoke("Connesso - In ascolto", true);
                OnLogMessage?.Invoke("Servizio avviato con successo");

                var inbox = _imapClient.Inbox;
                await inbox.OpenAsync(FolderAccess.ReadWrite);

                while (_isRunning)
                {
                    await ProcessNewEmails(inbox);
                    await Task.Delay(_checkInterval);
                }
            }
            catch (Exception ex)
            {
                OnStatusChanged?.Invoke($"Errore: {ex.Message}", false);
                OnLogMessage?.Invoke($"Errore: {ex.Message}");
                throw;
            }
        }

        private async Task ProcessNewEmails(IMailFolder inbox)
        {
            try
            {
                var uids = await inbox.SearchAsync(SearchQuery.NotSeen);
                
                foreach (var uid in uids)
                {
                    var message = await inbox.GetMessageAsync(uid);
                    var success = await ProcessEmail(message);
                    
                    if (_deleteAfterPrint && success)
                    {
                        await inbox.AddFlagsAsync(uid, MessageFlags.Deleted, true);
                        OnLogMessage?.Invoke($"Email eliminata dopo stampa da {message.From.FirstOrDefault()?.Name}");
                    }
                    else
                    {
                        await inbox.AddFlagsAsync(uid, MessageFlags.Seen, true);
                    }
                }

                if (_deleteAfterPrint && uids.Count > 0)
                {
                    await inbox.ExpungeAsync();
                }
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"Errore processamento email: {ex.Message}");
            }
        }

        private async Task<bool> ProcessEmail(MimeMessage message)
        {
            var sender = message.From.FirstOrDefault()?.Name ?? message.From.FirstOrDefault()?.ToString() ?? "Sconosciuto";
            OnLogMessage?.Invoke($"Email ricevuta da: {sender}");

            var pdfAttachments = message.Attachments
                .OfType<MimePart>()
                .Where(x => x.ContentType.MimeType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (!pdfAttachments.Any())
            {
                OnLogMessage?.Invoke("Nessun allegato PDF trovato");
                return false;
            }

            OnLogMessage?.Invoke($"Trovati {pdfAttachments.Count} allegati PDF");

            bool allSuccess = true;
            var printedFiles = new List<string>();

            foreach (var attachment in pdfAttachments)
            {
                var success = await ProcessPdfAttachment(attachment, sender);
                if (success)
                {
                    printedFiles.Add(attachment.FileName ?? "document.pdf");
                }
                allSuccess = allSuccess && success;
            }

            if (_sendConfirmation && printedFiles.Any())
            {
                await SendConfirmationEmail(message, printedFiles, sender);
            }

            return allSuccess;
        }

        private async Task<bool> ProcessPdfAttachment(MimePart attachment, string sender)
        {
            try
            {
                var fileName = attachment.FileName ?? $"document_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                var filePath = Path.Combine(_tempFolder, fileName);

                using (var stream = File.Create(filePath))
                {
                    await attachment.Content.DecodeToAsync(stream);
                }

                if (IsValidPdf(filePath))
                {
                    await PrintPdf(filePath, sender);
                    OnLogMessage?.Invoke($"Stampato: {fileName} da {sender}");
                    File.Delete(filePath);
                    return true;
                }
                else
                {
                    OnLogMessage?.Invoke($"PDF non valido: {fileName}");
                    File.Delete(filePath);
                    return false;
                }
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"Errore stampa allegato: {ex.Message}");
                return false;
            }
        }

        private bool IsValidPdf(string filePath)
        {
            try
            {
                using (var document = PdfReader.Open(filePath, PdfDocumentOpenMode.ReadOnly))
                {
                    return document.PageCount > 0;
                }
            }
            catch
            {
                // Se PdfSharp fallisce, prova una validazione semplice
                try
                {
                    var fileBytes = File.ReadAllBytes(filePath);
                    return fileBytes.Length > 100 && 
                           fileBytes[0] == 0x25 && fileBytes[1] == 0x50 && 
                           fileBytes[2] == 0x44 && fileBytes[3] == 0x46; // "%PDF"
                }
                catch
                {
                    return false;
                }
            }
        }

        private async Task PrintPdf(string filePath, string sender)
        {
            try
            {
                await Task.Run(() => PrintPdfDirect(filePath, _printerName));
                OnLogMessage?.Invoke($"Stampato: {Path.GetFileName(filePath)} da {sender}");

                var logPath = Path.Combine(_tempFolder, "print_log.txt");
                var logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Stampato '{Path.GetFileName(filePath)}' da {sender}";
                File.AppendAllText(logPath, logEntry + Environment.NewLine);
            }
            catch (Exception ex)
            {
                throw new Exception($"Errore durante la stampa: {ex.Message}");
            }
        }

        private void PrintPdfDirect(string filePath, string printerName)
        {
            try
            {
                // Metodo semplificato per stampare PDF
                var psi = new ProcessStartInfo
                {
                    FileName = filePath,
                    Verb = "print",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true
                };

                using (var process = Process.Start(psi))
                {
                    process?.WaitForExit(30000);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Impossibile stampare il PDF: {ex.Message}");
            }
        }

        private async Task SendConfirmationEmail(MimeMessage originalMessage, List<string> printedFiles, string sender)
        {
            try
            {
                if (_smtpClient == null) return;

                var confirmationMessage = new MimeMessage();
                confirmationMessage.From.Add(new MailboxAddress("Sistema Stampa", _emailUsername));
                confirmationMessage.To.Add(originalMessage.From.FirstOrDefault());
                confirmationMessage.Subject = $"Conferma stampa - {originalMessage.Subject}";

                var bodyText = $@"
Gentile {sender},

La sua richiesta di stampa è stata elaborata con successo.

Dettagli:
• Data e ora: {DateTime.Now:dd/MM/yyyy HH:mm}
• File stampati: {string.Join(", ", printedFiles)}
• Numero di documenti: {printedFiles.Count}

Cordiali saluti,
Sistema di Stampa Automatica
";

                confirmationMessage.Body = new TextPart("plain") { Text = bodyText };

                await _smtpClient.SendAsync(confirmationMessage);
                OnLogMessage?.Invoke($"Email di conferma inviata a {sender}");
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"Errore invio conferma: {ex.Message}");
            }
        }

        public void Stop()
        {
            _isRunning = false;
            _imapClient?.Disconnect(true);
            _smtpClient?.Disconnect(true);
            OnStatusChanged?.Invoke("Servizio fermo", false);
        }
    }

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
}
