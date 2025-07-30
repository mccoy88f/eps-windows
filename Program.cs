using System;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit;
using MailKit.Search;
using MailKit.Security;
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
        public const string Version = "1.0.1";
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
        public string SecureSender { get; set; } = ""; // Email del mittente sicuro (opzionale)

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
        private ToolTip _toolTip; // Aggiungi componente ToolTip

        // Controlli UI
        private TextBox txtEmailServer, txtEmailUsername, txtEmailPassword, txtLog;
        private NumericUpDown txtEmailPort, txtInterval, txtTimeout;
        private ComboBox cmbPrinter;
        private CheckBox chkSendConfirmation, chkDeleteAfterPrint, chkAutoStart;
        private Button btnSave, btnStart, btnStop, btnConfigurePrint;
        private Label lblStatus, lblPrintInfo, lblConfirmInfo;
        private RadioButton rbPrintAuto, rbPrintNet, rbPrintSumatra, rbPrintShell;
        private RadioButton rbConfirmAuto, rbConfirmTimed, rbConfirmManual;
        private TextBox txtSecureSender; // Campo per il mittente sicuro

        public MainForm()
        {
            try
            {
                _toolTip = new ToolTip(); // Inizializza prima
                InitializeComponent();
                InitializeTrayIcon();
                _service = new EmailPrintService();
                _settings = AppSettings.Load();
                _printSettings = PrintSettings.Load();
                LoadSettings();
                
                // Pulisce il log file all'avvio
                ClearLogFile();
                
                this.Show();
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore inizializzazione MainForm: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.Text = $"{AppInfo.Name} v{AppInfo.Version} - by {AppInfo.Author}";
            this.Size = new Size(1000, 700); // Più largo che alto
            this.FormBorderStyle = FormBorderStyle.Sizable; // Ridimensionabile
            this.MinimumSize = new Size(800, 600); // Dimensione minima
            this.MaximizeBox = true; // Abilita massimizzazione
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = true;
            this.Icon = CreateAppIcon();

            CreateControls();
        }

        private void CreateControls()
        {
            // Layout principale con 3 colonne per i blocchi centrali
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 2,
                Padding = new Padding(10)
            };

            // Configurazione delle colonne (25% - 25% - 25% - 25% per il log)
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            // Configurazione delle righe (70% per i controlli, 30% per il log)
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 70));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30));

            // BLOCCO 1: Configurazione Email
            var emailPanel = CreateEmailBlock();
            mainPanel.Controls.Add(emailPanel, 0, 0);

            // BLOCCO 2: Configurazione Stampante e processi di stampa
            var printerPanel = CreatePrinterBlock();
            mainPanel.Controls.Add(printerPanel, 1, 0);

            // BLOCCO 3: Opzioni e conferme
            var optionsPanel = CreateOptionsBlock();
            mainPanel.Controls.Add(optionsPanel, 2, 0);

            // AREA LOG (occupa tutta la larghezza nella riga inferiore)
            var logPanel = CreateLogBlock();
            mainPanel.Controls.Add(logPanel, 0, 1);
            mainPanel.SetColumnSpan(logPanel, 4);

            this.Controls.Add(mainPanel);
        }

        private Panel CreateEmailBlock()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(10)
            };

            var title = new Label
            {
                Text = "📧 Configurazione Email",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter
            };
            panel.Controls.Add(title);

            var contentPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 8,
                Padding = new Padding(5)
            };

            // Server IMAP
            contentPanel.Controls.Add(new Label { Text = "Server IMAP:", Anchor = AnchorStyles.Left }, 0, 0);
            txtEmailServer = new TextBox { Width = 200, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            contentPanel.Controls.Add(txtEmailServer, 1, 0);

            // Porta
            contentPanel.Controls.Add(new Label { Text = "Porta:", Anchor = AnchorStyles.Left }, 0, 1);
            txtEmailPort = new NumericUpDown { Minimum = 1, Maximum = 65535, Value = 993, Width = 200 };
            contentPanel.Controls.Add(txtEmailPort, 1, 1);

            // Username
            contentPanel.Controls.Add(new Label { Text = "Username:", Anchor = AnchorStyles.Left }, 0, 2);
            txtEmailUsername = new TextBox { Width = 200, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            contentPanel.Controls.Add(txtEmailUsername, 1, 2);

            // Password
            contentPanel.Controls.Add(new Label { Text = "Password:", Anchor = AnchorStyles.Left }, 0, 3);
            txtEmailPassword = new TextBox { UseSystemPasswordChar = true, Width = 200, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            _toolTip.SetToolTip(txtEmailPassword, "Per Gmail: usa una 'Password per le App' invece della password normale");
            contentPanel.Controls.Add(txtEmailPassword, 1, 3);

            // Mittente sicuro
            contentPanel.Controls.Add(new Label { Text = "Mittente sicuro:", Anchor = AnchorStyles.Left }, 0, 4);
            txtSecureSender = new TextBox { Width = 200, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            _toolTip.SetToolTip(txtSecureSender, "Se specificato, processa solo email da questi indirizzi. Più indirizzi separati da ;");
            contentPanel.Controls.Add(txtSecureSender, 1, 4);

            // Intervallo controllo
            contentPanel.Controls.Add(new Label { Text = "Intervallo (sec):", Anchor = AnchorStyles.Left }, 0, 5);
            txtInterval = new NumericUpDown { Minimum = 5, Maximum = 3600, Value = 10, Width = 200 };
            contentPanel.Controls.Add(txtInterval, 1, 5);

            panel.Controls.Add(contentPanel);
            return panel;
        }

        private Panel CreatePrinterBlock()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(10)
            };

            var title = new Label
            {
                Text = "🖨️ Configurazione Stampante",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter
            };
            panel.Controls.Add(title);

            var contentPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 8,
                Padding = new Padding(5)
            };

            // Stampante
            contentPanel.Controls.Add(new Label { Text = "Stampante:", Anchor = AnchorStyles.Left }, 0, 0);
            cmbPrinter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
            LoadPrinters();
            _toolTip.SetToolTip(cmbPrinter, "Seleziona una stampante specifica o usa quella predefinita");
            contentPanel.Controls.Add(cmbPrinter, 1, 0);

            // Configura stampa
            btnConfigurePrint = new Button { Text = "Configura Stampa", Width = 150 };
            btnConfigurePrint.Click += BtnConfigurePrint_Click;
            _toolTip.SetToolTip(btnConfigurePrint, "Apre il dialog di Windows per configurare formato carta, fronte/retro, etc.");
            contentPanel.Controls.Add(btnConfigurePrint, 1, 1);

            // Metodo di stampa
            contentPanel.Controls.Add(new Label { Text = "Metodo:", Anchor = AnchorStyles.Left }, 0, 2);
            contentPanel.SetColumnSpan(contentPanel.Controls[contentPanel.Controls.Count - 1], 2);
            
            rbPrintAuto = new RadioButton { Text = "🤖 Automatico", Width = 200, Checked = true };
            contentPanel.Controls.Add(rbPrintAuto, 0, 3);
            contentPanel.SetColumnSpan(rbPrintAuto, 2);

            rbPrintSumatra = new RadioButton { Text = "⚡ SumatraPDF", Width = 200 };
            contentPanel.Controls.Add(rbPrintSumatra, 0, 4);
            contentPanel.SetColumnSpan(rbPrintSumatra, 2);

            rbPrintNet = new RadioButton { Text = "🏗️ .NET", Width = 200 };
            contentPanel.Controls.Add(rbPrintNet, 0, 5);
            contentPanel.SetColumnSpan(rbPrintNet, 2);

            rbPrintShell = new RadioButton { Text = "🪟 Windows Shell", Width = 200 };
            contentPanel.Controls.Add(rbPrintShell, 0, 6);
            contentPanel.SetColumnSpan(rbPrintShell, 2);

            // Info stampa
            lblPrintInfo = new Label { Text = "Clicca 'Configura Stampa' per impostare formato, fronte/retro, etc.", ForeColor = Color.Blue, AutoSize = true };
            contentPanel.Controls.Add(lblPrintInfo, 0, 7);
            contentPanel.SetColumnSpan(lblPrintInfo, 2);

            panel.Controls.Add(contentPanel);
            return panel;
        }

        private Panel CreateOptionsBlock()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(10)
            };

            var title = new Label
            {
                Text = "⚙️ Opzioni e Conferme",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter
            };
            panel.Controls.Add(title);

            var contentPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 10,
                Padding = new Padding(5)
            };

            // Conferma stampa
            contentPanel.Controls.Add(new Label { Text = "Conferma:", Anchor = AnchorStyles.Left }, 0, 0);
            contentPanel.SetColumnSpan(contentPanel.Controls[contentPanel.Controls.Count - 1], 2);
            
            rbConfirmAuto = new RadioButton { Text = "🤖 Automatica", Width = 200, Checked = true };
            rbConfirmAuto.CheckedChanged += (s, e) => { if (rbConfirmAuto.Checked) ShowConfirmInfo("Stampa diretta"); };
            contentPanel.Controls.Add(rbConfirmAuto, 0, 1);
            contentPanel.SetColumnSpan(rbConfirmAuto, 2);

            rbConfirmTimed = new RadioButton { Text = "⏰ Con timer", Width = 200 };
            rbConfirmTimed.CheckedChanged += (s, e) => { if (rbConfirmTimed.Checked) ShowConfirmInfo("Popup con timer"); };
            contentPanel.Controls.Add(rbConfirmTimed, 0, 2);
            contentPanel.SetColumnSpan(rbConfirmTimed, 2);

            rbConfirmManual = new RadioButton { Text = "👤 Manuale", Width = 200 };
            rbConfirmManual.CheckedChanged += (s, e) => { if (rbConfirmManual.Checked) ShowConfirmInfo("Popup obbligatorio"); };
            contentPanel.Controls.Add(rbConfirmManual, 0, 3);
            contentPanel.SetColumnSpan(rbConfirmManual, 2);

            // Timeout conferma
            contentPanel.Controls.Add(new Label { Text = "Timeout (sec):", Anchor = AnchorStyles.Left }, 0, 4);
            txtTimeout = new NumericUpDown { Minimum = 5, Maximum = 300, Value = 15, Width = 200 };
            contentPanel.Controls.Add(txtTimeout, 1, 4);

            // Opzioni
            chkSendConfirmation = new CheckBox { Text = "📧 Invia conferma email", Width = 200 };
            _toolTip.SetToolTip(chkSendConfirmation, "Invia email di conferma al mittente dopo la stampa");
            contentPanel.Controls.Add(chkSendConfirmation, 0, 5);
            contentPanel.SetColumnSpan(chkSendConfirmation, 2);

            chkDeleteAfterPrint = new CheckBox { Text = "🗑️ Elimina email dopo stampa", Width = 200 };
            _toolTip.SetToolTip(chkDeleteAfterPrint, "Elimina l'email dal server dopo aver stampato gli allegati");
            contentPanel.Controls.Add(chkDeleteAfterPrint, 0, 6);
            contentPanel.SetColumnSpan(chkDeleteAfterPrint, 2);

            chkAutoStart = new CheckBox { Text = "🚀 Avvio automatico Windows", Width = 200 };
            _toolTip.SetToolTip(chkAutoStart, "Avvia automaticamente l'applicazione all'avvio di Windows");
            contentPanel.Controls.Add(chkAutoStart, 0, 7);
            contentPanel.SetColumnSpan(chkAutoStart, 2);

            // Info conferma
            lblConfirmInfo = new Label { Text = "Modalità conferma stampa", ForeColor = Color.Blue, AutoSize = true };
            contentPanel.Controls.Add(lblConfirmInfo, 0, 8);
            contentPanel.SetColumnSpan(lblConfirmInfo, 2);

            panel.Controls.Add(contentPanel);
            return panel;
        }

        private Panel CreateLogBlock()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(10)
            };

            var title = new Label
            {
                Text = "📋 Log Attività",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter
            };
            panel.Controls.Add(title);

            // Area log
            txtLog = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.Black,
                ForeColor = Color.LightGreen,
                Dock = DockStyle.Fill
            };
            panel.Controls.Add(txtLog);

            // Pulsanti sotto il log (da sinistra a destra)
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 40,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(5)
            };

            btnSave = new Button { Text = "💾 Salva Configurazione", Width = 150, Height = 30 };
            btnSave.Click += BtnSave_Click;
            buttonPanel.Controls.Add(btnSave);

            btnStart = new Button { Text = "▶️ Avvia Servizio", Width = 120, Height = 30 };
            btnStart.Click += BtnStart_Click;
            buttonPanel.Controls.Add(btnStart);

            btnStop = new Button { Text = "⏹️ Ferma Servizio", Width = 120, Height = 30 };
            btnStop.Click += BtnStop_Click;
            buttonPanel.Controls.Add(btnStop);

            var btnCheckEmail = new Button { Text = "📧 Controlla Email", Width = 120, Height = 30 };
            btnCheckEmail.Click += async (s, e) => await ForceEmailCheck();
            buttonPanel.Controls.Add(btnCheckEmail);

            var btnResetTimestamp = new Button { Text = "🔄 Reset Timestamp", Width = 120, Height = 30 };
            btnResetTimestamp.Click += (s, e) => ResetEmailTimestamp();
            buttonPanel.Controls.Add(btnResetTimestamp);

            var btnOpenLog = new Button { Text = "📂 Apri Log", Width = 100, Height = 30 };
            btnOpenLog.Click += (s, e) => OpenTodayLog();
            buttonPanel.Controls.Add(btnOpenLog);

            var btnAbout = new Button { Text = "ℹ️ Info", Width = 80, Height = 30 };
            btnAbout.Click += BtnAbout_Click;
            buttonPanel.Controls.Add(btnAbout);

            // Status label
            lblStatus = new Label
            {
                Text = "⏸️ Servizio fermo",
                ForeColor = Color.Red,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                AutoSize = true,
                Anchor = AnchorStyles.Right
            };
            buttonPanel.Controls.Add(lblStatus);

            panel.Controls.Add(buttonPanel);

            return panel;
        }

        private void InitializeTrayIcon()
        {
            _trayIcon = new NotifyIcon()
            {
                Icon = CreateAppIcon(),
                Text = $"{AppInfo.Name} v{AppInfo.Version}",
                Visible = true
            };

            var contextMenu = new ContextMenuStrip();
            contextMenu.AutoClose = true; // Importante: chiude automaticamente
            contextMenu.ShowImageMargin = false; // Rimuove margine icone
            contextMenu.Width = 250; // Larghezza fissa per vedere tutto
            
            contextMenu.Items.Add("📧 Controlla Email Ora", null, async (s, e) => await ForceEmailCheck());
            contextMenu.Items.Add("📄 Mostra Finestra", null, (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; this.BringToFront(); });
            contextMenu.Items.Add("🔄 Riavvia Servizio", null, async (s, e) => await RestartService());
            contextMenu.Items.Add("-");
            contextMenu.Items.Add("⏰ Reset Timestamp Email", null, (s, e) => ResetEmailTimestamp());
            contextMenu.Items.Add("ℹ️ Info Stato", null, (s, e) => ShowStatusInfo());
            contextMenu.Items.Add("📁 Apri Cartella Log", null, (s, e) => OpenLogFolder());
            contextMenu.Items.Add("📄 Mostra Log di Oggi", null, (s, e) => OpenTodayLog());
            contextMenu.Items.Add("-");
            contextMenu.Items.Add($"👨‍💻 About {AppInfo.Name}", null, (s, e) => ShowAboutDialog());
            contextMenu.Items.Add("❌ Esci Completamente", null, (s, e) => ExitApplication());

            _trayIcon.ContextMenuStrip = contextMenu;
            _trayIcon.DoubleClick += (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; this.BringToFront(); };
        }

        private void ExitApplication()
        {
            try
            {
                var result = MessageBox.Show(
                    "Sei sicuro di voler chiudere completamente Email Print Service?\n\nIl servizio smetterà di monitorare le email.",
                    "Conferma Chiusura",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (result == DialogResult.Yes)
                {
                    // Ferma il servizio
                    if (_isServiceRunning)
                    {
                        _service?.Stop();
                    }
                    
                    // Chiusura completa
                    _trayIcon.Visible = false;
                    _toolTip?.Dispose();
                    _trayIcon?.Dispose();
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore durante la chiusura: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _toolTip?.Dispose();
                _trayIcon?.Dispose();
                Application.Exit(); // Forza la chiusura in caso di errore
            }
        }

        private Icon CreateAppIcon()
        {
            try
            {
                // Prova a caricare l'icona dal file app.ico
                var iconPath = Path.Combine(Application.StartupPath, "app.ico");
                if (File.Exists(iconPath))
                {
                    var icon = new Icon(iconPath);
                    // Verifica che l'icona sia valida
                    if (icon.Width > 0 && icon.Height > 0)
                    {
                        return icon;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log dell'errore per debug
                System.Diagnostics.Debug.WriteLine($"Errore caricamento app.ico: {ex.Message}");
            }
            
            // Fallback: crea un'icona personalizzata 16x16 per la system tray
            var bitmap = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bitmap))
            {
                // Sfondo blu
                g.FillRectangle(Brushes.DodgerBlue, 0, 0, 16, 16);
                
                // Bordo bianco
                g.DrawRectangle(Pens.White, 0, 0, 15, 15);
                
                // Simbolo stampante semplificato (rettangoli bianchi)
                g.FillRectangle(Brushes.White, 3, 4, 10, 3);  // Parte superiore
                g.FillRectangle(Brushes.White, 2, 7, 12, 5);  // Corpo stampante
                g.FillRectangle(Brushes.White, 4, 9, 8, 1);   // Carta
                
                // Punto rosso per indicare "email"
                g.FillEllipse(Brushes.Red, 12, 2, 2, 2);
                
                return Icon.FromHandle(bitmap.GetHicon());
            }
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
                txtSecureSender.Text = _settings.SecureSender; // Carica il mittente sicuro
                
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
                _settings.SecureSender = txtSecureSender.Text; // Salva il mittente sicuro
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
            btnStart.Enabled = false; // Disabilita temporaneamente per evitare doppi click
            btnStop.Enabled = false;
            _isServiceRunning = false;
            lblStatus.Text = "Servizio in arresto...";
            lblStatus.ForeColor = Color.Orange;

            try
            {
                // Rimuovi gli event handler PRIMA di fermare il servizio
                if (_service != null)
                {
                    _service.OnStatusChanged -= Service_OnStatusChanged;
                    _service.OnLogMessage -= Service_OnLogMessage;
                    
                    _service.Stop();
                    LogMessage("🛑 Servizio fermato");
                }
                
                lblStatus.Text = "Servizio fermo";
                lblStatus.ForeColor = Color.Red;
                btnStart.Enabled = true;
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Errore durante l'arresto: {ex.Message}");
                lblStatus.Text = "Errore arresto servizio";
                lblStatus.ForeColor = Color.Red;
                btnStart.Enabled = true;
            }
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
                    (int)txtInterval.Value,
                    txtSecureSender.Text // Passa il mittente sicuro
                );

                // Rimuovi vecchi event handler per evitare duplicati
                _service.OnStatusChanged -= Service_OnStatusChanged;
                _service.OnLogMessage -= Service_OnLogMessage;
                
                // Aggiungi nuovi event handler
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
            var lastProcessed = _isServiceRunning ? _service.GetLastProcessedTime().ToString("dd/MM/yyyy HH:mm:ss") : "Non disponibile";
            
            var info = $"Stato Servizio: {status}\n" +
                      $"Email: {_settings.EmailUsername}\n" +
                      $"Intervallo: {_settings.CheckInterval} secondi\n" +
                      $"Metodo Stampa: {method}\n" +
                      $"Ultima email processata: {lastProcessed}\n" +
                      $"{nextCheck}\n\n" +
                      $"📁 Log salvati in:\n%LOCALAPPDATA%\\EmailPrintService\\";
            
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
                    LogMessage($"📁 Aperta cartella log: {logFolder}");
                    LogMessage($"📄 File log di oggi: log_{DateTime.Now:yyyy-MM-dd}.txt");
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

        private async void ResetEmailTimestamp()
        {
            try
            {
                var result = MessageBox.Show(
                    "Questo resetterà il timestamp dell'ultima email processata.\n\n" +
                    "ATTENZIONE: Dopo il reset, il servizio riprocesserà tutte le email ricevute dopo l'ultima email presente nella casella.\n\n" +
                    "Vuoi continuare?",
                    "Reset Timestamp Email",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                if (result == DialogResult.Yes)
                {
                    await _service.ResetLastProcessedTime();
                    LogMessage("⏰ Timestamp email resettato - verranno processate tutte le email ricevute dopo l'ultima email presente");
                    _trayIcon.ShowBalloonTip(3000, "Timestamp Reset", "Timestamp resettato con successo. Il servizio processerà tutte le email ricevute dopo l'ultima email presente.", ToolTipIcon.Info);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore nel reset timestamp: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenTodayLog()
        {
            try
            {
                var logFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailPrintService");
                var logFile = Path.Combine(logFolder, "log.txt");
                
                if (File.Exists(logFile))
                {
                    Process.Start("notepad.exe", logFile);
                    LogMessage($"📄 Aperto log: {logFile}");
                }
                else
                {
                    MessageBox.Show("File log non ancora creato.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Impossibile aprire log: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            
            // Log nell'interfaccia
            txtLog.AppendText(logEntry + Environment.NewLine);
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
            
            // Log su file singolo che si svuota ad ogni avvio
            try
            {
                var logFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailPrintService");
                Directory.CreateDirectory(logFolder);
                var logFile = Path.Combine(logFolder, "log.txt");
                
                var fileLogEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                File.AppendAllText(logFile, fileLogEntry + Environment.NewLine);
            }
            catch
            {
                // Ignora errori di scrittura log su file
            }
        }

        private void ClearLogFile()
        {
            try
            {
                var logFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailPrintService");
                Directory.CreateDirectory(logFolder);
                var logFile = Path.Combine(logFolder, "log.txt");
                
                // Svuota il file di log all'avvio
                if (File.Exists(logFile))
                {
                    File.WriteAllText(logFile, string.Empty);
                }
            }
            catch
            {
                // Ignora errori di pulizia log
            }
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(value);
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
                _toolTip?.Dispose();
                _trayIcon?.Dispose();
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
        private DateTime _lastProcessedTime;
        private CancellationTokenSource _cancellationTokenSource; // Per fermare il loop pulitamente
        private string _secureSender; // Mittente sicuro (opzionale)
        private readonly object _imapLock = new object(); // Lock per sincronizzare l'accesso a ImapClient

        public event Action<string, bool> OnStatusChanged;
        public event Action<string> OnLogMessage;

        public EmailPrintService()
        {
            _tempFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailPrintService");
            Directory.CreateDirectory(_tempFolder);
            _printSettings = PrintSettings.Load();
            
            // Carica l'ultimo timestamp processato o usa ora se è la prima volta
            _lastProcessedTime = LoadLastProcessedTime();
        }

        private DateTime LoadLastProcessedTime()
        {
            try
            {
                var timestampFile = Path.Combine(_tempFolder, "last_processed.txt");
                if (File.Exists(timestampFile))
                {
                    var timestampText = File.ReadAllText(timestampFile);
                    if (DateTime.TryParse(timestampText, out var savedTime))
                    {
                        return savedTime;
                    }
                }
            }
            catch { }
            
            // Se non c'è file o c'è errore, usa ora come default
            // Il timestamp verrà aggiornato durante il primo controllo email
            return DateTime.Now;
        }

        private async Task<DateTime> GetLatestEmailTime()
        {
            try
            {
                var tempClient = new ImapClient();
                await tempClient.ConnectAsync(_emailServer, _emailPort, true);
                await tempClient.AuthenticateAsync(_emailUsername, _emailPassword);
                
                var inbox = tempClient.Inbox;
                await inbox.OpenAsync(FolderAccess.ReadWrite);
                
                if (inbox.Count > 0)
                {
                    // Prendi l'ultima email (la più recente)
                    var lastMessage = await inbox.GetMessageAsync(inbox.Count - 1);
                    var latestTime = lastMessage.Date.DateTime;
                    
                    await tempClient.DisconnectAsync(true);
                    tempClient.Dispose();
                    
                    return latestTime;
                }
                
                await tempClient.DisconnectAsync(true);
                tempClient.Dispose();
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"⚠️ Errore nel leggere l'ultima email: {ex.Message}");
            }
            
            // Se non riesce a leggere, usa ora
            return DateTime.Now;
        }

        private void SaveLastProcessedTime(DateTime timestamp)
        {
            try
            {
                var timestampFile = Path.Combine(_tempFolder, "last_processed.txt");
                File.WriteAllText(timestampFile, timestamp.ToString("yyyy-MM-dd HH:mm:ss"));
                _lastProcessedTime = timestamp;
            }
            catch { }
        }

        public void Configure(string server, int port, string username, string password, 
                            string printer, bool sendConfirmation, bool deleteAfterPrint, int interval,
                            string secureSender) // Aggiunto parametro secureSender
        {
            _emailServer = server;
            _emailPort = port;
            _emailUsername = username;
            _emailPassword = password;
            _printerName = printer;
            _sendConfirmation = sendConfirmation;
            _deleteAfterPrint = deleteAfterPrint;
            _checkInterval = interval * 1000;
            _secureSender = secureSender; // Salva il mittente sicuro
            
            _printSettings = PrintSettings.Load();
        }

        public void UpdatePrintSettings()
        {
            _printSettings = PrintSettings.Load();
        }

        public Task ForceEmailCheck()
        {
            if (!_isRunning)
            {
                throw new Exception("Servizio non avviato");
            }

            if (_cancellationTokenSource?.Token.IsCancellationRequested == true)
            {
                throw new Exception("Servizio in fase di arresto");
            }

            try
            {
                OnLogMessage?.Invoke("🔍 Controllo email forzato dall'utente");
                
                // Usa lock per evitare conflitti con il loop principale
                lock (_imapLock)
                {
                    if (_imapClient?.IsConnected != true)
                    {
                        OnLogMessage?.Invoke("🔄 Riconnessione IMAP necessaria...");
                        _imapClient = new ImapClient();
                        _imapClient.ConnectAsync(_emailServer, _emailPort, true).Wait();
                        _imapClient.AuthenticateAsync(_emailUsername, _emailPassword).Wait();
                        OnLogMessage?.Invoke("✅ Riconnessione IMAP completata");
                    }
                    
                    var inbox = _imapClient.Inbox;
                    
                    if (!inbox.IsOpen)
                    {
                        inbox.OpenAsync(FolderAccess.ReadWrite).Wait();
                    }
                    else if (inbox.Access != FolderAccess.ReadWrite)
                    {
                        // Se la cartella è aperta in modalità diversa, riaprila in read-write
                        inbox.CloseAsync().Wait();
                        inbox.OpenAsync(FolderAccess.ReadWrite).Wait();
                    }
                    
                    OnLogMessage?.Invoke($"⏰ Cercando email ricevute dopo: {_lastProcessedTime:dd/MM/yyyy HH:mm:ss}");
                    ProcessNewEmails(inbox).Wait();
                    OnLogMessage?.Invoke("✅ Controllo email forzato completato");
                }
                
                return Task.CompletedTask;
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
            _cancellationTokenSource = new CancellationTokenSource(); // Nuovo token per questo avvio
            OnStatusChanged?.Invoke("Connessione in corso...", false);

            try
            {
                OnLogMessage?.Invoke($"🔗 Connettendo a {_emailServer}:{_emailPort}...");
                OnLogMessage?.Invoke($"⏰ Servizio avviato alle: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                
                _imapClient = new ImapClient();
                await _imapClient.ConnectAsync(_emailServer, _emailPort, true);
                
                OnLogMessage?.Invoke($"✅ Connesso! Autenticando con {_emailUsername}...");
                
                try
                {
                    await _imapClient.AuthenticateAsync(_emailUsername, _emailPassword);
                    OnLogMessage?.Invoke("✅ Autenticazione IMAP riuscita!");
                }
                catch (AuthenticationException authEx)
                {
                    OnLogMessage?.Invoke($"❌ ERRORE AUTENTICAZIONE: {authEx.Message}");
                    OnLogMessage?.Invoke("🔑 SUGGERIMENTO: Se usi Gmail, serve una 'Password per le App':");
                    OnLogMessage?.Invoke("   1. Vai su https://myaccount.google.com/");
                    OnLogMessage?.Invoke("   2. Sicurezza → Verifica in due passaggi (deve essere ATTIVA)");
                    OnLogMessage?.Invoke("   3. Password per le app → Genera nuova password");
                    OnLogMessage?.Invoke("   4. Usa quella password invece della tua password normale");
                    throw new Exception($"Autenticazione fallita: {authEx.Message}. Controlla username/password o usa Password per le App per Gmail.");
                }

                // Inizializza il timestamp con l'ultima email presente
                var inbox = _imapClient.Inbox;
                await inbox.OpenAsync(FolderAccess.ReadWrite);
                OnLogMessage?.Invoke($"📧 Monitoraggio inbox: {inbox.Count} messaggi totali presenti");
                
                if (inbox.Count > 0)
                {
                    var lastMessage = await inbox.GetMessageAsync(inbox.Count - 1);
                    _lastProcessedTime = lastMessage.Date.DateTime;
                    SaveLastProcessedTime(_lastProcessedTime);
                    OnLogMessage?.Invoke($"📧 Ultima email presente: {_lastProcessedTime:dd/MM/yyyy HH:mm:ss}");
                    OnLogMessage?.Invoke($"🔍 Verranno processate solo le email ricevute DOPO questo orario");
                }
                else
                {
                    _lastProcessedTime = DateTime.Now;
                    SaveLastProcessedTime(_lastProcessedTime);
                    OnLogMessage?.Invoke($"📧 Inbox vuota - timestamp impostato a: {_lastProcessedTime:dd/MM/yyyy HH:mm:ss}");
                }

                if (_sendConfirmation)
                {
                    OnLogMessage?.Invoke("🔗 Connettendo al server SMTP...");
                    _smtpClient = new SmtpClient();
                    var smtpServer = _emailServer.Replace("imap", "smtp");
                    var smtpPort = _emailPort == 993 ? 587 : 25;
                    
                    await _smtpClient.ConnectAsync(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls);
                    
                    try
                    {
                        await _smtpClient.AuthenticateAsync(_emailUsername, _emailPassword);
                        OnLogMessage?.Invoke("✅ Autenticazione SMTP riuscita!");
                    }
                    catch (AuthenticationException authEx)
                    {
                        OnLogMessage?.Invoke($"❌ ERRORE AUTENTICAZIONE SMTP: {authEx.Message}");
                        throw new Exception($"Autenticazione SMTP fallita: {authEx.Message}");
                    }
                }

                OnStatusChanged?.Invoke("Connesso - In ascolto", true);
                OnLogMessage?.Invoke("🚀 Servizio avviato con successo");
                OnLogMessage?.Invoke($"⏰ Controllo automatico email ogni {_checkInterval/1000} secondi");
                if (!string.IsNullOrEmpty(_secureSender))
                {
                    OnLogMessage?.Invoke($"🔒 Mittente sicuro configurato: {_secureSender}");
                }
                else
                {
                    OnLogMessage?.Invoke($"🔓 Nessun mittente sicuro configurato - processa tutte le email");
                }

                // Loop principale con cancellation token
                while (_isRunning && !_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    try
                    {
                        lock (_imapLock)
                        {
                            ProcessNewEmails(inbox).Wait();
                        }
                        
                        await Task.Delay(_checkInterval, _cancellationTokenSource.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        // Task.Delay è stato cancellato, uscire dal loop
                        OnLogMessage?.Invoke("🛑 Loop servizio cancellato");
                        break;
                    }
                    catch (Exception ex)
                    {
                        OnLogMessage?.Invoke($"❌ Errore nel loop principale: {ex.Message}");
                        // Aspetta un po' prima di riprovare
                        await Task.Delay(5000, _cancellationTokenSource.Token);
                    }
                }
                
                OnLogMessage?.Invoke("🔚 Loop servizio terminato");
            }
            catch (Exception ex)
            {
                var errorMsg = $"Errore: {ex.Message}";
                OnStatusChanged?.Invoke(errorMsg, false);
                OnLogMessage?.Invoke($"❌ {errorMsg}");
                throw;
            }
            finally
            {
                // Cleanup sempre eseguito
                OnLogMessage?.Invoke("🧹 Cleanup risorse in corso...");
                CleanupResources();
            }
        }

        private async Task ProcessNewEmails(IMailFolder inbox)
        {
            try
            {
                // Verifica che la cartella sia aperta in modalità corretta
                if (inbox.Access != FolderAccess.ReadWrite)
                {
                    await inbox.OpenAsync(FolderAccess.ReadWrite);
                }
                
                // Cerca TUTTE le email nell'inbox (non solo quelle non lette)
                // Questo risolve il problema delle email che potrebbero essere già marcate come "lette"
                var allUids = await inbox.SearchAsync(SearchQuery.All);
                
                if (allUids.Count > 0)
                {
                    // Conta quante email sono più recenti dell'ultimo timestamp
                    int newEmailsCount = 0;
                    foreach (var uid in allUids)
                    {
                        var message = await inbox.GetMessageAsync(uid);
                        if (message.Date.DateTime > _lastProcessedTime)
                        {
                            newEmailsCount++;
                        }
                    }
                    
                    if (newEmailsCount > 0)
                    {
                        OnLogMessage?.Invoke($"📧 Trovate {newEmailsCount} nuove email da processare");
                    }
                }
                
                DateTime? latestEmailTime = null;
                int processedCount = 0;
                int skippedCount = 0;
                int unauthorizedCount = 0;
                
                foreach (var uid in allUids)
                {
                    var message = await inbox.GetMessageAsync(uid);
                    
                    // Controlla se l'email è più recente dell'ultimo timestamp processato
                    var emailTime = message.Date.DateTime;
                    
                    if (emailTime <= _lastProcessedTime)
                    {
                        skippedCount++;
                        // Non logghiamo ogni email saltata per evitare spam nel log
                        continue;
                    }
                    
                    OnLogMessage?.Invoke($"📧 Processando email del {emailTime:dd/MM/yyyy HH:mm:ss} (ricevuta dopo l'avvio - cutoff: {_lastProcessedTime:dd/MM/yyyy HH:mm:ss})");
                    
                    var success = await ProcessEmail(message);
                    if (success)
                    {
                        processedCount++;
                    }
                    else
                    {
                        unauthorizedCount++;
                    }
                    
                    // Tieni traccia dell'email più recente processata
                    if (latestEmailTime == null || emailTime > latestEmailTime)
                    {
                        latestEmailTime = emailTime;
                    }
                    
                    if (_deleteAfterPrint && success)
                    {
                        await inbox.AddFlagsAsync(uid, MessageFlags.Deleted, true);
                        OnLogMessage?.Invoke($"🗑️ Email eliminata dopo stampa da {message.From.FirstOrDefault()?.Name}");
                    }
                    else
                    {
                        await inbox.AddFlagsAsync(uid, MessageFlags.Seen, true);
                    }
                }

                // Aggiorna il timestamp dell'ultima email processata
                if (latestEmailTime.HasValue)
                {
                    SaveLastProcessedTime(latestEmailTime.Value);
                    OnLogMessage?.Invoke($"💾 Aggiornato timestamp ultima email processata: {latestEmailTime:dd/MM/yyyy HH:mm:ss}");
                }
                
                if (processedCount == 0)
                {
                    if (skippedCount > 0)
                    {
                        OnLogMessage?.Invoke($"📭 Nessuna nuova email da processare ({skippedCount} email saltate - già presenti all'avvio)");
                    }
                    else
                    {
                        OnLogMessage?.Invoke($"📭 Nessuna email trovata nell'inbox");
                    }
                }
                else
                {
                    OnLogMessage?.Invoke($"✅ Processate {processedCount} nuove email");
                }

                if (unauthorizedCount > 0)
                {
                    OnLogMessage?.Invoke($"🚫 Trovate {unauthorizedCount} nuove email da mittenti non sicuri, ignorate");
                }

                if (_deleteAfterPrint && processedCount > 0)
                {
                    await inbox.ExpungeAsync();
                }
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"❌ Errore processamento email: {ex.Message}");
                
                // Se la connessione IMAP è persa, prova a riconnettersi
                if (ex.Message.Contains("not connected") || ex.Message.Contains("disconnected"))
                {
                    OnLogMessage?.Invoke("🔄 Tentativo di riconnessione IMAP...");
                    try
                    {
                        if (_imapClient?.IsConnected == false)
                        {
                            await _imapClient.ConnectAsync(_emailServer, _emailPort, true);
                            await _imapClient.AuthenticateAsync(_emailUsername, _emailPassword);
                            OnLogMessage?.Invoke("✅ Riconnessione IMAP riuscita");
                        }
                    }
                    catch (Exception reconnectEx)
                    {
                        OnLogMessage?.Invoke($"❌ Riconnessione fallita: {reconnectEx.Message}");
                    }
                }
            }
        }

        private async Task<bool> ProcessEmail(MimeMessage message)
        {
            var sender = message.From.FirstOrDefault()?.Name ?? message.From.FirstOrDefault()?.ToString() ?? "Sconosciuto";
            var senderEmail = "";
            var fromAddress = message.From.FirstOrDefault();
            if (fromAddress is MailboxAddress mailboxAddress)
            {
                senderEmail = mailboxAddress.Address;
            }
            
            // Controllo mittente sicuro se configurato
            if (!string.IsNullOrEmpty(_secureSender))
            {
                var secureSenders = _secureSender.Split(';', StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .ToList();
                
                bool isAuthorized = secureSenders.Any(secureSender => 
                    senderEmail.Equals(secureSender, StringComparison.OrdinalIgnoreCase));
                
                if (!isAuthorized)
                {
                    OnLogMessage?.Invoke($"📧 Email da {sender} ({senderEmail}) ignorata - mittente non autorizzato");
                    return false;
                }
                OnLogMessage?.Invoke($"✅ Email da mittente sicuro: {sender} ({senderEmail})");
            }
            
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
                using (var document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import))
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
                // Prova prima SumatraPDF se disponibile
                var sumatraPath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "SumatraPDF.exe");
                if (File.Exists(sumatraPath))
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = sumatraPath,
                        Arguments = $"-print-to \"{printerName}\" \"{filePath}\"",
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        UseShellExecute = false
                    };

                    using (var process = Process.Start(psi))
                    {
                        process?.WaitForExit(30000);
                        if (process?.ExitCode == 0)
                        {
                            return; // Stampa riuscita
                        }
                    }
                }

                // Fallback: usa il metodo Windows Shell
                var shellPsi = new ProcessStartInfo
                {
                    FileName = filePath,
                    Verb = "print",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true
                };

                using (var process = Process.Start(shellPsi))
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

        public async Task ResetLastProcessedTime()
        {
            try
            {
                OnLogMessage?.Invoke("⏰ Reset timestamp richiesto - leggendo ultima email presente...");
                
                var tempClient = new ImapClient();
                await tempClient.ConnectAsync(_emailServer, _emailPort, true);
                await tempClient.AuthenticateAsync(_emailUsername, _emailPassword);
                
                var inbox = tempClient.Inbox;
                await inbox.OpenAsync(FolderAccess.ReadWrite);
                
                if (inbox.Count > 0)
                {
                    var lastMessage = await inbox.GetMessageAsync(inbox.Count - 1);
                    _lastProcessedTime = lastMessage.Date.DateTime;
                    SaveLastProcessedTime(_lastProcessedTime);
                    OnLogMessage?.Invoke($"⏰ Timestamp resettato all'ultima email presente: {_lastProcessedTime:dd/MM/yyyy HH:mm:ss}");
                }
                else
                {
                    _lastProcessedTime = DateTime.Now;
                    SaveLastProcessedTime(_lastProcessedTime);
                    OnLogMessage?.Invoke($"⏰ Inbox vuota - timestamp resettato a: {_lastProcessedTime:dd/MM/yyyy HH:mm:ss}");
                }
                
                await tempClient.DisconnectAsync(true);
                tempClient.Dispose();
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"❌ Errore nel reset timestamp: {ex.Message}");
                // Fallback: usa ora
                _lastProcessedTime = DateTime.Now;
                SaveLastProcessedTime(_lastProcessedTime);
                OnLogMessage?.Invoke($"⏰ Timestamp resettato a ora (fallback): {_lastProcessedTime:dd/MM/yyyy HH:mm:ss}");
            }
        }

        public DateTime GetLastProcessedTime()
        {
            return _lastProcessedTime;
        }

        private void CleanupResources()
        {
            try
            {
                if (_imapClient?.IsConnected == true)
                {
                    _imapClient.Disconnect(true);
                    OnLogMessage?.Invoke("📧 Disconnesso IMAP");
                }
                _imapClient?.Dispose();
                _imapClient = null;
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"⚠️ Errore disconnessione IMAP: {ex.Message}");
            }

            try
            {
                if (_smtpClient?.IsConnected == true)
                {
                    _smtpClient.Disconnect(true);
                    OnLogMessage?.Invoke("📤 Disconnesso SMTP");
                }
                _smtpClient?.Dispose();
                _smtpClient = null;
            }
            catch (Exception ex)
            {
                OnLogMessage?.Invoke($"⚠️ Errore disconnessione SMTP: {ex.Message}");
            }

            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
            
            OnLogMessage?.Invoke("✅ Cleanup completato");
        }

        public void Stop()
        {
            OnLogMessage?.Invoke("🛑 Arresto servizio richiesto...");
            
            _isRunning = false;
            
            // Cancella il token per interrompere il loop immediatamente
            _cancellationTokenSource?.Cancel();
            
            OnStatusChanged?.Invoke("Servizio fermo", false);
            OnLogMessage?.Invoke("🛑 Servizio fermato");
            
            // Il cleanup verrà fatto nel finally di StartAsync
        }
    }

    static class Program
    {
        private static Mutex _mutex = null;
        
        [STAThread]
        static void Main()
        {
            const string mutexName = "EmailPrintService_SingleInstance";
            bool createdNew;
            
            _mutex = new Mutex(true, mutexName, out createdNew);
            
            if (!createdNew)
            {
                // L'applicazione è già in esecuzione
                MessageBox.Show("Email Print Service è già in esecuzione!", 
                              "Applicazione già attiva", 
                              MessageBoxButtons.OK, 
                              MessageBoxIcon.Information);
                return;
            }
            
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                var form = new MainForm();
                Application.Run(form);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore avvio applicazione: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _mutex?.ReleaseMutex();
                _mutex?.Dispose();
            }
        }
    }
}