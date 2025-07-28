# 📧→🖨️ Email Print Service

**Professional Windows service for automatic PDF printing from email attachments**

[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-6.0+-blue.svg)](https://dotnet.microsoft.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Release](https://img.shields.io/badge/Release-v1.0.0-brightgreen.svg)](https://github.com/mccoy88f/email-print-service/releases)

> **Transform your email-to-print workflow!** Automatically receive emails with PDF attachments and print them silently on your network printer. Perfect for offices, remote printing, and automated document processing.

---

## 🌟 **Key Features**

### 📧 **Smart Email Processing**
- ✅ **Multi-provider support** (Gmail, Outlook, Yahoo, Exchange)
- ✅ **IMAP with SSL/TLS** secure connections
- ✅ **Multi-PDF handling** - process multiple attachments per email
- ✅ **Automatic email cleanup** (optional)
- ✅ **Email confirmation** with print status

### 🖨️ **Advanced Print Management**
- ✅ **Multiple print methods**: .NET PrintDocument, SumatraPDF, Windows Shell
- ✅ **3 confirmation modes**: Automatic, Timed popup, Manual approval
- ✅ **Detailed PDF selection** - choose which files to print
- ✅ **Full printer configuration** via Windows standard dialogs
- ✅ **Duplex, paper size, orientation** control

### 🎛️ **Professional GUI**
- ✅ **Complete Windows Forms interface** with real-time logging
- ✅ **System tray integration** with context menu
- ✅ **Manual email checking** - force immediate processing
- ✅ **Configurable intervals** (5-300 seconds)
- ✅ **Service restart** and monitoring controls

### 🔒 **Enterprise Ready**
- ✅ **No admin rights required** - runs in user space
- ✅ **Portable deployment** - single executable + optional SumatraPDF
- ✅ **Auto-start with Windows** (per-user registry)
- ✅ **Comprehensive logging** and error handling
- ✅ **Self-contained** - includes all dependencies

---

## 📸 **Screenshots**

### Main Configuration Interface
![Main Interface](docs/screenshots/main-interface.png)

### Multi-PDF Selection Dialog
![PDF Selection](docs/screenshots/pdf-selection-dialog.png)

### System Tray Integration
![System Tray](docs/screenshots/system-tray.png)

---

## 🚀 **Quick Start**

### **For End Users (Pre-built Release):**
1. **Download** the latest release from [Releases](https://github.com/mccoy88f/email-print-service/releases)
2. **Extract** to any folder (e.g., `C:\EmailPrintService\`)
3. **Optional**: Add `SumatraPDF.exe` to the same folder for best PDF compatibility
4. **Run** `EmailPrintService.exe`
5. **Configure** email and printer settings
6. **Start service** and enjoy automated printing!

### **For Developers:**
```bash
git clone https://github.com/mccoy88f/eps-windows.git
cd email-print-service
dotnet build -c Release
```

---

## 🛠️ **Development Setup**

### **Prerequisites**
- **Visual Studio 2019/2022** or **VS Code** with C# extension
- **.NET 6 SDK** ([Download](https://dotnet.microsoft.com/download/dotnet/6.0))
- **Windows 7/10/11** (development and runtime)

### **Clone and Build**
```bash
# Clone repository
git clone https://github.com/mccoy88f/eps-windows.git
cd email-print-service

# Restore NuGet packages
dotnet restore

# Build in Debug mode
dotnet build

# Build in Release mode
dotnet build -c Release

# Run from source
dotnet run
```

### **NuGet Dependencies**
```xml
<PackageReference Include="MailKit" Version="4.3.0" />
<PackageReference Include="MimeKit" Version="4.3.0" />
<PackageReference Include="PdfSharp" Version="1.50.5147" />
<PackageReference Include="System.Text.Json" Version="6.0.0" />
<PackageReference Include="System.Drawing.Common" Version="6.0.0" />
<PackageReference Include="PdfSharp.Gdi" Version="1.50.5147" />
```

---

## 📦 **Deployment**

### **Self-Contained Deployment (Recommended)**
```bash
# Create portable executable (includes .NET runtime)
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

# Output: bin/Release/net6.0-windows/win-x64/publish/EmailPrintService.exe (~70MB)
```

### **Framework-Dependent Deployment**
```bash
# Smaller executable (requires .NET 6 runtime on target machine)
dotnet publish -c Release -r win-x64 --self-contained false

# Output: Multiple files (~5MB total)
```

### **Enterprise Deployment Package**
```
📁 EmailPrintService_v1.0.0/
├── 📄 EmailPrintService.exe          # Main application
├── 📄 SumatraPDF.exe                 # Optional PDF renderer
├── 📄 README.txt                     # Quick setup guide
├── 📄 LICENSE.txt                    # MIT License
└── 📁 docs/                          # Documentation
    ├── 📄 Configuration-Guide.md     # Detailed setup
    ├── 📄 Troubleshooting.md         # Common issues
    └── 📄 API-Documentation.md       # For integrations
```

---

## ⚙️ **Configuration Guide**

### **1. Email Setup**
Configure your dedicated email account:

#### **Gmail (Recommended)**
```
Server: imap.gmail.com
Port: 993
Security: SSL/TLS
Username: your-printer@gmail.com
Password: [App-specific password]
```

**Setup Steps:**
1. Enable 2-factor authentication on Gmail
2. Generate an "App Password" in Google Account Security
3. Use the app password in EmailPrintService

#### **Outlook/Office 365**
```
Server: outlook.office365.com
Port: 993
Security: SSL/TLS
Username: printer@yourcompany.com
Password: [Your regular password]
```

#### **Exchange Server**
```
Server: mail.yourcompany.com
Port: 993 (or as configured)
Security: SSL/TLS
Username: DOMAIN\username or username@domain.com
Password: [Your domain password]
```

### **2. Print Configuration**
1. **Select printer** from dropdown (must be installed in Windows)
2. **Click "Configure Print"** to open Windows print dialog
3. **Set paper size, duplex, orientation** as needed
4. **Choose print method**: Automatic, .NET, SumatraPDF, or Windows Shell
5. **Select confirmation mode**: Automatic, Timed, or Manual

### **3. SumatraPDF Setup (Optional)**
For maximum PDF compatibility:
1. **Download** SumatraPDF portable from [official site](https://www.sumatrapdfreader.org/)
2. **Rename** to `SumatraPDF.exe`
3. **Copy** to same folder as `EmailPrintService.exe`
4. **Restart** the service

---

## 🎯 **Use Cases**

### **🏢 Office Environments**
- **Remote printing** for employees working from home
- **Client document processing** - customers email contracts/forms
- **Automated invoice printing** from accounting systems
- **Document scanning workflows** - scan to email, auto-print elsewhere

### **🏪 Retail/Service**
- **Customer order printing** from online systems  
- **Ticket/receipt printing** from web applications
- **Remote location printing** via email bridge

### **🏥 Healthcare/Legal**
- **Secure document handling** with manual approval
- **Patient form processing** with size verification
- **Legal document review** before printing

### **🏭 Manufacturing/Logistics**
- **Work order printing** from MES systems
- **Shipping label generation** from WMS
- **Quality control document** distribution

---

## 🔧 **API Integration**

### **Email Format**
Send emails to your configured address with PDF attachments:

```
To: printer@yourcompany.com
Subject: [Optional] Document description
Attachments: document1.pdf, document2.pdf, ...
```

### **Email Confirmation Response**
When enabled, you'll receive confirmation emails:

```
Subject: Print Confirmation - Your original subject

Dear Sender,

Your print request has been processed successfully.

Details:
• Date: 28/07/2025 14:30
• Files printed: document1.pdf, document2.pdf  
• Total documents: 2

Best regards,
Automatic Print System
```

### **Programmatic Integration**
```csharp
// Example: Send print job via SMTP
var message = new MimeMessage();
message.From.Add(new MailboxAddress("System", "system@company.com"));
message.To.Add(new MailboxAddress("Printer", "printer@company.com"));
message.Subject = "Invoice Batch #1234";

var attachment = new MimePart("application", "pdf")
{
    Content = new MimeContent(File.OpenRead("invoice.pdf")),
    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
    FileName = "invoice_1234.pdf"
};

var multipart = new Multipart("mixed");
multipart.Add(new TextPart("plain") { Text = "Please print attached invoice." });
multipart.Add(attachment);
message.Body = multipart;

// Send via your SMTP client
await smtpClient.SendAsync(message);
```

---

## 🛡️ **Security Considerations**

### **Email Security**
- ✅ **SSL/TLS encryption** for all email connections
- ✅ **App-specific passwords** recommended over main passwords
- ✅ **Dedicated email account** for print service only
- ✅ **Auto-cleanup** of processed emails (optional)

### **System Security**
- ✅ **No admin privileges** required
- ✅ **User-space execution** only
- ✅ **Local file storage** in user directory
- ✅ **PDF validation** before processing
- ✅ **Controlled network access** (IMAP/SMTP only)

### **Enterprise Deployment**
- 📋 **Group Policy** can control auto-start behavior
- 📋 **Firewall rules** can restrict email server access
- 📋 **User permissions** control printer access
- 📋 **Audit logging** available for compliance

---

## 🐛 **Troubleshooting**

### **Common Issues**

#### **"Cannot connect to email server"**
```bash
# Check network connectivity
ping imap.gmail.com

# Verify port access
telnet imap.gmail.com 993
```
- ✅ Verify server address and port
- ✅ Check firewall/antivirus blocking
- ✅ Confirm credentials are correct
- ✅ For Gmail: ensure app password is used

#### **"Printer not found"**
- ✅ Verify printer is installed in Windows
- ✅ Test manual printing from another application
- ✅ Check printer is online and accessible
- ✅ Ensure printer drivers are up to date

#### **"PDF files not printing correctly"**
- ✅ Try different print methods (.NET vs SumatraPDF vs Windows Shell)
- ✅ Verify PDF files open correctly in other applications
- ✅ Check printer supports PDF direct printing
- ✅ Install SumatraPDF for better compatibility

#### **"Application won't start"**
- ✅ Verify .NET 6 runtime is installed (self-contained versions don't need this)
- ✅ Check Windows Event Viewer for error details
- ✅ Run from command line to see error messages
- ✅ Ensure all required DLLs are present

### **Debug Mode**
```bash
# Run with verbose logging
EmailPrintService.exe --debug

# Check log files
# Location: %LOCALAPPDATA%\EmailPrintService\
dir %LOCALAPPDATA%\EmailPrintService\
```

### **Performance Tuning**
```json
// Optimize for high-volume environments
{
  "CheckInterval": 30,        // Reduce email checking frequency
  "ConfirmationMode": 0,      // Use automatic mode for speed
  "DeleteAfterPrint": true,   // Keep mailbox clean
  "PreferredMethod": 1        // Use SumatraPDF for reliability
}
```

---

## 🤝 **Contributing**

We welcome contributions! Here's how to get started:

### **Development Workflow**
1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Commit** your changes (`git commit -m 'Add amazing feature'`)
4. **Push** to the branch (`git push origin feature/amazing-feature`)
5. **Open** a Pull Request

### **Code Style**
- Follow **C# coding conventions**
- Use **meaningful variable names**
- Add **XML documentation** for public methods
- Include **unit tests** for new features
- Maintain **backward compatibility**

### **Testing**
```bash
# Run unit tests
dotnet test

# Run integration tests (requires email account setup)
dotnet test --configuration Integration

# Test different email providers
dotnet test --filter "Category=EmailProviders"
```

### **Feature Requests**
- 📧 **Email templates** for different document types
- 🔐 **OAuth2 authentication** for Gmail/Outlook
- 📊 **Dashboard** for print statistics
- 🌐 **Web interface** for remote configuration
- 📱 **Mobile app** for print queue management

---

## 📄 **License**

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

### **Third-Party Licenses**
- **MailKit/MimeKit**: MIT License
- **PdfSharp**: MIT License  
- **SumatraPDF**: GPLv3 (optional component)

---

## 👨‍💻 **Credits & Author**

### **Created by:**
**Antonello Migliorelli**  
- 🐙 GitHub: [@mccoy88f](https://github.com/mccoy88f)
- 📧 Email: [Contact via GitHub](https://github.com/mccoy88f)

### **Special Thanks**
- **MailKit Team** - Excellent IMAP/SMTP library
- **PdfSharp Team** - Reliable PDF processing
- **SumatraPDF Team** - Lightweight PDF rendering
- **Microsoft** - .NET Framework and development tools

---

## 🌟 **Star History**

If this project helped you, please consider giving it a ⭐!

[![Star History Chart](https://api.star-history.com/svg?repos=mccoy88f/email-print-service&type=Date)](https://star-history.com/#mccoy88f/email-print-service&Date)

---

## 📧 **Support & Contact**

- 🐛 **Bug Reports**: [GitHub Issues](https://github.com/mccoy88f/email-print-service/issues)
- 💡 **Feature Requests**: [GitHub Discussions](https://github.com/mccoy88f/email-print-service/discussions)
- 📖 **Documentation**: [Wiki](https://github.com/mccoy88f/email-print-service/wiki)
- 💬 **Community**: [Discord Server](https://discord.gg/email-print-service)

---

## 🚀 **Roadmap**

### **v1.1.0** (Next Release)
- [ ] OAuth2 support for Gmail/Outlook
- [ ] Web-based configuration interface
- [ ] Docker container support
- [ ] Linux compatibility (using Mono)

### **v1.2.0** (Future)
- [ ] Print queue management
- [ ] Advanced filtering rules
- [ ] Statistics dashboard
- [ ] Mobile app companion

### **v2.0.0** (Long-term)
- [ ] Multi-tenant support
- [ ] Enterprise SSO integration
- [ ] Advanced document processing
- [ ] Cloud service integration

---

<div align="center">

**Made with ❤️ for the automation community**

[![GitHub](https://img.shields.io/badge/GitHub-mccoy88f-blue?logo=github)](https://github.com/mccoy88f)
[![.NET](https://img.shields.io/badge/.NET-6.0+-purple?logo=.net)](https://dotnet.microsoft.com/)
[![Windows](https://img.shields.io/badge/Windows-Compatible-success?logo=windows)](https://www.microsoft.com/windows/)

**⭐ Star this repo if it's useful! ⭐**

</div>
