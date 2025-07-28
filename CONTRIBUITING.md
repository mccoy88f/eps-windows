# Contributing to Email Print Service

Thank you for your interest in contributing to Email Print Service! This document provides guidelines and information for contributors.

## 🤝 How to Contribute

### Reporting Bugs
1. **Check existing issues** first to avoid duplicates
2. **Use the bug report template** when creating new issues
3. **Provide detailed information**:
   - Operating system and version
   - .NET version (if applicable)
   - Steps to reproduce the issue
   - Expected vs actual behavior
   - Screenshots or error messages

### Suggesting Features
1. **Check if the feature already exists** in issues or discussions
2. **Use the feature request template**
3. **Explain the use case** and why it would be valuable
4. **Consider implementation complexity** and alternatives

### Code Contributions

#### Development Environment Setup
```bash
# Prerequisites
# - Visual Studio 2019/2022 or VS Code
# - .NET 6 SDK
# - Git

# Clone the repository
git clone https://github.com/mccoy88f/email-print-service.git
cd email-print-service

# Restore dependencies
dotnet restore

# Build the project
dotnet build

# Run tests
dotnet test
```

#### Development Workflow
1. **Fork the repository** to your GitHub account
2. **Create a feature branch** from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. **Make your changes** following the coding guidelines
4. **Write tests** for new functionality
5. **Test thoroughly** on Windows environments
6. **Commit with clear messages**:
   ```bash
   git commit -m "feat: add support for OAuth2 authentication"
   ```
7. **Push to your fork**:
   ```bash
   git push origin feature/your-feature-name
   ```
8. **Create a Pull Request** with detailed description

#### Commit Message Guidelines
We follow the [Conventional Commits](https://www.conventionalcommits.org/) specification:

- `feat:` new features
- `fix:` bug fixes
- `docs:` documentation changes
- `style:` formatting, missing semicolons, etc.
- `refactor:` code changes that neither fix bugs nor add features
- `test:` adding or updating tests
- `chore:` updating build tasks, package manager configs, etc.

Examples:
```
feat: add support for multi-tenant configuration
fix: resolve PDF printing issue with SumatraPDF
docs: update README with OAuth2 setup instructions
```

## 🎨 Coding Guidelines

### C# Style Guidelines
- Follow **Microsoft C# Coding Conventions**
- Use **meaningful variable and method names**
- Add **XML documentation** for public APIs
- Keep methods **small and focused**
- Use **async/await** for I/O operations
- Handle **exceptions appropriately**

### Code Organization
```
EmailPrintService/
├── Core/                   # Core business logic
├── Services/               # Email, print, and other services
├── UI/                     # Windows Forms UI components
├── Models/                 # Data models and DTOs
├── Configuration/          # Settings and configuration
├── Utils/                  # Utility classes
└── Tests/                  # Unit and integration tests
```

### Naming Conventions
- **Classes**: PascalCase (`EmailService`)
- **Methods**: PascalCase (`ProcessEmail`)
- **Properties**: PascalCase (`PrinterName`)
- **Fields**: camelCase with underscore (`_emailClient`)
- **Constants**: PascalCase (`DefaultTimeout`)
- **Enums**: PascalCase (`PrintMethod`)

### Error Handling
```csharp
// Good: Specific exception handling
try
{
    await ProcessEmailAsync(message);
}
catch (EmailConnectionException ex)
{
    _logger.LogError(ex, "Failed to connect to email server");
    // Handle specific email connection issues
}
catch (PrinterException ex)
{
    _logger.LogError(ex, "Printer error occurred");
    // Handle printer-specific issues
}

// Bad: Generic catch-all
try
{
    await ProcessEmailAsync(message);
}
catch (Exception ex)
{
    // Generic handling
}
```

### Async/Await Best Practices
```csharp
// Good: Proper async method
public async Task<bool> ProcessEmailAsync(MimeMessage message)
{
    var attachments = await ExtractPdfAttachmentsAsync(message);
    return await PrintAttachmentsAsync(attachments);
}

// Bad: Blocking async calls
public bool ProcessEmail(MimeMessage message)
{
    var attachments = ExtractPdfAttachmentsAsync(message).Result; // Don't do this
    return PrintAttachmentsAsync(attachments).Result;
}
```

## 🧪 Testing Guidelines

### Unit Tests
- Use **xUnit** testing framework
- **Mock external dependencies** (email servers, printers)
- Test **both success and failure scenarios**
- Use **descriptive test names**

```csharp
[Fact]
public async Task ProcessEmail_WithValidPdfAttachment_ShouldReturnTrue()
{
    // Arrange
    var emailService = CreateMockEmailService();
    var message = CreateTestMessageWithPdf();
    
    // Act
    var result = await emailService.ProcessEmailAsync(message);
    
    // Assert
    Assert.True(result);
}
```

### Integration Tests
- Test **complete workflows** from email to print
- Use **test email accounts** (not production)
- Verify **file system operations**
- Test **configuration loading/saving**

### Manual Testing
Before submitting PR, test:
- [ ] Email connection with different providers
- [ ] PDF printing with various file types
- [ ] GUI responsiveness and error handling
- [ ] System tray functionality
- [ ] Configuration persistence
- [ ] Windows startup integration

## 📚 Documentation

### Code Documentation
```csharp
/// <summary>
/// Processes an email message and extracts PDF attachments for printing.
/// </summary>
/// <param name="message">The MIME message to process</param>
/// <returns>True if all PDFs were processed successfully</returns>
/// <exception cref="EmailProcessingException">Thrown when email processing fails</exception>
public async Task<bool> ProcessEmailAsync(MimeMessage message)
{
    // Implementation
}
```

### README Updates
- Update **feature lists** when adding new functionality
- Add **configuration examples** for new settings
- Include **screenshots** for UI changes
- Update **troubleshooting** section for known issues

## 🔍 Code Review Process

### What We Look For
- **Functionality**: Does the code work as intended?
- **Security**: Are there any security vulnerabilities?
- **Performance**: Is the code efficient?
- **Maintainability**: Is the code easy to understand and modify?
- **Testing**: Are there adequate tests?
- **Documentation**: Is the code well-documented?

### Review Checklist
- [ ] Code follows style guidelines
- [ ] All tests pass
- [ ] No security vulnerabilities introduced
- [ ] Performance impact considered
- [ ] Documentation updated if needed
- [ ] Backward compatibility maintained
- [ ] Error handling is appropriate

## 🏷️ Versioning and Releases

We follow [Semantic Versioning](https://semver.org/):
- **MAJOR** version for incompatible API changes
- **MINOR** version for backward-compatible functionality additions
- **PATCH** version for backward-compatible bug fixes

## 🎯 Areas for Contribution

### High Priority
- [ ] OAuth2 authentication for Gmail/Outlook
- [ ] Linux compatibility (Mono/Avalonia UI)
- [ ] Docker container support
- [ ] Web-based configuration interface
- [ ] Advanced PDF processing options

### Medium Priority
- [ ] Print queue management
- [ ] Statistics and reporting
- [ ] Email filtering rules
- [ ] Multiple printer support
- [ ] Scheduled printing

### Low Priority
- [ ] Mobile companion app
- [ ] Cloud service integration
- [ ] Enterprise SSO integration
- [ ] Advanced logging/monitoring

## 🆘 Getting Help

- **GitHub Discussions**: For general questions and feature discussions
- **GitHub Issues**: For bug reports and specific problems
- **Discord**: Real-time chat with other contributors
- **Email**: Contact the maintainer for private discussions

## 📜 License

By contributing to Email Print Service, you agree that your contributions will be licensed under the MIT License.

## 🙏 Recognition

Contributors are recognized in:
- **README.md** contributors section
- **Release notes** for significant contributions
- **About dialog** in the application (for major contributors)

Thank you for helping make Email Print Service better! 🚀