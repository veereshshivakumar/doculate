# Contributing to Doculate

🎉 Thank you for your interest in contributing to the Doculate VS Code extension! We welcome contributions from the community to help improve this markdown to Word conversion tool.

## 🚀 Getting Started

### Prerequisites

- **Node.js** (v18 or higher)
- **VS Code** (latest version)
- **Git**
- **Pandoc** (for testing export functionality)

### Development Setup

1. **Fork and Clone**
   ```bash
   git clone https://github.com/Veereshs88/doculate-ai.git
   cd doculate-ai
   ```

2. **Install Dependencies**
   ```bash
   npm install
   ```

3. **Compile the Extension**
   ```bash
   npm run compile
   ```

4. **Open in VS Code**
   ```bash
   code .
   ```

5. **Launch Extension Development Host**
   - Press `F5` in VS Code
   - This opens a new window with the extension loaded

## 🛠 Development Workflow

### Project Structure

```
doculate-ai/
├── extension.ts              # VS Code extension entry point
├── package.json              # Extension manifest
├── tsconfig.extension.json   # TypeScript config for extension
├── README.md                 # Project documentation
├── ROADMAP.md               # Development roadmap
├── CONTRIBUTING.md          # Contribution guidelines
└── LICENSE                  # MIT license
```

### Making Changes

1. **Create a Feature Branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make Your Changes**
   - Extension logic goes in `extension.ts`
   - All UI is embedded within the extension's webview HTML
   - Update tests if applicable

3. **Test Your Changes**
   ```bash
   npm run compile
   # Press F5 to test in Extension Development Host
   ```

4. **Commit Your Changes**
   ```bash
   git add .
   git commit -m "feat: add your feature description"
   ```

## 📝 Code Guidelines

### TypeScript Standards

- Use **TypeScript** for all code
- Follow **strict mode** settings
- Add **type annotations** for public APIs
- Use **async/await** instead of Promises where possible

### Code Style

- Use **2 spaces** for indentation
- Use **single quotes** for strings
- Add **JSDoc comments** for public functions
- Follow **VS Code API patterns**

```typescript
/**
 * Exports a markdown file to Word format
 * @param filePath Path to the markdown file
 * @param templatePath Optional Word template path
 * @returns Promise resolving to export result
 */
async function exportToWord(filePath: string, templatePath?: string): Promise<ExportResult> {
  // Implementation
}
```

### VS Code Extension Best Practices

- **Lazy load** heavy operations
- Use **progress notifications** for long operations
- Handle **errors gracefully** with user-friendly messages
- Follow **VS Code UX patterns**
- Test on **multiple platforms**

## 🧪 Testing

### Manual Testing

1. **Basic Functionality**
   - Extension loads without errors
   - File discovery works in various workspace types
   - Preview renders markdown correctly
   - Export functionality works

2. **Edge Cases**
   - Empty workspaces
   - Large numbers of files
   - Corrupted markdown files
   - Network failures (for dependency installation)

3. **Cross-Platform**
   - Test on Windows, macOS, and Linux
   - Verify file path handling
   - Check keyboard shortcuts

### Testing Checklist

- [ ] Extension activates properly
- [ ] File discovery works
- [ ] Preview renders correctly
- [ ] Export functionality works
- [ ] Settings modal functions
- [ ] Error handling works
- [ ] No console errors
- [ ] Memory usage is reasonable

## 🐛 Bug Reports

When reporting bugs, please include:

- **VS Code Version**
- **Extension Version**
- **Operating System**
- **Steps to Reproduce**
- **Expected vs Actual Behavior**
- **Console Errors** (if any)
- **Workspace Structure** (if relevant)

## 💡 Feature Requests

Before submitting feature requests:

1. Check existing [GitHub Issues](https://github.com/your-username/doculate-ai/issues)
2. Review the [Roadmap](ROADMAP.md)
3. Consider if it fits the extension's scope
4. Provide clear use cases and examples

## 📋 Pull Request Process

### Before Submitting

- [ ] Code compiles without errors
- [ ] Manual testing completed
- [ ] Documentation updated (if needed)
- [ ] Commit messages follow conventional format

### Commit Message Format

Use [Conventional Commits](https://www.conventionalcommits.org/):

```
feat: add export to PDF functionality
fix: resolve file path issues on Windows
docs: update installation instructions
refactor: improve error handling in export
```

### Pull Request Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] Manual testing completed
- [ ] Cross-platform testing (if applicable)
- [ ] Edge cases considered

## Screenshots (if applicable)
Add screenshots for UI changes
```

## 🎯 Development Priorities

Current focus areas (see [Roadmap](ROADMAP.md)):

1. **Advanced Formatting** - Mermaid diagrams, complex tables
2. **Batch Processing** - Multi-file export capabilities
3. **Performance Optimization** - Speed and memory improvements
4. **Performance** - Optimization and caching

## 📚 Resources

### VS Code Extension Development

- [VS Code Extension API](https://code.visualstudio.com/api)
- [Extension Guidelines](https://code.visualstudio.com/api/references/extension-guidelines)
- [Publishing Extensions](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)

### Project-Specific

- [Project Roadmap](ROADMAP.md)
- [Architecture Overview](https://github.com/your-username/doculate-ai/wiki/Architecture)
- [API Documentation](https://github.com/your-username/doculate-ai/wiki/API)

## 🤝 Community

- **GitHub Discussions**: General questions and ideas
- **GitHub Issues**: Bug reports and feature requests
- **Pull Requests**: Code contributions

## 📄 License

By contributing to this project, you agree that your contributions will be licensed under the MIT License.

---

**Thank you for contributing to Doculate!** 🚀 