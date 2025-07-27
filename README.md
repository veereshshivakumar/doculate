# Doculate - Markdown to Word Converter

🚀 **A powerful VS Code extension for converting Markdown files to professional Word documents with Mermaid diagram support**

[![VS Code Extension](https://img.shields.io/badge/VS%20Code-Extension-blue)](https://code.visualstudio.com/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.5+-blue)](https://www.typescriptlang.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

## ⚠️ Prerequisites

**This extension requires the following dependencies:**

- **[Pandoc](https://pandoc.org/installing.html)** - Document conversion engine
- **[Node.js](https://nodejs.org/en/download/)** - Required for Mermaid diagram support
- **[Mermaid CLI](https://github.com/mermaid-js/mermaid-cli)** - Diagram rendering (installed via npm)

### Manual Installation Links

If you prefer to install dependencies manually before using the extension:

1. **Pandoc**: Download from [https://pandoc.org/installing.html](https://pandoc.org/installing.html)
   - Windows: Download the MSI installer
   - Mac: Use Homebrew (`brew install pandoc`) or download the PKG
   - Linux: Use your package manager (`sudo apt-get install pandoc`)

2. **Node.js**: Download from [https://nodejs.org/en/download/](https://nodejs.org/en/download/)
   - Choose the LTS version for stability
   - Installation includes npm package manager

3. **Mermaid CLI**: Install after Node.js using npm:
   ```bash
   npm install -g @mermaid-js/mermaid-cli
   ```

### Automatic Installation

**Don't worry!** The extension can automatically install these dependencies for you:
- After installing the extension, open the Doculate panel
- Click the installation buttons when prompted
- The extension will guide you through the setup process

## ✨ Features

- **📄 Markdown File Discovery**: Automatically discover and organize markdown files in your workspace
- **📝 Real-time Preview**: Live markdown preview with Mermaid diagram rendering
- **📊 Professional Word Export**: Export to .docx using Pandoc with custom reference templates
- **🎨 Mermaid Diagram Support**: Automatic conversion of Mermaid diagrams to images in exports
- **🔄 Cross-platform Installation**: Automated installation of dependencies (Pandoc, Node.js, Mermaid CLI)
- **⚙️ Template Management**: Store and manage Word reference documents for consistent styling
- **� VS Code Integration**: Native VS Code theming and seamless workflow integration

## 🚀 Quick Start

### Installation

1. **Install the Extension**:
   - Open VS Code
   - Go to Extensions (Ctrl+Shift+X)
   - Search for "Doculate"
   - Click Install

2. **Setup Dependencies**:
   - The extension will automatically detect missing dependencies
   - Click the install buttons for Pandoc, Node.js, and Mermaid CLI when prompted
   - Or install manually following the prompts

3. **Add Reference Templates** (Optional):
   - Open the extension panel (Command Palette: "Open Doculate Panel")
   - Go to Settings tab
   - Click "Add Reference Document" to add .docx templates
   - Your templates will be available for all exports

### Basic Usage

1. **Open a workspace** with markdown files
2. **Launch Doculate** from Command Palette or Explorer view
3. **Select a markdown file** to preview
4. **Export to Word** using the Export tab
5. **Choose your template** and output location

### Development Setup

```bash
# Clone the repository
git clone https://github.com/Veereshs88/doculate-ai.git
cd doculate-ai

# Install dependencies

```bash
# Clone the repository (replace with your actual GitHub URL)
git clone https://github.com/your-username/doculate-ai.git
cd doculate-ai

# Install dependencies
npm install

# Compile the extension
npm run compile

# Open in VS Code for development
code .

# Press F5 to launch Extension Development Host
```

## 📋 How It Works

### 1. Document Discovery
The extension automatically scans your workspace for markdown files and displays them in a simple list. Any markdown file can be previewed and exported to Word format.

### 2. Advanced Preview System
- **Syntax Highlighting**: Code blocks with language-specific coloring
- **Mermaid Diagrams**: Automatic rendering of flowcharts and diagrams
- **Real-time Updates**: Preview updates as you edit markdown files
- **Responsive Layout**: Optimized for different screen sizes

### 3. Professional Export
- **Word Documents**: Export to .docx with proper formatting and structure
- **Custom Templates**: Use your own Word templates for branded documents
- **Advanced Formatting**: Tables, code blocks, images, and complex layouts
- **Batch Processing**: Combine multiple files into a single document

## 🎯 Use Cases

### Technical Writers
- Convert markdown documentation to professional Word format
- Maintain consistent formatting across documents
- Export complex technical specifications with proper structure

### Product Managers
- Convert requirements documents to Word for stakeholder review
- Export user stories and acceptance criteria in professional format
- Create branded documentation with custom templates

### Software Architects
- Export system design documents with proper formatting
- Convert technical specifications to Word for team distribution
- Maintain professional documentation standards

### Teams & Organizations
- Standardize document formatting across projects
- Convert existing markdown documentation to Word
- Create branded documents with custom templates

## 🛠 Technologies

- **VS Code Extension API**: Native integration with VS Code
- **TypeScript**: Type-safe development
- **Marked**: Advanced markdown parsing and rendering
- **Docx**: Professional Word document generation
- **WebView**: Rich HTML interface within VS Code
- **Mermaid**: Diagram rendering and export

## 📊 Supported Document Types

### Any Markdown File
- **Technical Documentation**: API docs, system architecture, design decisions
- **Business Documents**: Requirements, specifications, project plans
- **Meeting Notes**: Agendas, minutes, action items
- **Process Documentation**: Procedures, guidelines, workflows
- **Academic Papers**: Research, reports, presentations
- **Blog Posts**: Articles, tutorials, guides

The extension works with **any markdown file** - no special naming or categorization required!

## 🔧 Configuration

Access settings via the ⚙️ icon in the extension panel:

### Export Settings
- **Default Output Directory**: Choose where exported files are saved
- **Word Templates**: Manage custom templates for consistent formatting
- **Formatting Options**: Configure default styles and preferences

### Preview Settings
- **Syntax Highlighting**: Enable/disable code block coloring
- **Mermaid Diagrams**: Configure diagram rendering options
- **Auto-refresh**: Control preview update frequency

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

### Development Workflow

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Run tests (`npm test`)
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🐛 Bug Reports & Feature Requests

- **Bug Reports**: [GitHub Issues](https://github.com/your-username/doculate-ai/issues)
- **Feature Requests**: [GitHub Discussions](https://github.com/your-username/doculate-ai/discussions)
- **Documentation**: [Wiki](https://github.com/your-username/doculate-ai/wiki)

## 🙏 Acknowledgments

- Built with [VS Code Extension API](https://code.visualstudio.com/api)
- Markdown processing by [Marked](https://marked.js.org/)
- Word document generation by [docx](https://docx.js.org/)
- Diagram rendering by [Mermaid](https://mermaid.js.org/)

---

**Made with ❤️ for the developer community**
