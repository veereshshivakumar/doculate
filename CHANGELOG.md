# Changelog

All notable changes to the Doculate extension will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Advanced formatting phase planning (Phase 5)
- Mermaid diagram support and batch processing
- Enhanced export capabilities design

## [0.1.0] - 2025-01-27

### Added
- Complete VS Code extension architecture
- Markdown file discovery and workspace scanning
- Live markdown preview with syntax highlighting
- Mermaid diagram rendering and conversion to images
- Professional Word document export using Pandoc
- Cross-platform dependency installation (Pandoc, Node.js, Mermaid CLI)
- Reference document template management system
- Automated dependency detection and installation prompts
- Clean, responsive webview interface with VS Code theming
- Template storage using VS Code's globalState API
- Debug tools for template storage inspection
- Comprehensive error handling and user feedback

### Features
- **Document Discovery**: Automatic markdown file detection in workspace
- **Live Preview**: Real-time markdown rendering with Mermaid diagram support
- **Professional Export**: Pandoc-powered Word document generation with custom templates
- **Template Management**: Store and manage .docx reference documents
- **Cross-platform Support**: Windows, macOS, and Linux compatibility
- **Dependency Management**: Automated installation of required tools
- **VS Code Integration**: Native theming and seamless workflow integration
- Professional Word document export functionality
- Template support for Word export
- VS Code/Cursor-style preview interface
- Advanced settings modal with export configuration
- Professional export workflow
- Export modal with file browsing capabilities
- Progress tracking for export operations
- Professional VS Code theming throughout
- Complete project documentation (README, ROADMAP, CONTRIBUTING)
- MIT license

### Features Implemented
- **File Management**: Automatic markdown file discovery and categorization (TDD, BRD, SDD, Custom)
- **Preview System**: Real-time markdown rendering with syntax highlighting
- **Export System**: Full Word document export with custom template support
- **Preview Interface**: Professional markdown preview UI
- **Settings System**: Comprehensive configuration management
- **Error Handling**: Robust error handling and user feedback

### Technical Implementation
- TypeScript-based VS Code extension
- WebView-based user interface
- Native VS Code file system integration
- `marked` library for markdown processing
- `docx` library for Word document generation
- CommonJS module system for VS Code compatibility

### Documentation
- Comprehensive README with features and setup instructions
- Development roadmap with phase-by-phase breakdown
- Contribution guidelines for open source development
- MIT license for open source distribution

## Development Phases Completed

### ✅ Phase 1: Foundation & Core Plugin
- VS Code extension project structure
- Basic WebView panel implementation
- Extension activation and command registration
- VS Code dark theme styling

### ✅ Phase 2: File Management & Preview
- Workspace markdown file scanning
- File selection interface with dropdown
- Real-time markdown preview
- File metadata extraction and display

### ✅ Phase 3: Export Functionality
- Word document export implementation
- Template selection and management
- Export progress tracking
- Professional export modal interface

### ✅ Phase 4: Advanced Preview & Export
- Enhanced markdown preview with syntax highlighting
- Professional Word export with template support
- Advanced formatting and batch processing
- Export workflow and progress tracking

## Upcoming Releases

### [0.2.0] - Planned
- Mermaid diagram support and rendering
- Advanced table formatting in Word export
- Batch processing for multiple files

### [0.3.0] - Planned
- PDF export capability
- HTML export with custom styling
- Performance optimizations and polish

### [1.0.0] - Target Release
- Production-ready extension
- VS Code Marketplace publication
- Complete formatting and export features
- Comprehensive testing and documentation 