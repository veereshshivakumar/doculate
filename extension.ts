import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { exec } from 'child_process';
import { promisify } from 'util';
const { marked } = require('marked');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ImageRun } = require('docx');

export function activate(context: vscode.ExtensionContext) {
    const provider = new DoculateProvider(context.extensionUri);
    
    context.subscriptions.push(
        vscode.window.registerWebviewViewProvider('doculateView', provider),
        vscode.commands.registerCommand('doculate.openPanel', () => {
            DoculatePanel.createOrShow(context.extensionUri, context);
        })
    );
}

class DoculateProvider implements vscode.WebviewViewProvider {
    public static readonly viewType = 'doculateView';
    private _view?: vscode.WebviewView;

    constructor(private readonly _extensionUri: vscode.Uri) {}

    public resolveWebviewView(webviewView: vscode.WebviewView) {
        this._view = webviewView;

        webviewView.webview.options = {
            enableScripts: true,
            localResourceRoots: [this._extensionUri]
        };

        webviewView.webview.html = this._getSidebarHtml();
        webviewView.webview.onDidReceiveMessage(async (data) => {
            if (data.type === 'openPanel') {
                DoculatePanel.createOrShow(this._extensionUri, undefined as any);
            }
        });
    }

    private _getSidebarHtml() {
        return `<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Doculate</title>
            <style>
                body {
                    font-family: var(--vscode-font-family);
                    padding: 20px;
                    color: var(--vscode-foreground);
                    background: var(--vscode-editor-background);
                }
                .header {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    margin-bottom: 20px;
                }
                .button {
                    background: var(--vscode-button-background);
                    color: var(--vscode-button-foreground);
                    border: none;
                    padding: 8px 16px;
                    border-radius: 4px;
                    cursor: pointer;
                    width: 100%;
                }
                .button:hover {
                    background: var(--vscode-button-hoverBackground);
                }
            </style>
        </head>
        <body>
            <div class="header">
                <div>📄</div>
                <h3>Doculate</h3>
            </div>
            <button class="button" onclick="openPanel()">
                📄 Open Document Panel
            </button>
            <script>
                const vscode = acquireVsCodeApi();
                function openPanel() {
                    vscode.postMessage({ type: 'openPanel' });
                }
            </script>
        </body>
        </html>`;
    }
}

class DoculatePanel {
    public static currentPanel: DoculatePanel | undefined;
    public static readonly viewType = 'doculate';

    private readonly _panel: vscode.WebviewPanel;
    private readonly _extensionUri: vscode.Uri;
    private readonly _context: vscode.ExtensionContext;
    private _disposables: vscode.Disposable[] = [];
    private _mermaidCache: Map<string, string> = new Map(); // Cache for Mermaid base64 images

    public static createOrShow(extensionUri: vscode.Uri, context: vscode.ExtensionContext) {
        const column = vscode.window.activeTextEditor?.viewColumn;

        if (DoculatePanel.currentPanel) {
            DoculatePanel.currentPanel._panel.reveal(column);
            return;
        }

        const panel = vscode.window.createWebviewPanel(
            DoculatePanel.viewType,
            'Doculate - Markdown to Word',
            column || vscode.ViewColumn.One,
            {
                enableScripts: true,
                localResourceRoots: [extensionUri]
            }
        );

        DoculatePanel.currentPanel = new DoculatePanel(panel, extensionUri, context);
    }

    private constructor(panel: vscode.WebviewPanel, extensionUri: vscode.Uri, context: vscode.ExtensionContext) {
        this._panel = panel;
        this._extensionUri = extensionUri;
        this._context = context;

        this._update();
        this._panel.onDidDispose(() => this.dispose(), null, this._disposables);
        this._panel.webview.onDidReceiveMessage(this._handleMessage.bind(this), null, this._disposables);
    }

    private async _handleMessage(message: any) {
        switch (message.command) {
            case 'getMarkdownFiles':
                await this._getMarkdownFiles();
                break;
            case 'previewFile':
                await this._previewFile(message.filePath);
                break;
            case 'browseOutputFolder':
                await this._browseOutputFolder();
                break;
            case 'installPandoc':
                await this._handlePandocInstallation();
                break;
            case 'checkPandocStatus':
                await this._checkAndUpdatePandocStatus();
                break;
            case 'installNodejs':
                await this._handleNodejsInstallation();
                break;
            case 'installMermaid':
                await this._handleMermaidInstallation();
                break;
            case 'checkMermaidStatus':
                await this._checkAndUpdateMermaidStatus();
                break;
            case 'addTemplate':
                await this._addTemplate();
                break;
            case 'removeTemplate':
                await this._removeTemplate(message.templateName);
                break;
            case 'getWorkspacePath':
                await this._getWorkspacePath();
                break;
            case 'exportToWord':
                await this._exportToWord(message.data);
                break;
            case 'showTemplateData':
                await this._showTemplateData();
                break;
            case 'showError':
                vscode.window.showErrorMessage(message.message);
                break;
            case 'showInfo':
                vscode.window.showInformationMessage(message.message);
                break;
        }
    }

    private async _getMarkdownFiles() {
        try {
            if (!vscode.workspace.workspaceFolders) {
                vscode.window.showErrorMessage('No workspace folder is open');
                return;
            }

            const files = await vscode.workspace.findFiles('**/*.md', '**/node_modules/**');
            const fileInfos = files.map(file => ({
                name: path.basename(file.fsPath),
                path: file.fsPath,
                relativePath: vscode.workspace.asRelativePath(file)
            }));

            this._panel.webview.postMessage({
                command: 'updateFiles',
                files: fileInfos
            });
        } catch (error) {
            vscode.window.showErrorMessage(`Error finding markdown files: ${error}`);
        }
    }

    private async _previewFile(filePath: string) {
        try {
            const content = fs.readFileSync(filePath, 'utf8');
            
            // Count Mermaid diagrams
            const mermaidMatches = content.match(/```mermaid[\s\S]*?```/g);
            const mermaidCount = mermaidMatches ? mermaidMatches.length : 0;
            
            // First, send the basic markdown preview immediately for fast response
            const basicHtml = marked(content);
            this._panel.webview.postMessage({
                command: 'updatePreview',
                content: basicHtml,
                fileName: path.basename(filePath),
                isProcessing: false,
                mermaidCount: 0
            });
            
            // Then check if Mermaid CLI is available and preprocess in background
            const isMermaidInstalled = await this._checkMermaidInstalled();
            
            if (isMermaidInstalled && mermaidCount > 0) {
                // Send processing indicator with count
                this._panel.webview.postMessage({
                    command: 'updatePreview',
                    content: basicHtml,
                    fileName: path.basename(filePath),
                    isProcessing: true,
                    mermaidCount: mermaidCount
                });
                
                // Process Mermaid diagrams asynchronously
                this._processMarkdownMermaidAsync(content, path.basename(filePath));
            }
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading file: ${error}`);
        }
    }

    private async _processMarkdownMermaidAsync(content: string, fileName: string) {
        try {
            console.log('Starting Mermaid processing for:', fileName);
            const processedContent = await this._preprocessMarkdownForPreviewProgressive(content, fileName);
            const html = marked(processedContent);
            
            console.log('Mermaid processing completed for:', fileName);
            this._panel.webview.postMessage({
                command: 'updatePreview',
                content: html,
                fileName: fileName,
                isProcessing: false,
                mermaidCount: 0
            });
        } catch (error) {
            console.warn('Mermaid processing failed for preview:', error);
            // Keep the original preview that was already sent
            this._panel.webview.postMessage({
                command: 'previewProcessingComplete',
                fileName: fileName,
                error: 'Mermaid processing failed'
            });
        }
    }

    private async _browseOutputFolder() {
        try {
            const result = await vscode.window.showOpenDialog({
                canSelectMany: false,
                canSelectFiles: false,
                canSelectFolders: true,
                openLabel: 'Select Output Folder'
            });

            if (result?.[0]) {
                this._panel.webview.postMessage({
                    command: 'outputFolderSelected',
                    path: result[0].fsPath
                });
            }
        } catch (error) {
            vscode.window.showErrorMessage(`Error selecting output folder: ${error}`);
        }
    }

    // Pandoc Installation Methods
    private async _checkPandocInstalled(): Promise<boolean> {
        try {
            const execAsync = promisify(exec);
            
            try {
                await execAsync('pandoc --version');
                return true;
            } catch (error) {
                // Try common installation locations on Windows
                const username = require('os').userInfo().username;
                const commonPaths = [
                    'C:\\Program Files\\Pandoc\\pandoc.exe',
                    'C:\\Program Files (x86)\\Pandoc\\pandoc.exe',
                    `C:\\Users\\${username}\\AppData\\Local\\Pandoc\\pandoc.exe`,
                    `C:\\Users\\${username}\\AppData\\Local\\Microsoft\\WinGet\\Packages\\JohnMacFarlane.Pandoc_*\\pandoc.exe`
                ];
                
                for (const pandocPath of commonPaths) {
                    try {
                        if (fs.existsSync(pandocPath) || pandocPath.includes('*')) {
                            // For wildcard paths, try to find the actual path
                            if (pandocPath.includes('*')) {
                                const baseDir = path.dirname(pandocPath);
                                if (fs.existsSync(baseDir)) {
                                    const files = fs.readdirSync(baseDir);
                                    for (const file of files) {
                                        if (file.startsWith('JohnMacFarlane.Pandoc_')) {
                                            const actualPath = path.join(baseDir, file, 'pandoc.exe');
                                            if (fs.existsSync(actualPath)) {
                                                await execAsync(`"${actualPath}" --version`);
                                                return true;
                                            }
                                        }
                                    }
                                }
                            } else {
                                await execAsync(`"${pandocPath}" --version`);
                                return true;
                            }
                        }
                    } catch {}
                }
                
                return false;
            }
        } catch {
            return false;
        }
    }

    private async _installPandoc(): Promise<void> {
        try {
            const execAsync = promisify(exec);
            const platform = process.platform;

            if (platform === 'win32') {
                // Try multiple Windows package managers in order of preference
                const windowsMethods = [
                    {
                        name: 'WinGet',
                        command: 'winget install --id JohnMacFarlane.Pandoc --source winget --accept-package-agreements --accept-source-agreements',
                        checkCommand: 'winget --version'
                    },
                    {
                        name: 'Chocolatey',
                        command: 'choco install pandoc --force -y',
                        checkCommand: 'choco --version'
                    },
                    {
                        name: 'Scoop',
                        command: 'scoop install pandoc',
                        checkCommand: 'scoop --version'
                    }
                ];

                let installed = false;
                for (const method of windowsMethods) {
                    try {
                        // Check if package manager is available
                        await execAsync(method.checkCommand);
                        console.log(`Trying to install Pandoc using ${method.name}...`);
                        await execAsync(method.command);
                        installed = true;
                        break;
                    } catch (error) {
                        console.warn(`${method.name} installation failed:`, error);
                        continue;
                    }
                }

                if (!installed) {
                    throw new Error('Failed to install Pandoc automatically. Please install manually from https://pandoc.org/installing.html\n\nAlternatively, install a package manager:\n• WinGet (comes with Windows 10+)\n• Chocolatey: https://chocolatey.org/install\n• Scoop: https://scoop.sh/');
                }
            } else if (platform === 'darwin') {
                try {
                    // Check if Homebrew is available
                    await execAsync('brew --version');
                    console.log('Installing Pandoc using Homebrew...');
                    await execAsync('brew install pandoc');
                } catch (brewError) {
                    try {
                        // Try MacPorts as fallback
                        await execAsync('port version');
                        console.log('Installing Pandoc using MacPorts...');
                        await execAsync('sudo port install pandoc');
                    } catch (portError) {
                        throw new Error('Failed to install Pandoc automatically. Please install manually:\n\n1. Install Homebrew: https://brew.sh/\n2. Run: brew install pandoc\n\nOr download from: https://pandoc.org/installing.html');
                    }
                }
            } else {
                // Linux - try multiple package managers
                const linuxMethods = [
                    {
                        name: 'apt-get (Debian/Ubuntu)',
                        command: 'sudo apt-get update && sudo apt-get install -y pandoc',
                        checkCommand: 'apt-get --version'
                    },
                    {
                        name: 'yum (RHEL/CentOS)',
                        command: 'sudo yum install -y pandoc',
                        checkCommand: 'yum --version'
                    },
                    {
                        name: 'dnf (Fedora)',
                        command: 'sudo dnf install -y pandoc',
                        checkCommand: 'dnf --version'
                    },
                    {
                        name: 'pacman (Arch)',
                        command: 'sudo pacman -S --noconfirm pandoc',
                        checkCommand: 'pacman --version'
                    },
                    {
                        name: 'zypper (openSUSE)',
                        command: 'sudo zypper install -y pandoc',
                        checkCommand: 'zypper --version'
                    }
                ];

                let installed = false;
                for (const method of linuxMethods) {
                    try {
                        await execAsync(method.checkCommand);
                        console.log(`Installing Pandoc using ${method.name}...`);
                        await execAsync(method.command);
                        installed = true;
                        break;
                    } catch (error) {
                        console.warn(`${method.name} installation failed:`, error);
                        continue;
                    }
                }

                if (!installed) {
                    throw new Error('Failed to install Pandoc automatically. Please install manually:\n\n• Download from: https://pandoc.org/installing.html\n• Or use your distribution\'s package manager');
                }
            }

            // Verify installation after a short delay
            setTimeout(async () => {
                const isInstalled = await this._checkPandocInstalled();
                if (isInstalled) {
                    vscode.window.showInformationMessage('Pandoc installed successfully! You may need to restart VS Code for PATH changes to take effect.');
                } else {
                    vscode.window.showWarningMessage('Pandoc installation completed, but the command is not yet available. Please restart VS Code or your terminal.');
                }
            }, 2000);
        } catch (error) {
            throw error;
        }
    }

    // Node.js Installation Methods
    private async _checkNodeInstalled(): Promise<boolean> {
        try {
            const execAsync = promisify(exec);
            await execAsync('node --version');
            await execAsync('npm --version');
            return true;
        } catch {
            return false;
        }
    }

    private async _installNodejs(): Promise<void> {
        const platform = process.platform;
        const execAsync = promisify(exec);
        
        try {
            if (platform === 'win32') {
                // Try multiple Windows package managers
                const windowsMethods = [
                    {
                        name: 'WinGet',
                        command: 'winget install OpenJS.NodeJS --source winget --accept-package-agreements --accept-source-agreements',
                        checkCommand: 'winget --version'
                    },
                    {
                        name: 'Chocolatey',
                        command: 'choco install nodejs --force -y',
                        checkCommand: 'choco --version'
                    },
                    {
                        name: 'Scoop',
                        command: 'scoop install nodejs',
                        checkCommand: 'scoop --version'
                    }
                ];

                let installed = false;
                for (const method of windowsMethods) {
                    try {
                        await execAsync(method.checkCommand);
                        console.log(`Installing Node.js using ${method.name}...`);
                        await execAsync(method.command);
                        installed = true;
                        break;
                    } catch (error) {
                        console.warn(`${method.name} installation failed:`, error);
                        continue;
                    }
                }

                if (!installed) {
                    throw new Error('Failed to install Node.js automatically. Please install manually:\n\n• Download from: https://nodejs.org/ (recommend v18+ LTS)\n• Or install a package manager like WinGet, Chocolatey, or Scoop');
                }
            } else if (platform === 'darwin') {
                try {
                    // Check if Homebrew is available
                    await execAsync('brew --version');
                    console.log('Installing Node.js using Homebrew...');
                    // Install LTS version
                    await execAsync('brew install node@18 && brew link node@18');
                } catch (brewError) {
                    try {
                        // Try MacPorts as fallback
                        await execAsync('port version');
                        console.log('Installing Node.js using MacPorts...');
                        await execAsync('sudo port install nodejs18');
                    } catch (portError) {
                        throw new Error('Failed to install Node.js automatically. Please install manually:\n\n1. Install Homebrew: https://brew.sh/\n2. Run: brew install node\n\nOr download from: https://nodejs.org/ (recommend v18+ LTS)');
                    }
                }
            } else {
                // Linux - try multiple methods
                const linuxMethods = [
                    {
                        name: 'NodeSource Repository (Recommended)',
                        command: 'curl -fsSL https://deb.nodesource.com/setup_lts.x | sudo -E bash - && sudo apt-get install -y nodejs',
                        checkCommand: 'curl --version && apt-get --version'
                    },
                    {
                        name: 'apt-get (Ubuntu/Debian)',
                        command: 'sudo apt-get update && sudo apt-get install -y nodejs npm',
                        checkCommand: 'apt-get --version'
                    },
                    {
                        name: 'yum (RHEL/CentOS)',
                        command: 'sudo yum install -y nodejs npm',
                        checkCommand: 'yum --version'
                    },
                    {
                        name: 'dnf (Fedora)',
                        command: 'sudo dnf install -y nodejs npm',
                        checkCommand: 'dnf --version'
                    },
                    {
                        name: 'pacman (Arch)',
                        command: 'sudo pacman -S --noconfirm nodejs npm',
                        checkCommand: 'pacman --version'
                    }
                ];

                let installed = false;
                for (const method of linuxMethods) {
                    try {
                        await execAsync(method.checkCommand);
                        console.log(`Installing Node.js using ${method.name}...`);
                        await execAsync(method.command);
                        installed = true;
                        break;
                    } catch (error) {
                        console.warn(`${method.name} installation failed:`, error);
                        continue;
                    }
                }

                if (!installed) {
                    throw new Error('Failed to install Node.js automatically. Please install manually:\n\n• Download from: https://nodejs.org/ (recommend v18+ LTS)\n• Or use your distribution\'s package manager');
                }
            }

            // Verify installation after a short delay
            setTimeout(async () => {
                const isInstalled = await this._checkNodeInstalled();
                if (isInstalled) {
                    vscode.window.showInformationMessage('Node.js installed successfully! You may need to restart VS Code for PATH changes to take effect.');
                } else {
                    vscode.window.showWarningMessage('Node.js installation completed, but the command is not yet available. Please restart VS Code or your terminal.');
                }
            }, 2000);
        } catch (error) {
            throw error;
        }
    }

    // Mermaid Installation Methods
    private async _checkMermaidInstalled(): Promise<boolean> {
        try {
            const execAsync = promisify(exec);
            await execAsync('mmdc --version');
            return true;
        } catch {
            return false;
        }
    }

    private async _installMermaid(): Promise<void> {
        try {
            const execAsync = promisify(exec);
            
            // Check if Node.js is installed first
            const nodeInstalled = await this._checkNodeInstalled();
            if (!nodeInstalled) {
                throw new Error('Node.js is required to install Mermaid CLI. Please install Node.js first.');
            }

            console.log('Installing Mermaid CLI globally...');
            
            // Try different npm installation approaches
            const installMethods = [
                {
                    name: 'npm (global)',
                    command: 'npm install -g @mermaid-js/mermaid-cli@latest'
                },
                {
                    name: 'npm (with force)',
                    command: 'npm install -g @mermaid-js/mermaid-cli@latest --force'
                },
                {
                    name: 'yarn (if available)',
                    command: 'yarn global add @mermaid-js/mermaid-cli@latest'
                }
            ];

            let installed = false;
            for (const method of installMethods) {
                try {
                    console.log(`Trying to install Mermaid CLI using ${method.name}...`);
                    await execAsync(method.command);
                    installed = true;
                    break;
                } catch (error) {
                    console.warn(`${method.name} installation failed:`, error);
                    continue;
                }
            }

            if (!installed) {
                throw new Error('Failed to install Mermaid CLI automatically. Please try manually:\n\n1. Open terminal/command prompt\n2. Run: npm install -g @mermaid-js/mermaid-cli@latest\n\nIf that fails, you may need to:\n• Update npm: npm install -g npm@latest\n• Clear npm cache: npm cache clean --force\n• Check permissions for global npm installs');
            }

            // Verify installation after a short delay
            setTimeout(async () => {
                const isInstalled = await this._checkMermaidInstalled();
                if (isInstalled) {
                    vscode.window.showInformationMessage('Mermaid CLI installed successfully! Diagram rendering is now available.');
                } else {
                    vscode.window.showWarningMessage('Mermaid CLI installation completed, but the command is not yet available. Please restart VS Code or your terminal.');
                }
            }, 3000); // Longer delay for npm global installs
        } catch (error) {
            throw error;
        }
    }

    private async _convertMermaidToImage(mermaidCode: string): Promise<string> {
        try {
            const tempDir = path.join(require('os').tmpdir(), 'doculate-mermaid');
            if (!fs.existsSync(tempDir)) {
                fs.mkdirSync(tempDir, { recursive: true });
            }

            const timestamp = Date.now();
            const inputFile = path.join(tempDir, `diagram-${timestamp}.mmd`);
            const outputFile = path.join(tempDir, `diagram-${timestamp}.png`);

            // Clean and prepare Mermaid code
            let cleanedCode = mermaidCode.trim();
            
            // Check for unsupported diagram types
            if (cleanedCode.includes('gitgraph')) {
                console.warn('Git graph diagrams are not supported by the current Mermaid CLI version');
                throw new Error('Git graph diagrams are not supported by the current Mermaid CLI version. Please update Mermaid CLI or use a different diagram type.');
            }
            
            // Handle different diagram types that might need special syntax
            if (cleanedCode.includes('classDiagram')) {
                console.log('Processing class diagram...');
                
                // Convert complex relationship syntax to simpler compatible syntax
                cleanedCode = cleanedCode
                    .replace(/\|\|--o\{/g, '-->')  // Convert ||--o{ to -->
                    .replace(/\s*:\s*"([^"]+)"/g, ' : $1');  // Remove quotes from relationships
            }

            // Write Mermaid code to temp file with explicit UTF-8 encoding (without BOM)
            const buffer = Buffer.from(cleanedCode, 'utf8');
            fs.writeFileSync(inputFile, buffer);

            // Convert using Mermaid CLI with more robust options
            const execAsync = promisify(exec);
            const command = `mmdc -i "${inputFile}" -o "${outputFile}" -t neutral -b white --scale 2`;
            
            console.log(`Executing Mermaid command: ${command}`);
            const result = await execAsync(command);
            console.log(`Mermaid CLI output: ${result.stdout || 'No stdout'}`);
            if (result.stderr) {
                console.warn(`Mermaid CLI stderr: ${result.stderr}`);
            }

            // Verify output file exists
            if (!fs.existsSync(outputFile)) {
                throw new Error('Mermaid conversion failed - no output file generated');
            }

            // Check if file has content
            const stats = fs.statSync(outputFile);
            if (stats.size === 0) {
                throw new Error('Mermaid conversion failed - empty output file generated');
            }

            console.log(`Successfully generated Mermaid image: ${outputFile} (${stats.size} bytes)`);
            return outputFile;
        } catch (error) {
            console.error('Detailed Mermaid conversion error:', error);
            throw new Error(`Mermaid conversion failed: ${error}`);
        }
    }

    private async _preprocessMarkdownForMermaid(markdownPath: string): Promise<string> {
        try {
            const content = fs.readFileSync(markdownPath, 'utf8');
            const lines = content.split('\n');
            let processedLines: string[] = [];
            let inMermaidBlock = false;
            let mermaidCode: string[] = [];

            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];

                if (line.trim().startsWith('```mermaid')) {
                    inMermaidBlock = true;
                    mermaidCode = [];
                    continue;
                }

                if (inMermaidBlock && line.trim() === '```') {
                    // Convert Mermaid to image
                    try {
                        const mermaidSource = mermaidCode.join('\n');
                        if (mermaidSource.trim()) {
                            const imagePath = await this._convertMermaidToImage(mermaidSource);
                            
                            // Add image reference to markdown
                            processedLines.push(`![Mermaid Diagram](${imagePath})`);
                            processedLines.push(''); // Empty line after image
                        }
                    } catch (error) {
                        // Fallback to code block if conversion fails
                        console.warn('Mermaid conversion failed, using code block:', error);
                        processedLines.push('```mermaid');
                        processedLines.push(...mermaidCode);
                        processedLines.push('```');
                    }
                    
                    inMermaidBlock = false;
                    mermaidCode = [];
                    continue;
                }

                if (inMermaidBlock) {
                    mermaidCode.push(line);
                } else {
                    processedLines.push(line);
                }
            }

            // Write processed markdown to temp file
            const tempDir = path.join(require('os').tmpdir(), 'doculate-processed');
            if (!fs.existsSync(tempDir)) {
                fs.mkdirSync(tempDir, { recursive: true });
            }

            const processedPath = path.join(tempDir, `processed-${Date.now()}.md`);
            fs.writeFileSync(processedPath, processedLines.join('\n'));

            return processedPath;
        } catch (error) {
            throw new Error(`Markdown preprocessing failed: ${error}`);
        }
    }

    private async _preprocessMarkdownForPreviewProgressive(content: string, fileName: string): Promise<string> {
        try {
            console.log('Starting progressive Mermaid preprocessing...');
            const lines = content.split('\n');
            let processedLines: string[] = [];
            let inMermaidBlock = false;
            let mermaidCode: string[] = [];
            let diagramCount = 0;
            let processedCount = 0;
            const maxDiagrams = 10; // Limit number of diagrams to process for performance
            
            // Count total diagrams first
            const totalDiagrams = (content.match(/```mermaid[\s\S]*?```/g) || []).length;

            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];

                if (line.trim().startsWith('```mermaid')) {
                    inMermaidBlock = true;
                    mermaidCode = [];
                    continue;
                }

                if (inMermaidBlock && line.trim() === '```') {
                    diagramCount++;
                    console.log(`Processing Mermaid diagram ${diagramCount}...`);
                    
                    // Limit processing for performance
                    if (diagramCount > maxDiagrams) {
                        console.warn(`Skipping Mermaid diagram ${diagramCount} (limit: ${maxDiagrams}) for performance`);
                        processedLines.push('```mermaid');
                        processedLines.push(...mermaidCode);
                        processedLines.push('```');
                    } else {
                        // Convert Mermaid to base64 image for preview
                        try {
                            const mermaidSource = mermaidCode.join('\n');
                            if (mermaidSource.trim()) {
                                // Check cache first
                                const cacheKey = Buffer.from(mermaidSource).toString('base64');
                                let dataUrl = this._mermaidCache.get(cacheKey);
                                
                                if (!dataUrl) {
                                    console.log(`Converting Mermaid diagram ${diagramCount} to image...`);
                                    // Add timeout for individual diagram processing
                                    const imagePath = await Promise.race([
                                        this._convertMermaidToImage(mermaidSource),
                                        new Promise<never>((_, reject) => 
                                            setTimeout(() => reject(new Error('Timeout')), 5000)
                                        )
                                    ]);
                                    
                                    // Convert image to base64 for webview
                                    const imageBuffer = fs.readFileSync(imagePath);
                                    const base64Image = imageBuffer.toString('base64');
                                    dataUrl = `data:image/png;base64,${base64Image}`;
                                    
                                    // Cache the result
                                    this._mermaidCache.set(cacheKey, dataUrl);
                                    console.log(`Cached Mermaid diagram ${diagramCount}`);
                                    
                                    // Limit cache size
                                    if (this._mermaidCache.size > 50) {
                                        const firstKey = this._mermaidCache.keys().next().value;
                                        if (firstKey) {
                                            this._mermaidCache.delete(firstKey);
                                        }
                                    }
                                    
                                    // Clean up temporary image file
                                    try {
                                        fs.unlinkSync(imagePath);
                                    } catch (cleanupError) {
                                        console.warn('Failed to clean up temporary image:', cleanupError);
                                    }
                                } else {
                                    console.log(`Using cached Mermaid diagram ${diagramCount}`);
                                }
                                
                                // Add image reference to markdown
                                processedLines.push(`![Mermaid Diagram ${diagramCount}](${dataUrl})`);
                                processedLines.push(''); // Empty line after image
                                
                                processedCount++;
                                
                                // Send progressive update
                                const remainingCount = totalDiagrams - processedCount;
                                if (remainingCount > 0) {
                                    const currentContent = [...processedLines, ...lines.slice(i + 1)].join('\n');
                                    const currentHtml = marked(currentContent);
                                    this._panel.webview.postMessage({
                                        command: 'updatePreview',
                                        content: currentHtml,
                                        fileName: fileName,
                                        isProcessing: true,
                                        mermaidCount: remainingCount
                                    });
                                }
                            }
                        } catch (error) {
                            // Fallback to code block if conversion fails
                            console.warn(`Mermaid conversion failed for diagram ${diagramCount}, using code block:`, error);
                            
                            // Add a helpful comment in the code block
                            const errorMessage = String(error).includes('not supported') 
                                ? `\n<!-- Note: ${String(error).split(':')[1]?.trim() || 'This diagram type is not supported'} -->`
                                : '\n<!-- Note: Mermaid conversion failed - displaying as code block -->';
                                
                            processedLines.push('```mermaid' + errorMessage);
                            processedLines.push(...mermaidCode);
                            processedLines.push('```');
                        }
                    }
                    
                    inMermaidBlock = false;
                    mermaidCode = [];
                    continue;
                }

                if (inMermaidBlock) {
                    mermaidCode.push(line);
                } else {
                    processedLines.push(line);
                }
            }

            console.log(`Progressive Mermaid preprocessing completed. Processed ${processedCount} diagrams.`);
            return processedLines.join('\n');
        } catch (error) {
            throw new Error(`Progressive markdown preprocessing failed for preview: ${error}`);
        }
    }

    private async _exportWithPandoc(markdownPath: string, outputPath: string, templatePath?: string): Promise<void> {
        const execAsync = promisify(exec);
        
        // Find pandoc executable
        let pandocPath = 'pandoc';
        
        try {
            await execAsync('pandoc --version');
        } catch {
            // Try common installation locations
            const username = require('os').userInfo().username;
            const commonPaths = [
                `C:\\Users\\${username}\\AppData\\Local\\Pandoc\\pandoc.exe`,
                'C:\\Program Files\\Pandoc\\pandoc.exe',
                'C:\\Program Files (x86)\\Pandoc\\pandoc.exe'
            ];
            
            let found = false;
            for (const path of commonPaths) {
                if (fs.existsSync(path)) {
                    pandocPath = path;
                    found = true;
                    break;
                }
            }
            
            if (!found) {
                throw new Error('Pandoc executable not found. Please install Pandoc.');
            }
        }
        
        // Check if Mermaid CLI is available and preprocess if needed
        const isMermaidInstalled = await this._checkMermaidInstalled();
        let processedMarkdownPath = markdownPath;
        let tempFiles: string[] = [];
        
        if (isMermaidInstalled) {
            try {
                // Preprocess markdown to convert Mermaid diagrams to images
                processedMarkdownPath = await this._preprocessMarkdownForMermaid(markdownPath);
                tempFiles.push(processedMarkdownPath);
            } catch (error) {
                console.warn('Mermaid preprocessing failed, using original markdown:', error);
                // Continue with original markdown if preprocessing fails
            }
        }
        
        // Validate input files
        if (!fs.existsSync(processedMarkdownPath)) {
            throw new Error(`Processed markdown file not found: ${processedMarkdownPath}`);
        }
        
        // Build pandoc command
        let command = `"${pandocPath}" "${processedMarkdownPath}" -o "${outputPath}"`;
        
        // Handle template/reference document
        if (templatePath && templatePath !== 'basic' && templatePath !== '') {
            if (!fs.existsSync(templatePath)) {
                const templateName = path.basename(templatePath);
                throw new Error(`Reference document "${templateName}" not found at: ${templatePath}. The file may have been moved or deleted. Please update your reference documents in settings.`);
            }
            
            // Validate the template file is readable
            try {
                const stats = fs.statSync(templatePath);
                if (stats.size === 0) {
                    const templateName = path.basename(templatePath);
                    throw new Error(`Reference document "${templateName}" is empty and cannot be used.`);
                }
            } catch (error) {
                const templateName = path.basename(templatePath);
                throw new Error(`Cannot access reference document "${templateName}": ${error}. Please check file permissions.`);
            }
            
            if (templatePath.toLowerCase().endsWith('.docx')) {
                command += ` --reference-doc="${templatePath}"`;
            } else {
                command += ` --template="${templatePath}"`;
            }
        }

        try {
            const result = await execAsync(command);
            
            // Verify output file was created
            if (!fs.existsSync(outputPath) || fs.statSync(outputPath).size === 0) {
                throw new Error('Pandoc completed but output file was not created properly');
            }
        } catch (error) {
            throw new Error(`Pandoc conversion failed: ${error}`);
        } finally {
            // Clean up temporary files
            for (const tempFile of tempFiles) {
                try {
                    if (fs.existsSync(tempFile)) {
                        fs.unlinkSync(tempFile);
                    }
                } catch (cleanupError) {
                    console.warn('Failed to clean up temporary file:', cleanupError);
                }
            }
            
            // Clean up temporary directories
            try {
                const mermaidTempDir = path.join(require('os').tmpdir(), 'doculate-mermaid');
                const processedTempDir = path.join(require('os').tmpdir(), 'doculate-processed');
                
                if (fs.existsSync(mermaidTempDir)) {
                    fs.rmSync(mermaidTempDir, { recursive: true, force: true });
                }
                if (fs.existsSync(processedTempDir)) {
                    fs.rmSync(processedTempDir, { recursive: true, force: true });
                }
            } catch (cleanupError) {
                console.warn('Failed to clean up temporary directories:', cleanupError);
            }
        }
    }

    private async _handlePandocInstallation() {
        try {
            const isInstalled = await this._checkPandocInstalled();
            if (isInstalled) {
                vscode.window.showInformationMessage('Pandoc is already installed!');
                return;
            }

            const choice = await vscode.window.showInformationMessage(
                'Pandoc is not installed. Would you like to install it now?',
                'Install Pandoc',
                'Cancel'
            );

            if (choice === 'Install Pandoc') {
                await vscode.window.withProgress({
                    location: vscode.ProgressLocation.Notification,
                    title: "Installing Pandoc...",
                    cancellable: false
                }, async () => {
                    await this._installPandoc();
                });
                vscode.window.showInformationMessage('Pandoc installed successfully! You may need to restart VS Code.');
                
                // Refresh status
                setTimeout(() => {
                    this._checkAndUpdatePandocStatus();
                }, 2000);
            }
        } catch (error) {
            vscode.window.showErrorMessage(`Error installing Pandoc: ${error}`);
        }
    }

    private async _handleNodejsInstallation() {
        try {
            const isInstalled = await this._checkNodeInstalled();
            if (isInstalled) {
                vscode.window.showInformationMessage('Node.js is already installed!');
                return;
            }

            const choice = await vscode.window.showInformationMessage(
                'Node.js is not installed. Would you like to install it now?',
                'Install Node.js',
                'Open Node.js Website',
                'Cancel'
            );

            if (choice === 'Install Node.js') {
                await vscode.window.withProgress({
                    location: vscode.ProgressLocation.Notification,
                    title: "Installing Node.js...",
                    cancellable: false
                }, async () => {
                    await this._installNodejs();
                });
                vscode.window.showInformationMessage('Node.js installed successfully! You may need to restart VS Code.');
                
                // Refresh status
                setTimeout(() => {
                    this._checkAndUpdateMermaidStatus();
                }, 2000);
            } else if (choice === 'Open Node.js Website') {
                vscode.env.openExternal(vscode.Uri.parse('https://nodejs.org/'));
            }
        } catch (error) {
            vscode.window.showErrorMessage(`Error installing Node.js: ${error}`);
        }
    }

    private async _handleMermaidInstallation() {
        try {
            const isInstalled = await this._checkMermaidInstalled();
            if (isInstalled) {
                vscode.window.showInformationMessage('Mermaid CLI is already installed!');
                return;
            }

            const nodeInstalled = await this._checkNodeInstalled();
            if (!nodeInstalled) {
                const choice = await vscode.window.showInformationMessage(
                    'Mermaid CLI requires Node.js which is not installed. Would you like to install Node.js first?',
                    'Install Node.js',
                    'Open Node.js Website',
                    'Cancel'
                );
                
                if (choice === 'Install Node.js') {
                    await this._handleNodejsInstallation();
                } else if (choice === 'Open Node.js Website') {
                    vscode.env.openExternal(vscode.Uri.parse('https://nodejs.org/'));
                }
                return;
            }

            const choice = await vscode.window.showInformationMessage(
                'Mermaid CLI is not installed. Would you like to install it now? This will install it globally using npm.',
                'Install Mermaid CLI',
                'Cancel'
            );

            if (choice === 'Install Mermaid CLI') {
                await vscode.window.withProgress({
                    location: vscode.ProgressLocation.Notification,
                    title: "Installing Mermaid CLI...",
                    cancellable: false
                }, async () => {
                    await this._installMermaid();
                });
                vscode.window.showInformationMessage('Mermaid CLI installed successfully!');
                
                // Refresh status
                setTimeout(() => {
                    this._checkAndUpdateMermaidStatus();
                }, 2000);
            }
        } catch (error) {
            vscode.window.showErrorMessage(`Error installing Mermaid CLI: ${error}`);
        }
    }

    private async _checkAndUpdatePandocStatus() {
        try {
            const isInstalled = await this._checkPandocInstalled();
            this._panel.webview.postMessage({
                command: 'updatePandocStatus',
                isInstalled: isInstalled
            });
        } catch (error) {
            this._panel.webview.postMessage({
                command: 'updatePandocStatus',
                isInstalled: false,
                error: error instanceof Error ? error.message : String(error)
            });
        }
    }

    private async _checkAndUpdateMermaidStatus() {
        try {
            const isInstalled = await this._checkMermaidInstalled();
            const nodeInstalled = await this._checkNodeInstalled();
            
            this._panel.webview.postMessage({
                command: 'updateMermaidStatus',
                isInstalled: isInstalled,
                nodeInstalled: nodeInstalled
            });
        } catch (error) {
            this._panel.webview.postMessage({
                command: 'updateMermaidStatus',
                isInstalled: false,
                nodeInstalled: false,
                error: error instanceof Error ? error.message : String(error)
            });
        }
    }

    private async _addTemplate() {
        try {
            const result = await vscode.window.showOpenDialog({
                canSelectMany: false,
                canSelectFiles: true,
                canSelectFolders: false,
                filters: {
                    'Word Documents': ['docx']
                },
                openLabel: 'Add Reference Document'
            });

            if (result?.[0]) {
                const templateName = path.basename(result[0].fsPath);
                const templatePath = result[0].fsPath;
                
                // Validate the file exists and is readable
                if (!fs.existsSync(templatePath)) {
                    vscode.window.showErrorMessage('Selected file does not exist.');
                    return;
                }
                
                // Check file size (reasonable limit for Word documents)
                const stats = fs.statSync(templatePath);
                if (stats.size === 0) {
                    vscode.window.showErrorMessage('Selected file is empty and cannot be used as a reference document.');
                    return;
                }
                
                if (stats.size > 50 * 1024 * 1024) { // 50MB limit
                    vscode.window.showWarningMessage('Selected file is very large (>50MB). This may cause performance issues during export.');
                }
                
                const templates = this._context.globalState.get('doculate.templates', []) as Array<{name: string, path: string}>;
                const newTemplate = { name: templateName, path: templatePath };
                
                const existingIndex = templates.findIndex(t => t.name === templateName);
                if (existingIndex >= 0) {
                    const choice = await vscode.window.showWarningMessage(
                        `A reference document named "${templateName}" already exists. Replace it?`,
                        'Replace',
                        'Cancel'
                    );
                    
                    if (choice !== 'Replace') {
                        return;
                    }
                    
                    templates[existingIndex] = newTemplate;
                } else {
                    templates.push(newTemplate);
                }
                
                await this._context.globalState.update('doculate.templates', templates);
                vscode.window.showInformationMessage(`Reference document added: ${templateName}`);
                
                // Reload templates in UI
                this._loadSavedTemplates();
            }
        } catch (error) {
            vscode.window.showErrorMessage(`Error adding template: ${error}`);
        }
    }

    private async _removeTemplate(templateName: string) {
        try {
            const templates = this._context.globalState.get('doculate.templates', []) as Array<{name: string, path: string}>;
            const updatedTemplates = templates.filter(t => t.name !== templateName);
            
            await this._context.globalState.update('doculate.templates', updatedTemplates);
            vscode.window.showInformationMessage(`Reference document removed: ${templateName}`);
            
            // Reload templates in UI
            this._loadSavedTemplates();
        } catch (error) {
            vscode.window.showErrorMessage(`Error removing template: ${error}`);
        }
    }

    private async _getWorkspacePath() {
        try {
            const workspaceFolders = vscode.workspace.workspaceFolders;
            const workspacePath = workspaceFolders?.[0]?.uri.fsPath || require('os').homedir();
            
            this._panel.webview.postMessage({
                command: 'workspacePathSelected',
                path: workspacePath
            });
        } catch (error) {
            vscode.window.showErrorMessage(`Error getting workspace path: ${error}`);
        }
    }

    private async _showTemplateData() {
        try {
            const templates = this._context.globalState.get('doculate.templates', []) as Array<{name: string, path: string}>;
            
            const message = templates.length > 0 
                ? `Templates stored in VS Code's internal database:\n\n${templates.map((t, i) => `${i + 1}. ${t.name}\n   Path: ${t.path}`).join('\n\n')}`
                : `No templates are currently stored in VS Code's internal database.`;
            
            vscode.window.showInformationMessage(message, 'OK');
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading template data: ${error}`);
        }
    }

    private async _exportToWord(exportData: any) {
        try {
            const { filename, templatePath, outputPath, sourceFile, engine: initialEngine } = exportData;
            let engine = initialEngine;
            
            await vscode.window.withProgress({
                location: vscode.ProgressLocation.Notification,
                title: "Exporting to Word...",
                cancellable: false
            }, async (progress) => {
                progress.report({ increment: 20, message: "Reading markdown file..." });
                
                const outputFilePath = path.join(outputPath, `${filename}.docx`);
                
                // Ensure output directory exists
                const outputDir = path.dirname(outputFilePath);
                if (!fs.existsSync(outputDir)) {
                    fs.mkdirSync(outputDir, { recursive: true });
                }

                if (engine === 'pandoc') {
                    progress.report({ increment: 40, message: "Converting with Pandoc..." });
                    
                    const isPandocInstalled = await this._checkPandocInstalled();
                    if (!isPandocInstalled) {
                        const installChoice = await vscode.window.showInformationMessage(
                            'Pandoc is not installed. Would you like to install it now?',
                            'Install Pandoc',
                            'Use Docx Library',
                            'Cancel'
                        );
                        
                        if (installChoice === 'Install Pandoc') {
                            progress.report({ increment: 50, message: "Installing Pandoc..." });
                            await this._installPandoc();
                        } else if (installChoice === 'Use Docx Library') {
                            engine = 'docx';
                        } else {
                            throw new Error('Export cancelled');
                        }
                    }
                    
                    if (engine === 'pandoc') {
                        await this._exportWithPandoc(sourceFile.path, outputFilePath, templatePath);
                    }
                }

                if (engine === 'docx') {
                    progress.report({ increment: 40, message: "Converting with Docx Library..." });
                    
                    const markdownContent = fs.readFileSync(sourceFile.path, 'utf8');
                    const wordElements = await this._parseMarkdownToWord(markdownContent);
                    
                    const doc = new Document({
                        numbering: {
                            config: [{
                                reference: "default-numbering",
                                levels: [{
                                    level: 0,
                                    format: "decimal",
                                    text: "%1.",
                                    alignment: AlignmentType.START,
                                    style: {
                                        paragraph: {
                                            indent: { left: 720, hanging: 260 },
                                        },
                                    },
                                }],
                            }],
                        },
                        sections: [{
                            properties: {},
                            children: wordElements,
                        }],
                    });
                    
                    progress.report({ increment: 70, message: "Generating document..." });
                    
                    const buffer = await Packer.toBuffer(doc);
                    if (!buffer || buffer.length === 0) {
                        throw new Error('Failed to generate Word document buffer');
                    }
                    
                    progress.report({ increment: 90, message: "Saving file..." });
                    
                    fs.writeFileSync(outputFilePath, buffer);
                    
                    if (!fs.existsSync(outputFilePath) || fs.statSync(outputFilePath).size === 0) {
                        throw new Error('File was not created successfully');
                    }
                }
                
                progress.report({ increment: 100, message: "Complete!" });
                
                const choice = await vscode.window.showInformationMessage(
                    `Document exported successfully to: ${path.basename(outputFilePath)}`,
                    'Open File',
                    'Show in Folder'
                );
                
                if (choice === 'Open File') {
                    await vscode.env.openExternal(vscode.Uri.file(outputFilePath));
                } else if (choice === 'Show in Folder') {
                    await vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(outputFilePath));
                }
                
                this._panel.webview.postMessage({
                    command: 'exportComplete',
                    path: outputFilePath
                });
            });
        } catch (error) {
            vscode.window.showErrorMessage(`Export failed: ${error}`);
        }
    }

    private async _parseMarkdownToWord(markdown: string): Promise<any[]> {
        const elements: any[] = [];
        const lines = markdown.split('\n');
        let currentParagraph: string[] = [];
        
        const flushParagraph = () => {
            if (currentParagraph.length > 0) {
                const text = currentParagraph.join(' ').trim();
                if (text) {
                    elements.push(new Paragraph({
                        children: this._parseInlineMarkdown(text),
                        spacing: { after: 200 }
                    }));
                }
                currentParagraph = [];
            }
        };

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            if (!line) {
                flushParagraph();
                continue;
            }
            
            // Headers
            if (line.startsWith('#')) {
                flushParagraph();
                const headerLevel = line.match(/^#+/)?.[0].length || 1;
                const headerText = line.replace(/^#+\s*/, '');
                
                const headingLevels = [
                    HeadingLevel.HEADING_1,
                    HeadingLevel.HEADING_2,
                    HeadingLevel.HEADING_3,
                    HeadingLevel.HEADING_4,
                    HeadingLevel.HEADING_5,
                    HeadingLevel.HEADING_6
                ];
                
                elements.push(new Paragraph({
                    text: headerText,
                    heading: headingLevels[Math.min(headerLevel - 1, 5)],
                    spacing: { before: 240, after: 120 }
                }));
                continue;
            }
            
            // Code blocks (including Mermaid)
            if (line.startsWith('```')) {
                flushParagraph();
                const codeType = line.replace('```', '').trim();
                const codeLines: string[] = [];
                i++;
                
                while (i < lines.length && !lines[i].trim().startsWith('```')) {
                    codeLines.push(lines[i]);
                    i++;
                }
                
                if (codeLines.length > 0) {
                    if (codeType === 'mermaid') {
                        // Try to convert Mermaid to image if CLI is available
                        try {
                            const isMermaidInstalled = await this._checkMermaidInstalled();
                            if (isMermaidInstalled) {
                                const mermaidCode = codeLines.join('\n');
                                const imagePath = await this._convertMermaidToImage(mermaidCode);
                                
                                // Add image to document
                                const imageBuffer = fs.readFileSync(imagePath);
                                elements.push(new Paragraph({
                                    children: [new ImageRun({
                                        data: imageBuffer,
                                        transformation: {
                                            width: 600,
                                            height: 400,
                                        },
                                    })],
                                    spacing: { before: 120, after: 120 }
                                }));
                                
                                // Clean up temp image
                                try {
                                    fs.unlinkSync(imagePath);
                                } catch {}
                                
                                continue;
                            }
                        } catch (error) {
                            console.warn('Failed to convert Mermaid diagram:', error);
                        }
                    }
                    
                    // Fallback to code block
                    elements.push(new Paragraph({
                        children: [new TextRun({
                            text: codeLines.join('\n'),
                            font: 'Courier New',
                            size: 20
                        })],
                        spacing: { before: 120, after: 120 },
                        indent: { left: 720 }
                    }));
                }
                continue;
            }
            
            // Lists
            if (line.match(/^[-*+]\s+/) || line.match(/^\d+\.\s+/)) {
                flushParagraph();
                const listText = line.replace(/^[-*+\d.]\s+/, '');
                const isNumbered = line.match(/^\d+\.\s+/);
                
                const paragraph = new Paragraph({
                    children: this._parseInlineMarkdown(listText),
                    indent: { left: 720 },
                    spacing: { after: 100 }
                });
                
                if (isNumbered) {
                    paragraph.numbering = { reference: "default-numbering", level: 0 };
                }
                
                elements.push(paragraph);
                continue;
            }
            
            // Blockquotes
            if (line.startsWith('>')) {
                flushParagraph();
                const quoteText = line.replace(/^>\s*/, '');
                elements.push(new Paragraph({
                    children: this._parseInlineMarkdown(quoteText),
                    indent: { left: 720 },
                    spacing: { after: 120 }
                }));
                continue;
            }
            
            currentParagraph.push(line);
        }
        
        flushParagraph();
        return elements;
    }

    private _parseInlineMarkdown(text: string): any[] {
        // Handle bold text
        if (text.includes('**')) {
            const parts = text.split('**');
            const runs: any[] = [];
            
            for (let i = 0; i < parts.length; i++) {
                if (parts[i]) {
                    runs.push(new TextRun({
                        text: parts[i],
                        bold: i % 2 === 1
                    }));
                }
            }
            
            return runs.length > 0 ? runs : [new TextRun(text)];
        }
        
        // Handle italic text
        if (text.includes('*') && !text.includes('**')) {
            const parts = text.split('*');
            const runs: any[] = [];
            
            for (let i = 0; i < parts.length; i++) {
                if (parts[i]) {
                    runs.push(new TextRun({
                        text: parts[i],
                        italics: i % 2 === 1
                    }));
                }
            }
            
            return runs.length > 0 ? runs : [new TextRun(text)];
        }
        
        // Handle inline code
        if (text.includes('`')) {
            const parts = text.split('`');
            const runs: any[] = [];
            
            for (let i = 0; i < parts.length; i++) {
                if (parts[i]) {
                    runs.push(new TextRun({
                        text: parts[i],
                        font: i % 2 === 1 ? 'Courier New' : undefined
                    }));
                }
            }
            
            return runs.length > 0 ? runs : [new TextRun(text)];
        }
        
        return [new TextRun(text)];
    }

    private _update() {
        this._panel.webview.html = this._getHtmlForWebview();
        this._loadSavedTemplates();
    }

    private async _loadSavedTemplates() {
        try {
            const templates = this._context.globalState.get('doculate.templates', []) as Array<{name: string, path: string}>;
            
            // Validate that template files still exist and filter out invalid ones
            const validTemplates: Array<{name: string, path: string}> = [];
            let removedCount = 0;
            
            for (const template of templates) {
                if (fs.existsSync(template.path)) {
                    validTemplates.push(template);
                } else {
                    console.warn(`Template file no longer exists: ${template.path}`);
                    removedCount++;
                }
            }
            
            // Update stored templates if any were removed
            if (removedCount > 0) {
                await this._context.globalState.update('doculate.templates', validTemplates);
                if (removedCount === 1) {
                    vscode.window.showWarningMessage(`1 reference document was removed because its file no longer exists.`);
                } else {
                    vscode.window.showWarningMessage(`${removedCount} reference documents were removed because their files no longer exist.`);
                }
            }
            
            this._panel.webview.postMessage({
                command: 'loadSavedTemplates',
                templates: validTemplates
            });
        } catch (error) {
            console.error('Error loading saved templates:', error);
        }
    }

    private _getHtmlForWebview() {
        return `<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Doculate - Markdown to Word</title>
            <style>
                * {
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                }
                
                body {
                    font-family: var(--vscode-font-family);
                    color: var(--vscode-foreground);
                    background: var(--vscode-editor-background);
                    height: 100vh;
                    display: flex;
                }

                .sidebar {
                    width: 320px;
                    background: var(--vscode-sideBar-background);
                    border-right: 1px solid var(--vscode-sideBar-border);
                    display: flex;
                    flex-direction: column;
                }

                .panel-header {
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    padding: 12px 16px;
                    border-bottom: 1px solid var(--vscode-sideBar-border);
                }

                .header-title {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    font-weight: 600;
                }

                .settings-button {
                    background: transparent;
                    border: none;
                    color: var(--vscode-foreground);
                    cursor: pointer;
                    padding: 4px;
                    border-radius: 3px;
                    font-size: 16px;
                    opacity: 0.7;
                }

                .settings-button:hover {
                    opacity: 1;
                    background: var(--vscode-toolbar-hoverBackground);
                }

                .sidebar-header {
                    padding: 16px;
                    border-bottom: 1px solid var(--vscode-sideBar-border);
                }

                .file-selector label {
                    display: block;
                    margin-bottom: 8px;
                    font-size: 12px;
                    color: var(--vscode-foreground);
                    opacity: 0.8;
                }

                .file-dropdown {
                    width: 100%;
                    padding: 8px;
                    background: var(--vscode-input-background);
                    border: 1px solid var(--vscode-input-border);
                    color: var(--vscode-input-foreground);
                    border-radius: 4px;
                }

                .current-file-status {
                    padding: 16px;
                    border-bottom: 1px solid var(--vscode-sideBar-border);
                }

                .file-info-display {
                    display: flex;
                    align-items: flex-start;
                    gap: 12px;
                }

                .file-icon {
                    width: 20px;
                    height: 20px;
                    background: var(--vscode-button-background);
                    color: var(--vscode-button-foreground);
                    border-radius: 3px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 10px;
                    margin-top: 2px;
                }

                .file-details {
                    flex: 1;
                    min-width: 0;
                }

                .file-name-display {
                    font-weight: 500;
                    margin-bottom: 4px;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                }

                .file-type-display {
                    font-size: 12px;
                    color: var(--vscode-foreground);
                    opacity: 0.7;
                    margin-bottom: 8px;
                }

                .doculate-badge {
                    background: rgba(0, 200, 0, 0.2);
                    color: #00c800;
                    border: 1px solid rgba(0, 200, 0, 0.3);
                    padding: 2px 8px;
                    border-radius: 3px;
                    font-size: 11px;
                    font-weight: 500;
                }

                .main-content {
                    flex: 1;
                    display: flex;
                    flex-direction: column;
                }

                .preview-header {
                    padding: 16px;
                    border-bottom: 1px solid var(--vscode-sideBar-border);
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    gap: 12px;
                }

                .preview-title {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    flex: 1;
                }

                .preview-actions {
                    display: flex;
                    gap: 8px;
                }

                .export-button {
                    padding: 6px 12px;
                    background: var(--vscode-button-background);
                    color: var(--vscode-button-foreground);
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    font-family: var(--vscode-font-family);
                    font-size: 12px;
                    display: flex;
                    align-items: center;
                    gap: 6px;
                }

                .export-button:hover:not(:disabled) {
                    background: var(--vscode-button-hoverBackground);
                }

                .export-button:disabled {
                    opacity: 0.5;
                    cursor: not-allowed;
                }

                .preview-content {
                    flex: 1;
                    overflow-y: auto;
                    padding: 24px;
                }

                .empty-state {
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    justify-content: center;
                    height: 100%;
                    color: var(--vscode-foreground);
                    opacity: 0.6;
                }

                .markdown-content {
                    max-width: 800px;
                    margin: 0 auto;
                    line-height: 1.6;
                }

                .markdown-content h1,
                .markdown-content h2,
                .markdown-content h3,
                .markdown-content h4,
                .markdown-content h5,
                .markdown-content h6 {
                    margin-top: 24px;
                    margin-bottom: 16px;
                    color: var(--vscode-foreground);
                }

                .markdown-content h1 {
                    border-bottom: 1px solid var(--vscode-sideBar-border);
                    padding-bottom: 8px;
                }

                .markdown-content p {
                    margin-bottom: 16px;
                }

                .markdown-content ul,
                .markdown-content ol {
                    margin-bottom: 16px;
                    padding-left: 24px;
                }

                .markdown-content li {
                    margin-bottom: 4px;
                }

                .markdown-content code {
                    background: var(--vscode-textCodeBlock-background);
                    padding: 2px 4px;
                    border-radius: 3px;
                    font-family: var(--vscode-editor-font-family);
                }

                .markdown-content pre {
                    background: var(--vscode-textCodeBlock-background);
                    padding: 16px;
                    border-radius: 4px;
                    overflow-x: auto;
                    margin-bottom: 16px;
                }

                .markdown-content blockquote {
                    border-left: 4px solid var(--vscode-button-background);
                    padding-left: 16px;
                    margin-bottom: 16px;
                    opacity: 0.8;
                }

                /* Modal Styles - Responsive */
                .modal {
                    position: fixed;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    background: rgba(0, 0, 0, 0.4);
                    display: none;
                    align-items: center;
                    justify-content: center;
                    z-index: 1000;
                    padding: 20px;
                }

                .modal-content {
                    background: var(--vscode-editor-background);
                    border: 1px solid var(--vscode-sideBar-border);
                    border-radius: 8px;
                    width: 100%;
                    max-width: 700px;
                    max-height: 90vh;
                    overflow: hidden;
                    display: flex;
                    flex-direction: column;
                    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
                }

                .modal-header {
                    padding: 20px 24px;
                    border-bottom: 1px solid var(--vscode-sideBar-border);
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    flex-shrink: 0;
                }

                .modal-title {
                    font-size: 18px;
                    font-weight: 600;
                    color: var(--vscode-foreground);
                }

                .close-button {
                    background: transparent;
                    border: none;
                    color: var(--vscode-foreground);
                    cursor: pointer;
                    padding: 8px;
                    border-radius: 4px;
                    font-size: 16px;
                    opacity: 0.7;
                    line-height: 1;
                    transition: all 0.2s ease;
                }

                .close-button:hover {
                    opacity: 1;
                    background: var(--vscode-toolbar-hoverBackground);
                }

                .modal-body {
                    padding: 24px;
                    overflow-y: auto;
                    flex: 1;
                }

                .modal-section {
                    margin-bottom: 32px;
                }

                .modal-section:last-child {
                    margin-bottom: 0;
                }

                .modal-section-title {
                    font-size: 16px;
                    font-weight: 600;
                    color: var(--vscode-foreground);
                    margin-bottom: 16px;
                    border-bottom: 1px solid var(--vscode-sideBar-border);
                    padding-bottom: 8px;
                }

                .modal-field {
                    margin-bottom: 20px;
                }

                .modal-label {
                    display: block;
                    font-size: 14px;
                    color: var(--vscode-foreground);
                    margin-bottom: 8px;
                    font-weight: 500;
                }

                .modal-input {
                    width: 100%;
                    padding: 10px 12px;
                    background: var(--vscode-input-background);
                    border: 1px solid var(--vscode-input-border);
                    color: var(--vscode-input-foreground);
                    border-radius: 6px;
                    font-family: var(--vscode-font-family);
                    font-size: 14px;
                    transition: border-color 0.2s ease;
                }

                .modal-input:focus {
                    outline: none;
                    border-color: var(--vscode-focusBorder);
                }

                .modal-note {
                    font-size: 12px;
                    color: var(--vscode-foreground);
                    opacity: 0.7;
                    margin-top: 6px;
                    line-height: 1.4;
                }

                .status-display {
                    padding: 12px 16px;
                    border-radius: 6px;
                    margin-bottom: 12px;
                    font-size: 14px;
                    font-weight: 500;
                    display: flex;
                    align-items: center;
                    gap: 8px;
                }

                .status-success {
                    background: rgba(0, 200, 0, 0.1);
                    color: #00c800;
                    border: 1px solid rgba(0, 200, 0, 0.3);
                }

                .status-error {
                    background: rgba(200, 0, 0, 0.1);
                    color: #c80000;
                    border: 1px solid rgba(200, 0, 0, 0.3);
                }

                .status-warning {
                    background: rgba(255, 165, 0, 0.1);
                    color: #ffa500;
                    border: 1px solid rgba(255, 165, 0, 0.3);
                }

                .template-list {
                    max-height: 200px;
                    overflow-y: auto;
                    border: 1px solid var(--vscode-input-border);
                    border-radius: 6px;
                    padding: 12px;
                    margin-bottom: 12px;
                    background: var(--vscode-input-background);
                }

                .template-item {
                    display: flex;
                    justify-content: space-between;
                    align-items: flex-start;
                    padding: 12px;
                    border-bottom: 1px solid var(--vscode-input-border);
                    font-size: 14px;
                }

                .template-item:last-child {
                    border-bottom: none;
                }

                .template-info {
                    flex: 1;
                    min-width: 0;
                }

                .template-name {
                    font-weight: 500;
                    color: var(--vscode-foreground);
                    margin-bottom: 4px;
                }

                .template-path {
                    color: var(--vscode-descriptionForeground);
                    font-size: 12px;
                    word-break: break-all;
                }

                .template-remove {
                    background: var(--vscode-errorForeground);
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 6px 12px;
                    font-size: 11px;
                    cursor: pointer;
                    flex-shrink: 0;
                    margin-left: 12px;
                    transition: opacity 0.2s ease;
                }

                .template-remove:hover {
                    opacity: 0.8;
                }

                .browse-button {
                    padding: 10px 16px;
                    background: var(--vscode-button-background);
                    color: var(--vscode-button-foreground);
                    border: none;
                    border-radius: 6px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 500;
                    margin-right: 12px;
                    transition: background-color 0.2s ease;
                }

                .browse-button:hover {
                    background: var(--vscode-button-hoverBackground);
                }

                .file-picker-container {
                    display: flex;
                    gap: 12px;
                    align-items: center;
                }

                .file-picker-input {
                    flex: 1;
                    background: var(--vscode-input-background);
                    border: 1px solid var(--vscode-input-border);
                    color: var(--vscode-input-foreground);
                    padding: 10px 12px;
                    border-radius: 6px;
                    font-family: var(--vscode-font-family);
                    font-size: 14px;
                }

                .modal-actions {
                    display: flex;
                    justify-content: flex-end;
                    gap: 12px;
                    margin-top: 24px;
                    padding-top: 20px;
                    border-top: 1px solid var(--vscode-sideBar-border);
                    flex-shrink: 0;
                }

                .cancel-button {
                    padding: 10px 20px;
                    background: transparent;
                    color: var(--vscode-foreground);
                    border: 1px solid var(--vscode-button-border);
                    border-radius: 6px;
                    cursor: pointer;
                    font-family: var(--vscode-font-family);
                    font-size: 14px;
                    transition: background-color 0.2s ease;
                }

                .cancel-button:hover {
                    background: var(--vscode-toolbar-hoverBackground);
                }

                .confirm-button {
                    padding: 10px 20px;
                    background: var(--vscode-button-background);
                    color: var(--vscode-button-foreground);
                    border: none;
                    border-radius: 6px;
                    cursor: pointer;
                    font-family: var(--vscode-font-family);
                    font-size: 14px;
                    font-weight: 500;
                    transition: background-color 0.2s ease;
                }

                .confirm-button:hover {
                    background: var(--vscode-button-hoverBackground);
                }

                .export-engine-display {
                    padding: 10px 12px;
                    background: var(--vscode-input-background);
                    border: 1px solid var(--vscode-input-border);
                    border-radius: 6px;
                    color: var(--vscode-input-foreground);
                    font-size: 14px;
                    font-family: var(--vscode-font-family);
                }

                /* Responsive Design */
                @media (max-width: 768px) {
                    .modal {
                        padding: 10px;
                    }
                    
                    .modal-content {
                        max-width: 100%;
                        max-height: 95vh;
                    }
                    
                    .modal-header {
                        padding: 16px 20px;
                    }
                    
                    .modal-title {
                        font-size: 16px;
                    }
                    
                    .modal-body {
                        padding: 20px;
                    }
                    
                    .file-picker-container {
                        flex-direction: column;
                        align-items: stretch;
                    }
                    
                    .browse-button {
                        margin-right: 0;
                        margin-bottom: 8px;
                    }
                }

                @media (max-height: 600px) {
                    .modal-content {
                        max-height: 95vh;
                    }
                    
                    .template-list {
                        max-height: 120px;
                    }
                }
                
                @keyframes spin {
                    0% { transform: rotate(0deg); }
                    100% { transform: rotate(360deg); }
                }
            </style>
        </head>
        <body>
            <div class="sidebar">
                <div class="panel-header">
                    <div class="header-title">
                        <div>📄</div>
                        <span>Doculate</span>
                    </div>
                    <button class="settings-button" onclick="openSettings()">⚙️</button>
                </div>
                
                <div class="sidebar-header">
                    <div class="file-selector">
                        <label>Select Document</label>
                        <select class="file-dropdown" id="fileDropdown">
                            <option value="">Choose a document...</option>
                        </select>
                    </div>
                </div>

                <div class="current-file-status" id="currentFileStatus" style="display: none;">
                    <div class="file-info-display">
                        <div class="file-icon">📄</div>
                        <div class="file-details">
                            <div class="file-name-display" id="currentFileName">No file selected</div>
                            <div class="file-type-display" id="currentFileType">Select a document to begin</div>
                            <div class="doculate-badge" id="doculateBadge" style="display: none;">Markdown</div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="main-content">
                <div class="preview-header">
                    <div class="preview-title">
                        <span>👁️</span>
                        <span id="previewTitle">Preview</span>
                    </div>
                    <div class="preview-actions">
                        <button class="export-button" onclick="openExportModal()" id="exportButton" disabled>
                            📄 Export to Word
                        </button>
                    </div>
                </div>

                <div class="preview-content">
                    <div class="empty-state" id="emptyState">
                        <h3>📄 Select a markdown file to preview</h3>
                        <p>Choose a .md file from your workspace to see the preview here</p>
                    </div>
                    <div class="markdown-content" id="markdownContent" style="display: none;"></div>
                </div>
            </div>

            <!-- Settings Modal -->
            <div class="modal" id="settingsModal">
                <div class="modal-content">
                    <div class="modal-header">
                        <div class="modal-title">Export Settings</div>
                        <button class="close-button" onclick="closeSettings()">✕</button>
                    </div>
                    <div class="modal-body">
                        <div class="modal-section">
                            <div class="modal-section-title">Export Engine</div>
                            
                            <div class="modal-field">
                                <label class="modal-label">Conversion Engine</label>
                                <select class="modal-input" id="exportEngine" onchange="updatePandocStatus()">
                                    <option value="pandoc">Pandoc (Recommended)</option>
                                    <option value="docx">Docx Library</option>
                                </select>
                                <div class="modal-note">Pandoc provides better formatting and supports Mermaid diagrams when Mermaid CLI is installed.</div>
                            </div>
                            
                            <div class="modal-field" id="pandocStatusSection" style="display: none;">
                                <label class="modal-label">Pandoc Status</label>
                                <div id="pandocStatus" class="status-display">Checking...</div>
                                <button class="browse-button" id="installPandocBtn" onclick="installPandoc()" style="display: none;">Install Pandoc</button>
                                <div class="modal-note">Pandoc is required for the best export quality and advanced formatting</div>
                            </div>
                        </div>

                        <div class="modal-section">
                            <div class="modal-section-title">Diagram Support</div>
                            
                            <div class="modal-field">
                                <label class="modal-label">Node.js Status</label>
                                <div id="nodejsStatus" class="status-display">Checking...</div>
                                <button class="browse-button" id="installNodejsBtn" onclick="installNodejs()" style="display: none;">Install Node.js</button>
                                <div class="modal-note">Node.js is required for Mermaid CLI installation and diagram rendering.</div>
                            </div>
                            
                            <div class="modal-field">
                                <label class="modal-label">Mermaid CLI Status</label>
                                <div id="mermaidStatus" class="status-display">Checking...</div>
                                <button class="browse-button" id="installMermaidBtn" onclick="installMermaid()" style="display: none;">Install Mermaid CLI</button>
                                <div class="modal-note">Mermaid CLI converts diagram code to images in Word documents.</div>
                            </div>
                        </div>

                        <div class="modal-section">
                            <div class="modal-section-title">Reference Documents</div>
                            
                            <div class="modal-field">
                                <label class="modal-label">Word Reference Documents</label>
                                <div class="template-list" id="templateList">
                                    <div class="template-item">
                                        <div class="template-info">
                                            <div class="template-name">Default Formatting</div>
                                            <div class="template-path">(Built-in)</div>
                                        </div>
                                    </div>
                                </div>
                                <button class="browse-button" onclick="addTemplate()">Add Reference Document</button>
                                <button class="browse-button" onclick="showTemplateData()" style="background: var(--vscode-button-secondaryBackground); margin-left: 8px;">Show Template Data</button>
                                <div class="modal-note">Add Word documents (.docx) to use as formatting references. Pandoc will copy the styles, headers, and formatting from these documents.</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Export Modal -->
            <div class="modal" id="exportModal">
                <div class="modal-content">
                    <div class="modal-header">
                        <div class="modal-title">Export to Word Document</div>
                        <button class="close-button" onclick="closeExportModal()">✕</button>
                    </div>
                    <div class="modal-body">
                        <div class="modal-field">
                            <label class="modal-label">Export Engine</label>
                            <div class="export-engine-display" id="exportEngineDisplay">
                                Pandoc (Recommended)
                            </div>
                            <div class="modal-note">Engine selected in settings</div>
                        </div>

                        <div class="modal-field">
                            <label class="modal-label">Document Name</label>
                            <input type="text" class="modal-input" id="exportFilename" placeholder="My Document" value="">
                            <div class="modal-note">The .docx extension will be added automatically</div>
                        </div>

                        <div class="modal-field">
                            <label class="modal-label">Reference Document</label>
                            <select class="modal-input" id="templateSelect">
                                <option value="">Default Formatting</option>
                            </select>
                            <div class="modal-note">Choose a Word document to copy formatting and styles from</div>
                        </div>

                        <div class="modal-field">
                            <label class="modal-label">Output Location</label>
                            <div class="file-picker-container">
                                <input type="text" class="file-picker-input" id="outputPath" placeholder="Workspace folder" readonly>
                                <button class="browse-button" onclick="browseOutputFolder()">Browse</button>
                            </div>
                            <div class="modal-note">Where to save the exported Word document</div>
                        </div>

                        <div class="modal-actions">
                            <button class="cancel-button" onclick="closeExportModal()">Cancel</button>
                            <button class="confirm-button" onclick="performExport()">Export Document</button>
                        </div>
                    </div>
                </div>
            </div>

            <script>
                const vscode = acquireVsCodeApi();
                let files = [];
                let selectedFile = null;

                // Load files on startup
                vscode.postMessage({ command: 'getMarkdownFiles' });

                // Modal Functions
                function openSettings() {
                    document.getElementById('settingsModal').style.display = 'flex';
                    
                    // Initialize export engine status display
                    updatePandocStatus();
                    
                    // Show checking indicators immediately
                    updatePandocStatusUI(false, true); // not installed, checking
                    updateMermaidStatusUI(false, false, true, true); // not installed, not node installed, checking both
                    
                    // Start status checks
                    vscode.postMessage({ command: 'checkPandocStatus' });
                    vscode.postMessage({ command: 'checkMermaidStatus' });
                }

                function closeSettings() {
                    document.getElementById('settingsModal').style.display = 'none';
                }

                function openExportModal() {
                    if (!selectedFile) {
                        vscode.postMessage({
                            command: 'showError',
                            message: 'Please select a document to export'
                        });
                        return;
                    }
                    
                    const defaultName = selectedFile.name.replace('.md', '');
                    document.getElementById('exportFilename').value = defaultName;
                    
                    const settingsEngine = document.getElementById('exportEngine');
                    const engineDisplay = document.getElementById('exportEngineDisplay');
                    if (settingsEngine && engineDisplay) {
                        const engineText = settingsEngine.options[settingsEngine.selectedIndex].text;
                        engineDisplay.textContent = engineText;
                    }
                    
                    vscode.postMessage({ command: 'getWorkspacePath' });
                    document.getElementById('exportModal').style.display = 'flex';
                }

                function closeExportModal() {
                    document.getElementById('exportModal').style.display = 'none';
                    document.getElementById('exportFilename').value = '';
                    document.getElementById('outputPath').value = '';
                }

                function installPandoc() {
                    vscode.postMessage({ command: 'installPandoc' });
                }

                function installNodejs() {
                    vscode.postMessage({ command: 'installNodejs' });
                }

                function installMermaid() {
                    vscode.postMessage({ command: 'installMermaid' });
                }

                function updatePandocStatus() {
                    const engine = document.getElementById('exportEngine').value;
                    const statusSection = document.getElementById('pandocStatusSection');
                    
                    if (engine === 'pandoc') {
                        statusSection.style.display = 'block';
                        vscode.postMessage({ command: 'checkPandocStatus' });
                    } else {
                        statusSection.style.display = 'none';
                    }
                }

                function addTemplate() {
                    vscode.postMessage({ command: 'addTemplate' });
                }

                function showTemplateData() {
                    vscode.postMessage({ command: 'showTemplateData' });
                }

                function removeTemplate(templateName) {
                    vscode.postMessage({
                        command: 'removeTemplate',
                        templateName: templateName
                    });
                }

                function browseOutputFolder() {
                    vscode.postMessage({ command: 'browseOutputFolder' });
                }

                function performExport() {
                    const settingsEngine = document.getElementById('exportEngine');
                    const engine = settingsEngine ? settingsEngine.value : 'pandoc';
                    
                    const filename = document.getElementById('exportFilename').value.trim();
                    const templateSelect = document.getElementById('templateSelect');
                    const selectedTemplate = templateSelect.value;
                    const outputPath = document.getElementById('outputPath').value.trim();
                    
                    if (!filename) {
                        vscode.postMessage({
                            command: 'showError',
                            message: 'Please enter a document name'
                        });
                        return;
                    }
                    
                    if (!outputPath) {
                        vscode.postMessage({
                            command: 'showError',
                            message: 'Please select an output folder'
                        });
                        return;
                    }
                    
                    vscode.postMessage({
                        command: 'exportToWord',
                        data: {
                            engine: engine,
                            filename: filename,
                            templatePath: selectedTemplate || null,
                            outputPath: outputPath,
                            sourceFile: selectedFile
                        }
                    });
                    
                    closeExportModal();
                }

                // Close modals when clicking outside
                document.getElementById('settingsModal').addEventListener('click', function(e) {
                    if (e.target === this) closeSettings();
                });

                document.getElementById('exportModal').addEventListener('click', function(e) {
                    if (e.target === this) closeExportModal();
                });

                // Handle dropdown selection
                document.getElementById('fileDropdown').addEventListener('change', (e) => {
                    if (e.target.value !== '') {
                        selectFile(parseInt(e.target.value));
                    }
                });

                function selectFile(index) {
                    selectedFile = files[index];
                    
                    document.getElementById('fileDropdown').value = index;
                    document.getElementById('previewTitle').textContent = \`Preview: \${selectedFile.name}\`;
                    
                    const currentFileStatus = document.getElementById('currentFileStatus');
                    const currentFileName = document.getElementById('currentFileName');
                    const currentFileType = document.getElementById('currentFileType');
                    const doculateBadge = document.getElementById('doculateBadge');
                    
                    currentFileStatus.style.display = 'block';
                    currentFileName.textContent = selectedFile.name;
                    currentFileType.textContent = 'Markdown Document';
                    doculateBadge.style.display = 'inline-block';
                    
                    vscode.postMessage({ 
                        command: 'previewFile', 
                        filePath: selectedFile.path 
                    });
                }

                function updatePreview(htmlContent, fileName, isProcessing = false, mermaidCount = 0) {
                    const emptyState = document.getElementById('emptyState');
                    const markdownContent = document.getElementById('markdownContent');
                    const exportButton = document.getElementById('exportButton');
                    
                    emptyState.style.display = 'none';
                    markdownContent.style.display = 'block';
                    
                    // Create or update loading banner
                    let loadingBanner = document.getElementById('mermaidLoadingBanner');
                    
                    if (isProcessing && mermaidCount > 0) {
                        if (!loadingBanner) {
                            loadingBanner = document.createElement('div');
                            loadingBanner.id = 'mermaidLoadingBanner';
                            loadingBanner.style.cssText = 'position: fixed; bottom: 20px; right: 20px; z-index: 1000; background: var(--vscode-editor-inactiveSelectionBackground); border: 1px solid var(--vscode-progressBar-background); padding: 10px 14px; border-radius: 6px; display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--vscode-foreground); box-shadow: 0 4px 12px rgba(0,0,0,0.2); backdrop-filter: blur(4px); max-width: 280px;';
                            document.body.appendChild(loadingBanner);
                        }
                        
                        loadingBanner.innerHTML = '<div style="width: 14px; height: 14px; border: 2px solid var(--vscode-progressBar-background); border-top: 2px solid transparent; border-radius: 50%; animation: spin 1s linear infinite; flex-shrink: 0;"></div><span style="white-space: nowrap;">Processing ' + mermaidCount + ' diagram' + (mermaidCount > 1 ? 's' : '') + '...</span>';
                    } else if (loadingBanner) {
                        // Remove loading banner when not processing
                        loadingBanner.remove();
                    }
                    
                    markdownContent.innerHTML = htmlContent;
                    exportButton.disabled = false;
                }

                function updateFileList(fileList) {
                    files = fileList;
                    const fileDropdown = document.getElementById('fileDropdown');
                    
                    fileDropdown.innerHTML = '<option value="">Choose a document...</option>';
                    
                    files.forEach((file, index) => {
                        const option = document.createElement('option');
                        option.value = index;
                        option.textContent = file.name;
                        fileDropdown.appendChild(option);
                    });
                }

                function updatePandocStatusUI(isInstalled, isChecking = false) {
                    const statusDiv = document.getElementById('pandocStatus');
                    const installBtn = document.getElementById('installPandocBtn');
                    
                    // Clear existing classes
                    statusDiv.className = 'status-display';
                    
                    if (isChecking) {
                        statusDiv.innerHTML = '<div style="width: 14px; height: 14px; border: 2px solid var(--vscode-progressBar-background); border-top: 2px solid transparent; border-radius: 50%; animation: spin 1s linear infinite; display: inline-block; margin-right: 8px;"></div>Checking Pandoc installation...';
                        statusDiv.classList.add('status-warning');
                        installBtn.style.display = 'none';
                    } else if (isInstalled) {
                        statusDiv.innerHTML = '✅ Pandoc is installed and ready (Latest version recommended)';
                        statusDiv.classList.add('status-success');
                        installBtn.style.display = 'none';
                    } else {
                        statusDiv.innerHTML = '❌ Pandoc is not installed (Required for advanced Word conversion)';
                        statusDiv.classList.add('status-error');
                        installBtn.style.display = 'inline-block';
                    }
                }

                function updateMermaidStatusUI(isInstalled, nodeInstalled, isCheckingNode = false, isCheckingMermaid = false) {
                    const nodejsStatusDiv = document.getElementById('nodejsStatus');
                    const installNodejsBtn = document.getElementById('installNodejsBtn');
                    const mermaidStatusDiv = document.getElementById('mermaidStatus');
                    const installMermaidBtn = document.getElementById('installMermaidBtn');
                    
                    // Clear existing classes
                    nodejsStatusDiv.className = 'status-display';
                    mermaidStatusDiv.className = 'status-display';
                    
                    // Update Node.js status
                    if (isCheckingNode) {
                        nodejsStatusDiv.innerHTML = '<div style="width: 14px; height: 14px; border: 2px solid var(--vscode-progressBar-background); border-top: 2px solid transparent; border-radius: 50%; animation: spin 1s linear infinite; display: inline-block; margin-right: 8px;"></div>Checking Node.js installation...';
                        nodejsStatusDiv.classList.add('status-warning');
                        installNodejsBtn.style.display = 'none';
                    } else if (nodeInstalled) {
                        nodejsStatusDiv.innerHTML = '✅ Node.js is installed and ready (v18+ recommended)';
                        nodejsStatusDiv.classList.add('status-success');
                        installNodejsBtn.style.display = 'none';
                    } else {
                        nodejsStatusDiv.innerHTML = '❌ Node.js is not installed (Required for Mermaid diagrams)';
                        nodejsStatusDiv.classList.add('status-error');
                        installNodejsBtn.style.display = 'inline-block';
                    }
                    
                    // Update Mermaid CLI status
                    if (isCheckingMermaid) {
                        mermaidStatusDiv.innerHTML = '<div style="width: 14px; height: 14px; border: 2px solid var(--vscode-progressBar-background); border-top: 2px solid transparent; border-radius: 50%; animation: spin 1s linear infinite; display: inline-block; margin-right: 8px;"></div>Checking Mermaid CLI installation...';
                        mermaidStatusDiv.classList.add('status-warning');
                        installMermaidBtn.style.display = 'none';
                    } else if (isInstalled) {
                        mermaidStatusDiv.innerHTML = '✅ Mermaid CLI is installed and ready (v11.6.0+)';
                        mermaidStatusDiv.classList.add('status-success');
                        installMermaidBtn.style.display = 'none';
                    } else if (!nodeInstalled) {
                        mermaidStatusDiv.innerHTML = '⚠️ Requires Node.js installation first';
                        mermaidStatusDiv.classList.add('status-warning');
                        installMermaidBtn.style.display = 'none';
                    } else {
                        mermaidStatusDiv.innerHTML = '❌ Mermaid CLI is not installed (Required for diagram rendering)';
                        mermaidStatusDiv.classList.add('status-error');
                        installMermaidBtn.style.display = 'inline-block';
                    }
                }

                function loadSavedTemplates(templates) {
                    const templateList = document.getElementById('templateList');
                    const templateSelect = document.getElementById('templateSelect');
                    
                    // Clear existing custom templates
                    const items = templateList.querySelectorAll('.template-item');
                    items.forEach(item => {
                        const pathSpan = item.querySelector('.template-path');
                        if (pathSpan && !pathSpan.textContent.includes('(Built-in)')) {
                            item.remove();
                        }
                    });
                    
                    const options = templateSelect.querySelectorAll('option');
                    options.forEach(option => {
                        if (option.value !== '') {
                            option.remove();
                        }
                    });
                    
                    // Add saved templates
                    templates.forEach(template => {
                        const templateItem = document.createElement('div');
                        templateItem.className = 'template-item';
                        templateItem.innerHTML = \`
                            <div class="template-info">
                                <div class="template-name">\${template.name}</div>
                                <div class="template-path">\${template.path}</div>
                            </div>
                            <button class="template-remove" onclick="removeTemplate('\${template.name}')">Remove</button>
                        \`;
                        templateList.appendChild(templateItem);
                        
                        const option = document.createElement('option');
                        option.value = template.path;
                        option.textContent = template.name;
                        templateSelect.appendChild(option);
                    });
                }

                // Handle messages from extension
                window.addEventListener('message', event => {
                    const message = event.data;
                    
                    switch (message.command) {
                        case 'updateFiles':
                            updateFileList(message.files);
                            break;
                        case 'updatePreview':
                            updatePreview(message.content, message.fileName, message.isProcessing, message.mermaidCount);
                            break;
                        case 'previewProcessingComplete':
                            // Remove loading indicator if processing failed
                            if (message.error) {
                                const loadingBanner = document.getElementById('mermaidLoadingBanner');
                                if (loadingBanner) {
                                    loadingBanner.remove();
                                }
                            }
                            break;
                        case 'outputFolderSelected':
                            document.getElementById('outputPath').value = message.path;
                            break;
                        case 'updatePandocStatus':
                            updatePandocStatusUI(message.isInstalled, false);
                            break;
                        case 'updateMermaidStatus':
                            updateMermaidStatusUI(message.isInstalled, message.nodeInstalled, false, false);
                            break;
                        case 'loadSavedTemplates':
                            loadSavedTemplates(message.templates);
                            break;
                        case 'workspacePathSelected':
                            document.getElementById('outputPath').value = message.path;
                            break;
                        case 'exportComplete':
                            // Export completed successfully
                            break;
                    }
                });
            </script>
        </body>
        </html>`;
    }

    public dispose() {
        DoculatePanel.currentPanel = undefined;
        this._panel.dispose();
        
        // Clear Mermaid cache
        this._mermaidCache.clear();
        
        while (this._disposables.length) {
            const disposable = this._disposables.pop();
            disposable?.dispose();
        }
    }
}

export function deactivate() {}