import * as vscode from "vscode";

export function activate(context: vscode.ExtensionContext) {
  console.log('Extension "markdown-tables" is now active');

  const provider = new MarkdownTableProvider(context.extensionUri);

  const disposables = [
    vscode.commands.registerCommand("markdown-tables.editTable", (args) => {
      provider.editTable(args);
    }),
    vscode.commands.registerCommand("markdown-tables.createTable", () => {
      provider.createTable();
    }),
    vscode.commands.registerCommand("markdown-tables.sortTableAsc", (args) => {
      provider.sortTable(args, true);
    }),
    vscode.commands.registerCommand("markdown-tables.sortTableDesc", (args) => {
      provider.sortTable(args, false);
    }),
    vscode.languages.registerCodeLensProvider(
      ["markdown", "mdx"],
      new MarkdownTableCodeLensProvider()
    ),
  ];

  context.subscriptions.push(...disposables);
}

class MarkdownTableCodeLensProvider implements vscode.CodeLensProvider {
  provideCodeLenses(
    document: vscode.TextDocument,
    token: vscode.CancellationToken
  ): vscode.CodeLens[] | Thenable<vscode.CodeLens[]> {
    const codeLenses: vscode.CodeLens[] = [];
    const text = document.getText();
    const lines = text.split("\n");

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      if (this.isTableRow(line)) {
        const tableStart = this.findTableStart(lines, i);
        const tableEnd = this.findTableEnd(lines, i);

        if (tableStart !== -1 && tableEnd !== -1) {
          const range = new vscode.Range(tableStart, 0, tableStart, 0);
          
          const editCommand: vscode.Command = {
            title: "Edit",
            command: "markdown-tables.editTable",
            arguments: [
              {
                document: document,
                startLine: tableStart,
                endLine: tableEnd,
              },
            ],
          };
          codeLenses.push(new vscode.CodeLens(range, editCommand));

          const sortAscCommand: vscode.Command = {
            title: "Sort (A-Z)",
            command: "markdown-tables.sortTableAsc",
            arguments: [
              {
                document: document,
                startLine: tableStart,
                endLine: tableEnd,
              },
            ],
          };
          codeLenses.push(new vscode.CodeLens(range, sortAscCommand));

          const sortDescCommand: vscode.Command = {
            title: "Sort (Z-A)",
            command: "markdown-tables.sortTableDesc",
            arguments: [
              {
                document: document,
                startLine: tableStart,
                endLine: tableEnd,
              },
            ],
          };
          codeLenses.push(new vscode.CodeLens(range, sortDescCommand));

          i = tableEnd;
        }
      }
    }

    return codeLenses;
  }

  private isTableRow(line: string): boolean {
    const trimmed = line.trim();
    return trimmed.includes("|") && trimmed.length > 2;
  }

  private findTableStart(lines: string[], currentLine: number): number {
    for (let i = currentLine; i >= 0; i--) {
      if (!this.isTableRow(lines[i])) {
        return i + 1;
      }
    }
    return 0;
  }

  private findTableEnd(lines: string[], currentLine: number): number {
    for (let i = currentLine; i < lines.length; i++) {
      if (!this.isTableRow(lines[i])) {
        return i - 1;
      }
    }
    return lines.length - 1;
  }
}

class MarkdownTableProvider {
  public static currentPanel: MarkdownTableProvider | undefined;
  private readonly _extensionUri: vscode.Uri;
  private _panel: vscode.WebviewPanel | undefined;

  constructor(extensionUri: vscode.Uri) {
    this._extensionUri = extensionUri;
  }

  public editTable(args: any) {
    const { document, startLine, endLine } = args;
    const tableText = this.extractTableText(document, startLine, endLine);
    const tableData = this.parseMarkdownTable(tableText);

    if (this._panel) {
      this._panel.dispose();
    }

    this._panel = vscode.window.createWebviewPanel(
      "markdownTableEditor",
      "Edit Markdown Table",
      vscode.ViewColumn.Beside,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
      }
    );

    this._panel.webview.html = this.getWebviewContent(tableData);

    this._panel.webview.onDidReceiveMessage((message) => {
      switch (message.command) {
        case "save":
          this.saveTable(document, startLine, endLine, message.data);
          this._panel?.dispose();
          break;
        case "cancel":
          this._panel?.dispose();
          break;
      }
    });
  }

  public createTable() {
    const activeEditor = vscode.window.activeTextEditor;
    if (!activeEditor) {
      vscode.window.showErrorMessage("No active editor found");
      return;
    }

    const document = activeEditor.document;
    if (!["markdown", "mdx"].includes(document.languageId)) {
      vscode.window.showErrorMessage(
        "This command only works in Markdown files"
      );
      return;
    }

    const defaultTableData = [
      ["", "", ""],
      ["", "", ""],
    ];

    if (this._panel) {
      this._panel.dispose();
    }

    this._panel = vscode.window.createWebviewPanel(
      "markdownTableEditor",
      "Create Markdown Table",
      vscode.ViewColumn.Beside,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
      }
    );

    this._panel.webview.html = this.getWebviewContent(defaultTableData);

    this._panel.webview.onDidReceiveMessage((message) => {
      switch (message.command) {
        case "save":
          this.insertNewTable(activeEditor, message.data);
          this._panel?.dispose();
          break;
        case "cancel":
          this._panel?.dispose();
          break;
      }
    });
  }

  public sortTable(args: any, ascending: boolean) {
    const { document, startLine, endLine } = args;
    const tableText = this.extractTableText(document, startLine, endLine);
    const tableData = this.parseMarkdownTable(tableText);

    if (tableData.length <= 1) {
      vscode.window.showInformationMessage("Table has no data rows to sort");
      return;
    }

    const headerRow = tableData[0];
    const dataRows = tableData.slice(1);

    const sortedDataRows = [...dataRows].sort((a, b) => {
      const aValue = (a[0] || "").trim().toLowerCase();
      const bValue = (b[0] || "").trim().toLowerCase();
      
      if (ascending) {
        return aValue.localeCompare(bValue);
      } else {
        return bValue.localeCompare(aValue);
      }
    });

    const sortedTableData = [headerRow, ...sortedDataRows];
    this.saveTable(document, startLine, endLine, sortedTableData);
  }

  private extractTableText(
    document: vscode.TextDocument,
    startLine: number,
    endLine: number
  ): string {
    const lines: string[] = [];
    for (let i = startLine; i <= endLine; i++) {
      lines.push(document.lineAt(i).text);
    }
    return lines.join("\n");
  }

  private parseMarkdownTable(tableText: string): string[][] {
    const lines = tableText.split("\n").filter((line) => line.trim());
    const rows: string[][] = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      if (i === 1 && this.isSeparatorRow(line)) {
        continue;
      }

      const cells = line
        .split("|")
        .map((cell) => cell.trim())
        .filter((cell, index, array) => {
          return (
            !(index === 0 && cell === "") &&
            !(index === array.length - 1 && cell === "")
          );
        });

      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    return rows;
  }

  private isSeparatorRow(line: string): boolean {
    return /^[\s\|:\-]+$/.test(line);
  }

  private saveTable(
    document: vscode.TextDocument,
    startLine: number,
    endLine: number,
    tableData: string[][]
  ) {
    const markdownTable = this.generateMarkdownTable(tableData);

    const edit = new vscode.WorkspaceEdit();
    const range = new vscode.Range(startLine, 0, endLine + 1, 0);
    edit.replace(document.uri, range, markdownTable + "\n");

    vscode.workspace.applyEdit(edit);
  }

  private insertNewTable(editor: vscode.TextEditor, tableData: string[][]) {
    const markdownTable = this.generateMarkdownTable(tableData);
    const position = editor.selection.active;

    const currentLine = editor.document.lineAt(position.line);
    const isLineEmpty = currentLine.text.trim() === "";

    let insertText = markdownTable;
    if (!isLineEmpty) {
      insertText = "\n\n" + markdownTable + "\n\n";
    } else {
      insertText = markdownTable + "\n\n";
    }

    editor.edit((editBuilder) => {
      if (isLineEmpty) {
        editBuilder.replace(
          new vscode.Range(
            position.line,
            0,
            position.line,
            currentLine.text.length
          ),
          insertText
        );
      } else {
        editBuilder.insert(position, insertText);
      }
    });
  }

  private generateMarkdownTable(data: string[][]): string {
    if (data.length === 0) return "";

    const maxCols = Math.max(...data.map((row) => row.length));

    const normalizedData = data.map((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push("");
      }
      return normalizedRow;
    });

    const colWidths = new Array(maxCols).fill(0);
    for (const row of normalizedData) {
      for (let i = 0; i < row.length; i++) {
        colWidths[i] = Math.max(colWidths[i], row[i].length);
      }
    }

    const lines: string[] = [];

    for (let rowIndex = 0; rowIndex < normalizedData.length; rowIndex++) {
      const row = normalizedData[rowIndex];
      const paddedCells = row.map((cell, i) => cell.padEnd(colWidths[i]));
      lines.push(`| ${paddedCells.join(" | ")} |`);

      if (rowIndex === 0) {
        const separatorCells = colWidths.map((width) => "-".repeat(width));
        lines.push(`| ${separatorCells.join(" | ")} |`);
      }
    }

    return lines.join("\n");
  }

  private getWebviewContent(tableData: string[][]): string {
    const tableDataJson = JSON.stringify(tableData);

    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Edit Markdown Table</title>
    <style>
        * {
            box-sizing: border-box;
        }
        
        body {
            font-family: var(--vscode-font-family);
            color: var(--vscode-foreground);
            background-color: var(--vscode-editor-background);
            padding: 24px;
            margin: 0;
            line-height: 1.5;
        }
        
        h2 {
            margin: 0 0 24px 0;
            font-size: 20px;
            font-weight: 600;
            color: var(--vscode-foreground);
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
        }
        
        .controls-section {
            margin-bottom: 24px;
        }
        
        .row-controls {
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
        }
        
        .row-controls button {
            background-color: var(--vscode-button-background);
            color: var(--vscode-button-foreground);
            border: none;
            padding: 8px 12px;
            border-radius: 5px;
            cursor: pointer;
            font-family: inherit;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.2s ease;
            min-width: auto;
        }
        
        .row-controls button:hover {
            background-color: var(--vscode-button-hoverBackground);
        }
        
        .word-wrap-toggle {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-left: 16px;
            padding: 8px 12px;
            background-color: var(--vscode-input-background);
            border-radius: 5px;
            border: 1px solid var(--vscode-input-border);
        }
        
        .word-wrap-toggle label {
            font-size: 14px;
            color: var(--vscode-foreground);
            cursor: pointer;
            user-select: none;
        }
        
        .word-wrap-toggle input[type="checkbox"] {
            width: auto;
            margin: 0;
            cursor: pointer;
        }
        
        .table-section {
            margin-bottom: 24px;
        }
        
        .table-container {
            overflow-x: auto;
            border-radius: 8px;
            border: 1px solid var(--vscode-panel-border);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }
        
        table {
            border-collapse: collapse;
            width: 100%;
            min-width: 500px;
            background-color: var(--vscode-editor-background);
            table-layout: auto;
        }
        
        th, td {
            border: 1px solid var(--vscode-panel-border);
            text-align: left;
            position: relative;
            vertical-align: top;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        
        th {
            padding: 12px 16px;
            background-color: var(--vscode-editor-background);
            font-weight: 400;
            font-size: 14px;
            border-bottom: 1px solid var(--vscode-panel-border);
            height: auto;
            min-width: 120px;
            max-width: none;
        }
        
        td {
            padding: 12px 16px;
            background-color: var(--vscode-editor-background);
            transition: background-color 0.2s ease;
            min-width: 120px;
            max-width: none;
        }
        
        .word-wrap-enabled th,
        .word-wrap-enabled td {
            max-width: 300px;
            white-space: normal;
            word-break: break-word;
        }
        
        .word-wrap-disabled th,
        .word-wrap-disabled td {
            white-space: nowrap;
            overflow: hidden;
        }
        
        tr.selected {
            background-color: var(--vscode-list-activeSelectionBackground) !important;
        }
        
        tr.selected th,
        tr.selected td {
            background-color: var(--vscode-list-activeSelectionBackground) !important;
        }
        
        th.selected,
        td.selected {
            background-color: var(--vscode-list-activeSelectionBackground) !important;
        }
        
        input, textarea {
            width: 100%;
            background: transparent;
            border: none;
            color: var(--vscode-foreground);
            font-family: inherit;
            font-size: 14px;
            padding: 4px;
            border-radius: 3px;
            transition: all 0.2s ease;
            resize: none;
            overflow: hidden;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        
        input:focus, textarea:focus {
            outline: none;
            background-color: var(--vscode-input-background);
            box-shadow: 0 0 0 2px var(--vscode-focusBorder);
        }
        
        .word-wrap-enabled input,
        .word-wrap-enabled textarea {
            white-space: normal;
            word-break: break-word;
            min-height: 20px;
        }
        
        .word-wrap-disabled input,
        .word-wrap-disabled textarea {
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        
        .drag-handle {
            position: absolute;
            left: 4px;
            top: 50%;
            transform: translateY(-50%);
            cursor: grab;
            color: var(--vscode-descriptionForeground);
            font-size: 14px;
            opacity: 0.5;
            user-select: none;
            padding: 2px;
            border-radius: 2px;
            transition: all 0.2s ease;
        }
        
        .drag-handle:hover {
            opacity: 1;
        }
        
        .drag-handle:active {
            cursor: grabbing;
            background-color: var(--vscode-list-activeSelectionBackground);
        }
        
        .column-drag-handle {
            position: absolute;
            top: 2px;
            left: 50%;
            transform: translateX(-50%);
            cursor: grab;
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
            opacity: 0.5;
            user-select: none;
            padding: 2px;
            border-radius: 2px;
            transition: all 0.2s ease;
        }
        
        .column-drag-handle:hover {
            opacity: 1;
        }
        
        .column-drag-handle:active {
            cursor: grabbing;
            background-color: var(--vscode-list-activeSelectionBackground);
        }
        
        tr.dragging {
            opacity: 0.6;
            background-color: var(--vscode-list-activeSelectionBackground);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            transform: scale(1.02);
        }
        
        th.dragging, td.dragging {
            opacity: 0.6;
            background-color: var(--vscode-list-activeSelectionBackground);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        }
        
        .drop-indicator {
            border: 2px dashed var(--vscode-focusBorder);
            background-color: var(--vscode-list-hoverBackground);
        }
        
        .actions-section {
            display: flex;
            gap: 12px;
            justify-content: flex-end;
            padding-top: 20px;
            border-top: 1px solid var(--vscode-panel-border);
        }
        
        .actions-section button {
            background-color: var(--vscode-button-background);
            color: var(--vscode-button-foreground);
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            cursor: pointer;
            font-family: inherit;
            font-size: 14px;
            font-weight: 500;
            min-width: 80px;
            transition: all 0.2s ease;
        }
        
        .actions-section button:hover {
            background-color: var(--vscode-button-hoverBackground);
        }
        
        .actions-section .secondary {
            background-color: var(--vscode-button-secondaryBackground);
            color: var(--vscode-button-secondaryForeground);
        }
        
        .actions-section .secondary:hover {
            background-color: var(--vscode-button-secondaryHoverBackground);
        }
        
        .cell-wrapper {
            position: relative;
            display: flex;
            align-items: center;
            min-height: 24px;
        }
        
        .selection-info {
            margin-bottom: 12px;
            font-size: 12px;
            color: var(--vscode-descriptionForeground);
            font-style: italic;
        }
        
        @media (max-width: 768px) {
            body {
                padding: 16px;
            }
            
            .row-controls {
                gap: 6px;
            }
            
            .word-wrap-toggle {
                margin-left: 0;
            }
            
            .row-controls button {
                padding: 8px 12px;
                font-size: 14px;
            }
            
            th, td {
                padding: 8px 12px;
            }
            
            input, textarea {
                font-size: 13px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="controls-section">
            <div class="selection-info" id="selectionInfo"></div>
            <div class="row-controls">
                <button onclick="addRow()">Add Row</button>
                <button onclick="removeLastRow()">Remove Row</button>
                <button onclick="addColumn()">Add Column</button>
                <button onclick="removeLastColumn()">Remove Column</button>
                <div class="word-wrap-toggle">
                    <input type="checkbox" id="wordWrapToggle" checked onchange="toggleWordWrap()">
                    <label for="wordWrapToggle">Word Wrap</label>
                </div>
            </div>
        </div>
        
        <div class="table-section">
            <div class="table-container">
                <table id="editableTable" class="word-wrap-enabled">
                </table>
            </div>
        </div>
        
        <div class="actions-section">
            <button class="secondary" onclick="cancel()">Cancel</button>
            <button onclick="saveTable()">Save</button>
        </div>
    </div>

    <script>
        const vscode = acquireVsCodeApi();
        let tableData = ${tableDataJson};
        let draggedRowIndex = -1;
        let draggedColumnIndex = -1;
        let selectedRowIndex = -1;
        let selectedColumnIndex = -1;
        let wordWrapEnabled = true;
        
        function toggleWordWrap() {
            wordWrapEnabled = !wordWrapEnabled;
            const table = document.getElementById('editableTable');
            const checkbox = document.getElementById('wordWrapToggle');
            
            if (wordWrapEnabled) {
                table.className = 'word-wrap-enabled';
                checkbox.checked = true;
            } else {
                table.className = 'word-wrap-disabled';
                checkbox.checked = false;
            }
            
            adjustInputTypes();
        }
        
        function adjustInputTypes() {
            const table = document.getElementById('editableTable');
            const inputs = table.querySelectorAll('input, textarea');
            
            inputs.forEach(input => {
                const value = input.value;
                const cell = input.parentElement;
                
                if (wordWrapEnabled && (value.length > 50 || value.includes('\\n'))) {
                    if (input.tagName === 'INPUT') {
                        const textarea = document.createElement('textarea');
                        textarea.value = value;
                        textarea.style.paddingLeft = input.style.paddingLeft;
                        textarea.rows = Math.max(1, Math.ceil(value.length / 50));
                        
                        const rowIndex = parseInt(cell.parentElement.dataset.rowIndex);
                        const cellIndex = parseInt(cell.dataset.columnIndex);
                        
                        textarea.addEventListener('input', (e) => {
                            tableData[rowIndex][cellIndex] = e.target.value;
                            autoResizeTextarea(e.target);
                        });
                        
                        textarea.addEventListener('focus', () => {
                            clearSelection();
                        });
                        
                        cell.replaceChild(textarea, input);
                        autoResizeTextarea(textarea);
                    }
                } else if (!wordWrapEnabled && input.tagName === 'TEXTAREA') {
                    const newInput = document.createElement('input');
                    newInput.type = 'text';
                    newInput.value = value;
                    newInput.style.paddingLeft = input.style.paddingLeft;
                    
                    const rowIndex = parseInt(cell.parentElement.dataset.rowIndex);
                    const cellIndex = parseInt(cell.dataset.columnIndex);
                    
                    newInput.addEventListener('input', (e) => {
                        tableData[rowIndex][cellIndex] = e.target.value;
                    });
                    
                    newInput.addEventListener('focus', () => {
                        clearSelection();
                    });
                    
                    cell.replaceChild(newInput, input);
                }
            });
        }
        
        function autoResizeTextarea(textarea) {
            textarea.style.height = 'auto';
            textarea.style.height = Math.max(20, textarea.scrollHeight) + 'px';
        }
        
        function clearSelection() {
            selectedRowIndex = -1;
            selectedColumnIndex = -1;
            
            const table = document.getElementById('editableTable');
            const rows = table.querySelectorAll('tr');
            rows.forEach(row => row.classList.remove('selected'));
            
            const cells = table.querySelectorAll('th, td');
            cells.forEach(cell => cell.classList.remove('selected'));
        }
        
        function selectRow(rowIndex) {
            clearSelection();
            selectedRowIndex = rowIndex;
            
            const table = document.getElementById('editableTable');
            const row = table.querySelector(\`tr[data-row-index="\${rowIndex}"]\`);
            if (row) {
                row.classList.add('selected');
            }
        }
        
        function selectColumn(columnIndex) {
            clearSelection();
            selectedColumnIndex = columnIndex;
            
            const table = document.getElementById('editableTable');
            const rows = table.querySelectorAll('tr');
            rows.forEach(row => {
                const cell = row.children[columnIndex];
                if (cell) {
                    cell.classList.add('selected');
                }
            });
        }
        
        function deleteSelectedRow() {
            if (selectedRowIndex !== -1 && tableData.length > 1) {
                if (selectedRowIndex === 0) {
                    alert('Cannot delete header row');
                    return;
                }
                tableData.splice(selectedRowIndex, 1);
                clearSelection();
                renderTable();
            }
        }
        
        function deleteSelectedColumn() {
            if (selectedColumnIndex !== -1 && tableData.length > 0 && tableData[0].length > 1) {
                tableData.forEach(row => row.splice(selectedColumnIndex, 1));
                clearSelection();
                renderTable();
            }
        }
        
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Backspace' && !e.target.matches('input') && !e.target.matches('textarea')) {
                e.preventDefault();
                if (selectedRowIndex !== -1) {
                    deleteSelectedRow();
                } else if (selectedColumnIndex !== -1) {
                    deleteSelectedColumn();
                }
            } else if (e.key === 'Escape') {
                clearSelection();
            }
        });
        
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.drag-handle') && !e.target.closest('th') && !e.target.closest('.column-drag-handle')) {
                clearSelection();
            }
        });
        
        function renderTable() {
            const table = document.getElementById('editableTable');
            table.innerHTML = '';
            
            if (tableData.length === 0) {
                tableData = [['Header 1', 'Header 2'], ['Cell 1', 'Cell 2']];
            }
            
            tableData.forEach((row, rowIndex) => {
                const tr = document.createElement('tr');
                tr.draggable = true;
                tr.dataset.rowIndex = rowIndex;
                
                tr.addEventListener('dragstart', (e) => {
                    draggedRowIndex = rowIndex;
                    tr.classList.add('dragging');
                    e.dataTransfer.effectAllowed = 'move';
                });
                
                tr.addEventListener('dragend', (e) => {
                    tr.classList.remove('dragging');
                    draggedRowIndex = -1;
                });
                
                tr.addEventListener('dragover', (e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                });
                
                tr.addEventListener('drop', (e) => {
                    e.preventDefault();
                    if (draggedRowIndex !== -1 && draggedRowIndex !== rowIndex) {
                        moveRow(draggedRowIndex, rowIndex);
                    }
                });
                
                row.forEach((cell, cellIndex) => {
                    const cellElement = rowIndex === 0 ? document.createElement('th') : document.createElement('td');
                    cellElement.dataset.columnIndex = cellIndex;
                    
                    if (rowIndex === 0) {
                        cellElement.draggable = true;
                        
                        cellElement.addEventListener('click', (e) => {
                            if (!e.target.matches('input') && !e.target.matches('textarea')) {
                                e.stopPropagation();
                                selectColumn(cellIndex);
                            }
                        });
                        
                        cellElement.addEventListener('dragstart', (e) => {
                            draggedColumnIndex = cellIndex;
                            markColumnDragging(cellIndex, true);
                            e.dataTransfer.effectAllowed = 'move';
                            e.stopPropagation();
                        });
                        
                        cellElement.addEventListener('dragend', (e) => {
                            markColumnDragging(cellIndex, false);
                            draggedColumnIndex = -1;
                            e.stopPropagation();
                        });
                        
                        cellElement.addEventListener('dragover', (e) => {
                            e.preventDefault();
                            e.dataTransfer.dropEffect = 'move';
                            e.stopPropagation();
                        });
                        
                        cellElement.addEventListener('drop', (e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            if (draggedColumnIndex !== -1 && draggedColumnIndex !== cellIndex) {
                                moveColumn(draggedColumnIndex, cellIndex);
                            }
                        });
                        
                        const columnDragHandle = document.createElement('span');
                        columnDragHandle.className = 'column-drag-handle';
                        columnDragHandle.innerHTML = '⋯';
                        columnDragHandle.title = 'Drag to reorder column';
                        cellElement.appendChild(columnDragHandle);
                    }
                    
                    if (cellIndex === 0) {
                        const rowDragHandle = document.createElement('span');
                        rowDragHandle.className = 'drag-handle';
                        rowDragHandle.innerHTML = '⋮';
                        rowDragHandle.title = 'Click to select row, drag to reorder';
                        
                        rowDragHandle.addEventListener('click', (e) => {
                            e.stopPropagation();
                            selectRow(rowIndex);
                        });
                        
                        cellElement.appendChild(rowDragHandle);
                    }
                    
                    const shouldUseTextarea = wordWrapEnabled && (cell.length > 50 || cell.includes('\\n'));
                    const inputElement = shouldUseTextarea ? document.createElement('textarea') : document.createElement('input');
                    
                    if (!shouldUseTextarea) {
                        inputElement.type = 'text';
                    } else {
                        inputElement.rows = Math.max(1, Math.ceil(cell.length / 50));
                    }
                    
                    inputElement.value = cell;
                    inputElement.style.paddingLeft = cellIndex === 0 ? '20px' : '0';
                    
                    inputElement.addEventListener('input', (e) => {
                        tableData[rowIndex][cellIndex] = e.target.value;
                        if (shouldUseTextarea) {
                            autoResizeTextarea(e.target);
                        }
                    });
                    
                    inputElement.addEventListener('focus', () => {
                        clearSelection();
                    });
                    
                    if (shouldUseTextarea) {
                        setTimeout(() => autoResizeTextarea(inputElement), 0);
                    }
                    
                    cellElement.appendChild(inputElement);
                    tr.appendChild(cellElement);
                });
                
                table.appendChild(tr);
            });
        }
        
        function markColumnDragging(columnIndex, isDragging) {
            const table = document.getElementById('editableTable');
            const rows = table.querySelectorAll('tr');
            rows.forEach(row => {
                const cell = row.children[columnIndex];
                if (cell) {
                    if (isDragging) {
                        cell.classList.add('dragging');
                    } else {
                        cell.classList.remove('dragging');
                    }
                }
            });
        }
        
        function moveRow(fromIndex, toIndex) {
            const movedRow = tableData.splice(fromIndex, 1)[0];
            tableData.splice(toIndex, 0, movedRow);
            clearSelection();
            renderTable();
        }
        
        function moveColumn(fromIndex, toIndex) {
            tableData.forEach(row => {
                const movedCell = row.splice(fromIndex, 1)[0];
                row.splice(toIndex, 0, movedCell);
            });
            clearSelection();
            renderTable();
        }
        
        function addRow() {
            const colCount = tableData.length > 0 ? tableData[0].length : 2;
            const newRow = new Array(colCount).fill('');
            tableData.push(newRow);
            renderTable();
        }
        
        function addColumn() {
            if (tableData.length === 0) {
                tableData = [['Header 1', 'Header 2']];
            } else {
                tableData.forEach(row => row.push(''));
            }
            renderTable();
        }
        
        function removeLastRow() {
            if (tableData.length > 1) {
                tableData.pop();
                clearSelection();
                renderTable();
            }
        }
        
        function removeLastColumn() {
            if (tableData.length > 0 && tableData[0].length > 1) {
                tableData.forEach(row => row.pop());
                clearSelection();
                renderTable();
            }
        }
        
        function saveTable() {
            vscode.postMessage({
                command: 'save',
                data: tableData
            });
        }
        
        function cancel() {
            vscode.postMessage({
                command: 'cancel'
            });
        }
        
        renderTable();
    </script>
</body>
</html>`;
  }
}

export function deactivate() {}
