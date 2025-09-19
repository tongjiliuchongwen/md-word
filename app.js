// Simple Markdown to HTML converter
class SimpleMarkdownParser {
    parse(markdown) {
        let html = markdown;
        
        // Headers
        html = html.replace(/^### (.*$)/gim, '<h3>$1</h3>');
        html = html.replace(/^## (.*$)/gim, '<h2>$1</h2>');
        html = html.replace(/^# (.*$)/gim, '<h1>$1</h1>');
        
        // Bold
        html = html.replace(/\*\*(.*)\*\*/gim, '<strong>$1</strong>');
        html = html.replace(/__(.*?)__/gim, '<strong>$1</strong>');
        
        // Italic
        html = html.replace(/\*(.*?)\*/gim, '<em>$1</em>');
        html = html.replace(/_(.*?)_/gim, '<em>$1</em>');
        
        // Code blocks
        html = html.replace(/```([\s\S]*?)```/gim, '<pre><code>$1</code></pre>');
        
        // Inline code
        html = html.replace(/`([^`]+)`/gim, '<code>$1</code>');
        
        // Line breaks
        html = html.replace(/\n\n/gim, '</p><p>');
        html = html.replace(/\n/gim, '<br>');
        
        // Wrap in paragraphs
        html = '<p>' + html + '</p>';
        
        // Clean up empty paragraphs and fix formatting
        html = html.replace(/<p><\/p>/gim, '');
        html = html.replace(/<p>(<h[1-6]>.*?<\/h[1-6]>)<\/p>/gim, '$1');
        html = html.replace(/<p>(<pre>.*?<\/pre>)<\/p>/gim, '$1');
        
        return html;
    }
}

// DocxGenerator class for creating Word documents with docx.js
class DocxGenerator {
    constructor() {}
    
    async generateWordDocument(markdown, mathExpressions) {
        try {
            if (typeof docx === 'undefined') {
                throw new Error('docx.js库未加载，请确保已正确引入docx.js');
            }
            
            const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle } = docx;
            
            // 创建一个新的文档
            const doc = new Document({
                sections: [{
                    properties: {},
                    children: []
                }]
            });
            
            // 处理带有数学公式的Markdown
            const sections = this.processMarkdownWithMath(markdown, mathExpressions);
            
            // 添加内容到文档
            for (const section of sections) {
                switch (section.type) {
                    case 'heading1':
                        doc.addParagraph(new Paragraph({
                            text: section.content,
                            heading: HeadingLevel.HEADING_1
                        }));
                        break;
                    
                    case 'heading2':
                        doc.addParagraph(new Paragraph({
                            text: section.content,
                            heading: HeadingLevel.HEADING_2
                        }));
                        break;
                    
                    case 'heading3':
                        doc.addParagraph(new Paragraph({
                            text: section.content,
                            heading: HeadingLevel.HEADING_3
                        }));
                        break;
                    
                    case 'paragraph':
                        // 处理带有行内公式的段落
                        if (section.hasMath) {
                            const children = [];
                            let currentText = '';
                            
                            for (let i = 0; i < section.segments.length; i++) {
                                const segment = section.segments[i];
                                
                                if (segment.type === 'text') {
                                    currentText += segment.content;
                                } else if (segment.type === 'math') {
                                    // 先添加前面的文本
                                    if (currentText) {
                                        children.push(new TextRun(currentText));
                                        currentText = '';
                                    }
                                    
                                    // 添加公式（使用斜体模拟，因为直接的公式支持有限）
                                    children.push(
                                        new TextRun({
                                            text: segment.content,
                                            italics: true
                                        })
                                    );
                                }
                            }
                            
                            // 添加最后的文本
                            if (currentText) {
                                children.push(new TextRun(currentText));
                            }
                            
                            doc.addParagraph(new Paragraph({ children }));
                        } else {
                            // 没有公式的普通段落
                            doc.addParagraph(new Paragraph({
                                text: section.content
                            }));
                        }
                        break;
                    
                    case 'display-math':
                        // 块级公式居中显示
                        doc.addParagraph(new Paragraph({
                            text: section.content,
                            alignment: AlignmentType.CENTER,
                            italics: true
                        }));
                        break;
                    
                    case 'code':
                        // 代码块使用等宽字体和边框
                        doc.addParagraph(new Paragraph({
                            text: section.content,
                            style: {
                                font: "Courier New",
                                size: 20,
                            },
                            border: {
                                top: { style: BorderStyle.SINGLE, size: 1, color: "auto" },
                                bottom: { style: BorderStyle.SINGLE, size: 1, color: "auto" },
                                left: { style: BorderStyle.SINGLE, size: 1, color: "auto" },
                                right: { style: BorderStyle.SINGLE, size: 1, color: "auto" }
                            }
                        }));
                        break;
                }
            }
            
            return doc;
        } catch (error) {
            console.error('生成Word文档时出错:', error);
            throw error;
        }
    }
    
    // 处理Markdown文本和数学公式
    processMarkdownWithMath(markdown, mathExpressions) {
        const sections = [];
        const lines = markdown.split('\n');
        
        let currentParagraph = '';
        
        // 构建一个查找表，用于快速查找占位符对应的公式
        const mathLookup = {};
        mathExpressions.forEach(math => {
            mathLookup[math.placeholder] = math;
        });
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            // 处理标题
            if (line.startsWith('# ')) {
                if (currentParagraph) {
                    sections.push({ type: 'paragraph', content: currentParagraph });
                    currentParagraph = '';
                }
                sections.push({ type: 'heading1', content: line.substring(2) });
                continue;
            }
            
            if (line.startsWith('## ')) {
                if (currentParagraph) {
                    sections.push({ type: 'paragraph', content: currentParagraph });
                    currentParagraph = '';
                }
                sections.push({ type: 'heading2', content: line.substring(3) });
                continue;
            }
            
            if (line.startsWith('### ')) {
                if (currentParagraph) {
                    sections.push({ type: 'paragraph', content: currentParagraph });
                    currentParagraph = '';
                }
                sections.push({ type: 'heading3', content: line.substring(4) });
                continue;
            }
            
            // 检查是否为数学公式占位符
            const displayMathMatch = line.match(/^__MATH_DISPLAY_(\d+)__$/);
            if (displayMathMatch) {
                if (currentParagraph) {
                    sections.push({ type: 'paragraph', content: currentParagraph });
                    currentParagraph = '';
                }
                
                const mathIndex = parseInt(displayMathMatch[1]);
                const math = mathExpressions[mathIndex];
                
                if (math) {
                    sections.push({ 
                        type: 'display-math', 
                        content: math.content 
                    });
                }
                continue;
            }
            
            // 检查是否是空行（段落分隔符）
            if (line.trim() === '') {
                if (currentParagraph) {
                    sections.push({ type: 'paragraph', content: currentParagraph });
                    currentParagraph = '';
                }
                continue;
            }
            
            // 处理行内公式
            if (line.includes('__MATH_INLINE_')) {
                // 检查这一行是否包含行内公式
                const hasMath = /(__MATH_INLINE_\d+__)/.test(line);
                
                if (hasMath) {
                    // 分割文本和公式
                    const segments = [];
                    let remainingLine = line;
                    
                    // 提取所有行内公式
                    const inlineMathRegex = /__MATH_INLINE_(\d+)__/g;
                    let match;
                    let lastIndex = 0;
                    
                    while ((match = inlineMathRegex.exec(line)) !== null) {
                        // 添加公式前的文本
                        const beforeMath = line.substring(lastIndex, match.index);
                        if (beforeMath) {
                            segments.push({ type: 'text', content: beforeMath });
                        }
                        
                        // 添加公式
                        const mathIndex = parseInt(match[1]);
                        const math = mathExpressions[mathIndex];
                        
                        if (math) {
                            segments.push({ 
                                type: 'math', 
                                content: math.content 
                            });
                        }
                        
                        lastIndex = match.index + match[0].length;
                    }
                    
                    // 添加最后一个公式后的文本
                    const afterLastMath = line.substring(lastIndex);
                    if (afterLastMath) {
                        segments.push({ type: 'text', content: afterLastMath });
                    }
                    
                    if (currentParagraph) {
                        sections.push({ type: 'paragraph', content: currentParagraph });
                        currentParagraph = '';
                    }
                    
                    sections.push({ 
                        type: 'paragraph', 
                        hasMath: true, 
                        segments: segments 
                    });
                    continue;
                }
            }
            
            // 普通文本行
            if (currentParagraph) {
                currentParagraph += ' ' + line;
            } else {
                currentParagraph = line;
            }
        }
        
        // 添加最后的段落
        if (currentParagraph) {
            sections.push({ type: 'paragraph', content: currentParagraph });
        }
        
        return sections;
    }
    
    // 保存并下载文档
    async saveDocumentToFile(doc, filename) {
        try {
            // 使用docx库的Packer来打包文档
            const blob = await doc.save();
            
            // 创建下载链接
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            document.body.appendChild(a);
            a.style.display = "none";
            a.href = url;
            a.download = filename;
            a.click();
            
            // 清理
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        } catch (error) {
            console.error('保存文档时出错:', error);
            throw error;
        }
    }
}

// Main application class
class MarkdownToWordConverter {
    constructor() {
        this.parser = new SimpleMarkdownParser();
        this.docxGenerator = new DocxGenerator();
        this.mathJaxLoaded = false;
        this.initializeElements();
        this.setupEventListeners();
        this.checkMathJaxAvailability();
    }

    initializeElements() {
        this.markdownInput = document.getElementById('markdown-input');
        this.previewContent = document.getElementById('preview-content');
        this.convertBtn = document.getElementById('convert-btn');
        this.fileInputElement = document.getElementById('file-input-element');
        this.fileDropZone = document.getElementById('file-drop-zone');
        this.inputMethodRadios = document.querySelectorAll('input[name="input-method"]');
        this.textInputDiv = document.getElementById('text-input');
        this.fileInputDiv = document.getElementById('file-input');
    }

    setupEventListeners() {
        // Text input changes
        this.markdownInput.addEventListener('input', () => this.updatePreview());
        
        // Convert button
        this.convertBtn.addEventListener('click', () => this.convertToWord());
        
        // Input method switching
        this.inputMethodRadios.forEach(radio => {
            radio.addEventListener('change', () => this.switchInputMethod(radio.value));
        });
        
        // File input
        this.fileInputElement.addEventListener('change', (e) => this.handleFileSelect(e));
        
        // Drag and drop
        this.fileDropZone.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.fileDropZone.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        this.fileDropZone.addEventListener('drop', (e) => this.handleDrop(e));
        
        // Initial preview update
        this.updatePreview();
    }

    checkMathJaxAvailability() {
        // Check if MathJax is available
        if (typeof MathJax !== 'undefined') {
            MathJax.startup.promise.then(() => {
                this.mathJaxLoaded = true;
                console.log('MathJax loaded successfully');
            }).catch(() => {
                this.mathJaxLoaded = false;
                console.log('MathJax failed to load');
            });
        } else {
            setTimeout(() => this.checkMathJaxAvailability(), 1000);
        }
    }

    switchInputMethod(method) {
        if (method === 'text') {
            this.textInputDiv.classList.add('active');
            this.fileInputDiv.classList.remove('active');
        } else {
            this.textInputDiv.classList.remove('active');
            this.fileInputDiv.classList.add('active');
        }
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        if (file) {
            this.readFile(file);
        }
    }

    handleDragOver(event) {
        event.preventDefault();
        this.fileDropZone.classList.add('drag-over');
    }

    handleDragLeave(event) {
        event.preventDefault();
        this.fileDropZone.classList.remove('drag-over');
    }

    handleDrop(event) {
        event.preventDefault();
        this.fileDropZone.classList.remove('drag-over');
        
        const files = event.dataTransfer.files;
        if (files.length > 0) {
            this.readFile(files[0]);
        }
    }

    readFile(file) {
        if (!file.name.match(/\.(md|txt)$/i)) {
            alert('请选择 .md 或 .txt 文件');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            this.markdownInput.value = e.target.result;
            // Switch to text input method to show the content
            document.querySelector('input[name="input-method"][value="text"]').checked = true;
            this.switchInputMethod('text');
            this.updatePreview();
        };
        reader.readAsText(file);
    }

    // Process math formulas in markdown text
    preprocessMath(text) {
        // Store math expressions to prevent markdown processing
        const mathExpressions = [];
        let mathIndex = 0;

        // Handle display math $$...$$ and \[...\]
        text = text.replace(/\$\$([\s\S]*?)\$\$/g, (match, content) => {
            const placeholder = `__MATH_DISPLAY_${mathIndex}__`;
            mathExpressions[mathIndex] = {
                type: 'display',
                content: content.trim(),
                placeholder,
                original: match
            };
            mathIndex++;
            return placeholder;
        });

        text = text.replace(/\\\[([\s\S]*?)\\\]/g, (match, content) => {
            const placeholder = `__MATH_DISPLAY_${mathIndex}__`;
            mathExpressions[mathIndex] = {
                type: 'display',
                content: content.trim(),
                placeholder,
                original: match
            };
            mathIndex++;
            return placeholder;
        });

        // Handle inline math $...$
        text = text.replace(/\$([^$\n]+?)\$/g, (match, content) => {
            const placeholder = `__MATH_INLINE_${mathIndex}__`;
            mathExpressions[mathIndex] = {
                type: 'inline',
                content: content.trim(),
                placeholder,
                original: match
            };
            mathIndex++;
            return placeholder;
        });

        return { text, mathExpressions };
    }

    // Restore math expressions after markdown processing
    restoreMath(html, mathExpressions) {
        mathExpressions.forEach((math) => {
            if (math.type === 'display') {
                html = html.replace(
                    math.placeholder,
                    `<div class="math-display">\\[${math.content}\\]</div>`
                );
            } else {
                html = html.replace(
                    math.placeholder,
                    `<span class="math-inline">\\(${math.content}\\)</span>`
                );
            }
        });
        return html;
    }

    updatePreview() {
        const markdownText = this.markdownInput.value;
        
        if (!markdownText.trim()) {
            this.previewContent.innerHTML = '<p class="placeholder">输入 Markdown 文本后，预览将显示在这里</p>';
            return;
        }

        try {
            // Preprocess math formulas
            const { text: processedText, mathExpressions } = this.preprocessMath(markdownText);
            
            // Convert markdown to HTML
            let html = this.parser.parse(processedText);
            
            // Restore math expressions
            html = this.restoreMath(html, mathExpressions);
            
            this.previewContent.innerHTML = html;
            
            // Render math if MathJax is available
            this.renderMath();
            
        } catch (error) {
            console.error('Error processing markdown:', error);
            this.previewContent.innerHTML = '<p style="color: red;">预览出现错误，请检查您的 Markdown 语法</p>';
        }
    }

    renderMath() {
        if (this.mathJaxLoaded && typeof MathJax !== 'undefined') {
            MathJax.typesetPromise([this.previewContent]).catch((err) => {
                console.log('MathJax rendering error:', err);
            });
        }
    }

    async convertToWord() {
        const markdownText = this.markdownInput.value;
        
        if (!markdownText.trim()) {
            alert('请先输入一些 Markdown 文本');
            return;
        }

        this.convertBtn.disabled = true;
        this.convertBtn.textContent = '转换中...';
        
        try {
            // 检查docx.js是否可用
            if (typeof docx === 'undefined') {
                throw new Error('docx.js库未加载，请确保已正确引入docx.js');
            }
            
            // 预处理数学公式
            const { text: processedText, mathExpressions } = this.preprocessMath(markdownText);
            
            // 使用docx.js生成Word文档
            const doc = await this.docxGenerator.generateWordDocument(processedText, mathExpressions);
            
            // 保存并下载文档
            const fileName = `markdown-document-${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.docx`;
            await this.docxGenerator.saveDocumentToFile(doc, fileName);
            
        } catch (error) {
            console.error('Error converting to Word:', error);
            alert(`转换过程中出现错误: ${error.message}`);
        } finally {
            this.convertBtn.disabled = false;
            this.convertBtn.textContent = '转换为 Word';
        }
    }
}

// Initialize the application when the page loads
document.addEventListener('DOMContentLoaded', () => {
    new MarkdownToWordConverter();
});
