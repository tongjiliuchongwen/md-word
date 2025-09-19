const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MATH_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

function escapeHtml(value) {
    if (value === undefined || value === null) {
        return '';
    }

    const stringValue = String(value);

    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    };

    return stringValue.replace(/[&<>"']/g, (char) => map[char]);
}

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

    stripOmmlNamespaces(omml) {
        if (!omml) {
            return '';
        }

        return omml
            .replace(/\sxmlns:m="[^"]*"/g, '')
            .replace(/\sxmlns:w="[^"]*"/g, '');
    }

    createInlineMathRun(omml, ImportedXmlComponent) {
        if (!omml || !ImportedXmlComponent) {
            return null;
        }

        const sanitizedOmml = this.stripOmmlNamespaces(omml);
        if (!sanitizedOmml.trim()) {
            return null;
        }

        const runXml = `
            <w:r xmlns:w="${WORD_NAMESPACE}" xmlns:m="${MATH_NAMESPACE}">
                <w:rPr>
                    <w:rFonts ascii="Cambria Math" hAnsi="Cambria Math"/>
                </w:rPr>
                ${sanitizedOmml}
            </w:r>
        `.trim();

        try {
            return ImportedXmlComponent.fromXmlString(runXml);
        } catch (error) {
            console.warn('无法将 MathML 内联转换为 Word 运行，使用文本回退。', error);
            return null;
        }
    }

    createDisplayMathComponent(omml, ImportedXmlComponent) {
        if (!omml || !ImportedXmlComponent) {
            return null;
        }

        let sanitizedOmml = this.stripOmmlNamespaces(omml).trim();
        if (!sanitizedOmml) {
            return null;
        }

        if (!/^<m:oMath/.test(sanitizedOmml)) {
            sanitizedOmml = `<m:oMath>${sanitizedOmml}</m:oMath>`;
        }

        const mathParaXml = `<m:oMathPara xmlns:m="${MATH_NAMESPACE}" xmlns:w="${WORD_NAMESPACE}">${sanitizedOmml}</m:oMathPara>`;

        try {
            return ImportedXmlComponent.fromXmlString(mathParaXml.trim());
        } catch (error) {
            console.warn('无法创建块级数学组件，使用文本回退。', error);
            return null;
        }
    }

    createDisplayMathParagraph(omml, rawTex, mathml, Paragraph, ImportedXmlComponent) {
        const mathComponent = this.createDisplayMathComponent(omml, ImportedXmlComponent);

        if (mathComponent) {
            return new Paragraph({
                children: [mathComponent]
            });
        }

        return new Paragraph({
            text: mathml || rawTex || ''
        });
    }
    
    async generateWordDocument(markdown, mathExpressions) {
        try {
            if (typeof docx === 'undefined') {
                throw new Error('docx.js库未加载，请确保已正确引入docx.js');
            }
            
            const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, ImportedXmlComponent } = docx;
            
            // 处理带有数学公式的Markdown
            const sections = this.processMarkdownWithMath(markdown, mathExpressions);
            
            // 将sections转换为docx库需要的格式
            const children = [];
            
            for (const section of sections) {
                switch (section.type) {
                    case 'heading1':
                        children.push(
                            new Paragraph({
                                text: section.content,
                                heading: HeadingLevel.HEADING_1
                            })
                        );
                        break;
                    
                    case 'heading2':
                        children.push(
                            new Paragraph({
                                text: section.content,
                                heading: HeadingLevel.HEADING_2
                            })
                        );
                        break;
                    
                    case 'heading3':
                        children.push(
                            new Paragraph({
                                text: section.content,
                                heading: HeadingLevel.HEADING_3
                            })
                        );
                        break;
                    
                    case 'paragraph':
                        if (section.hasMath) {
                            const paragraphChildren = [];

                            for (const segment of section.segments) {
                                if (segment.type === 'text') {
                                    if (segment.content) {
                                        paragraphChildren.push(
                                            new TextRun({
                                                text: segment.content
                                            })
                                        );
                                    }
                                } else if (segment.type === 'math') {
                                    const inlineMath = this.createInlineMathRun(segment.omml, ImportedXmlComponent);

                                    if (inlineMath) {
                                        paragraphChildren.push(inlineMath);
                                    } else {
                                        const fallbackContent = segment.mathml || segment.rawTex;

                                        if (fallbackContent) {
                                            paragraphChildren.push(
                                                new TextRun({
                                                    text: fallbackContent
                                                })
                                            );
                                        }
                                    }
                                }
                            }

                            if (paragraphChildren.length > 0) {
                                children.push(
                                    new Paragraph({
                                        children: paragraphChildren
                                    })
                                );
                            }
                        } else {
                            children.push(
                                new Paragraph({
                                    text: section.content
                                })
                            );
                        }
                        break;

                    case 'display-math':
                        children.push(
                            this.createDisplayMathParagraph(section.omml, section.rawTex, section.mathml, Paragraph, ImportedXmlComponent)
                        );
                        break;
                    
                    case 'code':
                        // 代码块使用等宽字体
                        children.push(
                            new Paragraph({
                                text: section.content,
                                font: {
                                    name: "Courier New"
                                }
                            })
                        );
                        break;
                }
            }
            
            // 创建文档
            const doc = new Document({
                sections: [{
                    children: children
                }]
            });
            
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
            if (math && math.placeholder) {
                mathLookup[math.placeholder] = math;
            }
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
            const displayMathMatch = line.match(/^(__MATH_DISPLAY_\d+__)$/);
            if (displayMathMatch) {
                if (currentParagraph) {
                    sections.push({ type: 'paragraph', content: currentParagraph });
                    currentParagraph = '';
                }

                const placeholder = displayMathMatch[1];
                const math = mathLookup[placeholder];

                if (math) {
                    sections.push({
                        type: 'display-math',
                        mathType: math.type,
                        mathml: math.mathml || '',
                        omml: math.omml || '',
                        rawTex: math.content
                    });
                } else {
                    sections.push({ type: 'paragraph', content: line });
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
                    // 提取所有行内公式
                    const inlineMathRegex = /__MATH_INLINE_\d+__/g;
                    let match;
                    let lastIndex = 0;

                    while ((match = inlineMathRegex.exec(line)) !== null) {
                        // 添加公式前的文本
                        const beforeMath = line.substring(lastIndex, match.index);
                        if (beforeMath) {
                            segments.push({ type: 'text', content: beforeMath });
                        }
                        
                        // 添加公式
                        const placeholder = match[0];
                        const math = mathLookup[placeholder];

                        if (math) {
                            segments.push({
                                type: 'math',
                                mathType: math.type,
                                mathml: math.mathml || '',
                                omml: math.omml || '',
                                rawTex: math.content
                            });
                        } else {
                            segments.push({ type: 'text', content: placeholder });
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
            const { Packer } = docx;

            if (!Packer || typeof Packer.toBlob !== 'function') {
                throw new Error('当前 docx.js 版本不支持直接保存文档，请检查库是否正确加载');
            }

            const blob = await Packer.toBlob(doc);
            
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
                this.updatePreview();
            }).catch(() => {
                this.mathJaxLoaded = false;
                console.log('MathJax failed to load');
            });
        } else {
            setTimeout(() => this.checkMathJaxAvailability(), 1000);
        }
    }

    async ensureMathJaxReady() {
        if (this.mathJaxLoaded && typeof MathJax !== 'undefined') {
            return;
        }

        if (typeof MathJax !== 'undefined' && MathJax.startup && MathJax.startup.promise) {
            try {
                await MathJax.startup.promise;
                this.mathJaxLoaded = true;
            } catch (error) {
                console.warn('MathJax 初始化失败，使用备用 MathML 生成逻辑。', error);
            }
        }
    }

    convertTeXToMathML(tex, isDisplay) {
        if (!tex) {
            return '';
        }

        if (this.mathJaxLoaded && typeof MathJax !== 'undefined' && typeof MathJax.tex2mml === 'function') {
            try {
                return MathJax.tex2mml(tex, { display: isDisplay });
            } catch (error) {
                console.warn('MathJax 转换 MathML 时出错，改用备用方案。', error);
            }
        }

        return this.createFallbackMathML(tex, isDisplay);
    }

    createFallbackMathML(tex, isDisplay) {
        const safeContent = escapeHtml(tex);
        const displayAttr = isDisplay ? ' display="block"' : '';
        return '<math xmlns="http://www.w3.org/1998/Math/MathML"' + displayAttr + '><mrow><mtext>' + safeContent + '</mtext></mrow></math>';
    }

    convertMathMLToOMML(mathML) {
        if (!mathML || typeof window === 'undefined') {
            return '';
        }

        if (typeof window.mml2omml !== 'function') {
            return '';
        }

        try {
            return window.mml2omml(mathML);
        } catch (error) {
            console.warn('MathML 转换为 OMML 时出错，将在 Word 中使用备用文本。', error);
            return '';
        }
    }

    populateMathExpressions(mathExpressions) {
        if (!Array.isArray(mathExpressions)) {
            return;
        }

        mathExpressions.forEach((math) => {
            if (!math) {
                return;
            }

            math.mathml = this.convertTeXToMathML(math.content, math.type === 'display');
            math.omml = this.convertMathMLToOMML(math.mathml);
        });
    }

    async prepareMathExpressions(mathExpressions) {
        await this.ensureMathJaxReady();
        this.populateMathExpressions(mathExpressions);
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
                original: match,
                mathml: null,
                omml: null
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
                original: match,
                mathml: null,
                omml: null
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
                original: match,
                mathml: null,
                omml: null
            };
            mathIndex++;
            return placeholder;
        });

        return { text, mathExpressions };
    }

    // Restore math expressions after markdown processing
    restoreMath(html, mathExpressions) {
        if (!Array.isArray(mathExpressions) || mathExpressions.length === 0) {
            return html;
        }

        mathExpressions.forEach((math) => {
            if (!math || !math.placeholder) {
                return;
            }

            const mathML = math.mathml || this.convertTeXToMathML(math.content, math.type === 'display');
            const safeMathMarkup = mathML || this.createFallbackMathML(math.content, math.type === 'display');
            const mathMLCode = escapeHtml(safeMathMarkup);

            if (math.type === 'display') {
                const replacement = `<div class="math-preview math-display" data-placeholder="${math.placeholder}">
                        <div class="mathml-label">块级公式</div>
                        <div class="mathml-formula">${safeMathMarkup}</div>
                        <details class="mathml-source">
                            <summary>查看 MathML 代码</summary>
                            <pre class="mathml-code">${mathMLCode}</pre>
                        </details>
                    </div>`;
                html = html.replace(math.placeholder, replacement);
            } else {
                const replacement = `<span class="math-preview math-inline" data-placeholder="${math.placeholder}">
                        <span class="mathml-label">行内公式</span>
                        <span class="mathml-formula">${safeMathMarkup}</span>
                        <details class="mathml-source">
                            <summary>MathML</summary>
                            <code class="mathml-code mathml-inline-code">${mathMLCode}</code>
                        </details>
                    </span>`;
                html = html.replace(math.placeholder, replacement);
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

            // 准备 MathML 表达式
            this.populateMathExpressions(mathExpressions);

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
        if (this.mathJaxLoaded && typeof MathJax !== 'undefined' && this.previewContent) {
            const mathContainers = Array.from(this.previewContent.querySelectorAll('.mathml-formula'));
            const mathElements = mathContainers.flatMap((container) => Array.from(container.querySelectorAll('math')));

            if (!mathElements.length) {
                return;
            }

            MathJax.typesetPromise(mathElements).catch((err) => {
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

            // 确保 MathML 已准备好
            await this.prepareMathExpressions(mathExpressions);

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
