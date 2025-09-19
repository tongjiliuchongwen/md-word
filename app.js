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
    constructor() {
        this.mathOmmlCache = new Map();
    }

    async generateWordDocument(markdown, mathExpressions) {
        try {
            if (typeof docx === 'undefined') {
                throw new Error('docx.js库未加载，请确保已正确引入docx.js');
            }

            const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;

            this.resetCache();

            const sections = this.processMarkdownWithMath(markdown, mathExpressions);
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
                            const inlineChildren = [];

                            for (const segment of section.segments) {
                                if (segment.type === 'text') {
                                    inlineChildren.push(
                                        new TextRun({
                                            text: segment.content
                                        })
                                    );
                                } else if (segment.type === 'math') {
                                    const mathComponent = await this.createMathComponent(segment.math, false);
                                    if (mathComponent) {
                                        inlineChildren.push(mathComponent);
                                    }
                                }
                            }

                            if (inlineChildren.length === 0 && section.content) {
                                inlineChildren.push(
                                    new TextRun({
                                        text: section.content
                                    })
                                );
                            }

                            children.push(
                                new Paragraph({
                                    children: inlineChildren
                                })
                            );
                        } else {
                            children.push(
                                new Paragraph({
                                    text: section.content
                                })
                            );
                        }
                        break;

                    case 'display-math':
                        if (section.math) {
                            const displayMath = await this.createMathComponent(section.math, true);
                            if (displayMath) {
                                children.push(
                                    new Paragraph({
                                        children: [displayMath],
                                        alignment: AlignmentType.CENTER
                                    })
                                );
                                break;
                            }
                        }

                        children.push(
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: section.content || '',
                                        italics: true
                                    })
                                ],
                                alignment: AlignmentType.CENTER
                            })
                        );
                        break;

                    case 'code':
                        children.push(
                            new Paragraph({
                                text: section.content,
                                font: {
                                    name: "Courier New",
                                }
                            })
                        );
                        break;
                }
            }

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

    async createMathComponent(math, isDisplay) {
        if (!math) {
            return null;
        }

        try {
            const ommlString = await this.getOmmlString(math, isDisplay);
            return docx.ImportedXmlComponent.fromXmlString(ommlString);
        } catch (error) {
            console.error('数学公式转换失败:', error);
            return new docx.TextRun({
                text: math.original || math.content || '',
                italics: true
            });
        }
    }

    async getOmmlString(math, isDisplay) {
        const cacheKey = `${math.placeholder || math.content}|${isDisplay ? 'display' : 'inline'}`;

        if (this.mathOmmlCache.has(cacheKey)) {
            return this.mathOmmlCache.get(cacheKey);
        }

        await this.ensureMathEnginesReady();

        const latex = math.content;
        if (!latex) {
            throw new Error('无法读取公式内容');
        }

        const mathml = await MathJax.tex2mmlPromise(latex, { display: isDisplay });
        const ommlString = this.convertMathMLToOmml(mathml, isDisplay);

        this.mathOmmlCache.set(cacheKey, ommlString);
        return ommlString;
    }

    async ensureMathEnginesReady() {
        if (typeof window === 'undefined') {
            throw new Error('浏览器环境缺失，无法转换数学公式');
        }

        if (typeof window.mml2omml !== 'function') {
            throw new Error('mathml2omml 转换器未加载');
        }

        if (typeof MathJax === 'undefined' || typeof MathJax.tex2mmlPromise !== 'function') {
            throw new Error('MathJax 未加载或不支持 LaTeX 转换');
        }

        if (MathJax.startup && MathJax.startup.promise) {
            await MathJax.startup.promise;
        }
    }

    convertMathMLToOmml(mathmlString, isDisplay) {
        const ommlString = window.mml2omml(mathmlString);
        if (typeof ommlString !== 'string' || !ommlString.trim()) {
            throw new Error('mathml2omml 转换失败');
        }

        const parser = new DOMParser();
        const ommlDoc = parser.parseFromString(ommlString, 'application/xml');
        if (ommlDoc.getElementsByTagName('parsererror').length > 0) {
            throw new Error('无法解析生成的 OMML');
        }

        const OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
        const serializer = new XMLSerializer();

        if (isDisplay) {
            const existingPara = ommlDoc.getElementsByTagNameNS(OMML_NS, 'oMathPara')[0];
            if (existingPara) {
                return serializer.serializeToString(existingPara);
            }

            const oMath = ommlDoc.getElementsByTagNameNS(OMML_NS, 'oMath')[0];
            if (!oMath) {
                throw new Error('缺少 oMath 节点');
            }

            const wrapperDoc = document.implementation.createDocument(OMML_NS, 'm:oMathPara', null);
            const wrapper = wrapperDoc.documentElement;
            wrapper.setAttribute('xmlns:m', OMML_NS);
            wrapper.appendChild(wrapperDoc.importNode(oMath, true));
            return serializer.serializeToString(wrapper);
        }

        const oMath = ommlDoc.getElementsByTagNameNS(OMML_NS, 'oMath')[0];
        if (!oMath) {
            throw new Error('缺少 oMath 节点');
        }

        if (!oMath.getAttribute('xmlns:m')) {
            oMath.setAttribute('xmlns:m', OMML_NS);
        }

        return serializer.serializeToString(oMath);
    }

    resetCache() {
        this.mathOmmlCache.clear();
    }

    processMarkdownWithMath(markdown, mathExpressions) {
        const sections = [];
        const lines = markdown.split('\n');

        let currentParagraph = '';

        const mathLookup = {};
        mathExpressions.forEach((math) => {
            if (math && math.placeholder) {
                mathLookup[math.placeholder] = math;
            }
        });

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];

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
                        math,
                        content: math.original || math.content,
                    });
                } else {
                    sections.push({ type: 'paragraph', content: line });
                }
                continue;
            }

            if (line.trim() === '') {
                if (currentParagraph) {
                    sections.push({ type: 'paragraph', content: currentParagraph });
                    currentParagraph = '';
                }
                continue;
            }

            if (line.includes('__MATH_INLINE_')) {
                const hasMath = /(__MATH_INLINE_\d+__)/.test(line);

                if (hasMath) {
                    const segments = [];
                    const inlineMathRegex = /__MATH_INLINE_\d+__/g;
                    let match;
                    let lastIndex = 0;

                    while ((match = inlineMathRegex.exec(line)) !== null) {
                        const beforeMath = line.substring(lastIndex, match.index);
                        if (beforeMath) {
                            segments.push({ type: 'text', content: beforeMath });
                        }

                        const placeholder = match[0];
                        const math = mathLookup[placeholder];

                        if (math) {
                            segments.push({
                                type: 'math',
                                math,
                            });
                        } else {
                            segments.push({ type: 'text', content: placeholder });
                        }

                        lastIndex = match.index + match[0].length;
                    }

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
                        segments: segments,
                        content: segments
                            .map((segment) => {
                                if (segment.type === 'text') {
                                    return segment.content;
                                }
                                return segment.math?.original || segment.math?.content || '';
                            })
                            .join(''),
                    });
                    continue;
                }
            }

            if (currentParagraph) {
                currentParagraph += ' ' + line;
            } else {
                currentParagraph = line;
            }
        }

        if (currentParagraph) {
            sections.push({ type: 'paragraph', content: currentParagraph });
        }

        return sections;
    }
    async saveDocumentToFile(doc, filename) {
        try {
            const { Packer } = docx;

            if (!Packer || typeof Packer.toBlob !== 'function') {
                throw new Error('当前 docx.js 版本不支持直接保存文档，请检查库是否正确加载');
            }

            const blob = await Packer.toBlob(doc);

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            document.body.appendChild(a);
            a.style.display = "none";
            a.href = url;
            a.download = filename;
            a.click();

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
