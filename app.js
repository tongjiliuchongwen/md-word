const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MATH_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

function escapeHtml(value) {
    if (value === undefined || value === null) {
        return '';
    }

    return String(value).replace(/[&<>"']/g, (char) => {
        switch (char) {
            case '&':
                return '&amp;';
            case '<':
                return '&lt;';
            case '>':
                return '&gt;';
            case '"':
                return '&quot;';
            case "'":
                return '&#39;';
            default:
                return char;
        }
    });
}

class MarkdownParser {
    constructor() {
        this.markedInitialized = false;
    }

    parse(markdown) {
        if (!markdown) {
            return '';
        }

        if (window.marked && typeof window.marked.parse === 'function') {
            if (!this.markedInitialized && typeof window.marked.setOptions === 'function') {
                window.marked.setOptions({ breaks: true });
                this.markedInitialized = true;
            }

            return window.marked.parse(markdown);
        }

        let html = markdown;

        html = html.replace(/\r\n/g, '\n');

        html = html.replace(/^######\s+(.*)$/gim, '<h6>$1</h6>');
        html = html.replace(/^#####\s+(.*)$/gim, '<h5>$1</h5>');
        html = html.replace(/^####\s+(.*)$/gim, '<h4>$1</h4>');
        html = html.replace(/^###\s+(.*)$/gim, '<h3>$1</h3>');
        html = html.replace(/^##\s+(.*)$/gim, '<h2>$1</h2>');
        html = html.replace(/^#\s+(.*)$/gim, '<h1>$1</h1>');

        html = html.replace(/\*\*(.*?)\*\*/gim, '<strong>$1</strong>');
        html = html.replace(/__(.*?)__/gim, '<strong>$1</strong>');

        html = html.replace(/\*(.*?)\*/gim, '<em>$1</em>');
        html = html.replace(/_(.*?)_/gim, '<em>$1</em>');

        html = html.replace(/`([^`]+)`/gim, '<code>$1</code>');

        html = html.replace(/^>\s?(.*)$/gim, '<blockquote>$1</blockquote>');

        html = html.replace(/\n\n+/g, '</p><p>');
        html = html.replace(/\n/g, '<br>');

        html = `<p>${html}</p>`;

        html = html.replace(/<p><\/p>/gim, '');
        html = html.replace(/<p>(<h[1-6][^>]*>.*?<\/h[1-6]>)<\/p>/gim, '$1');
        html = html.replace(/<p>(<blockquote>.*?<\/blockquote>)<\/p>/gim, '$1');

        return html;
    }
}

class DocxGenerator {
    unwrapImportedComponent(component) {
        if (!component) {
            return null;
        }

        if (component.rootKey) {
            return component;
        }

        if (Array.isArray(component.root)) {
            for (const child of component.root) {
                if (child && typeof child === 'object') {
                    const unwrapped = this.unwrapImportedComponent(child);
                    if (unwrapped) {
                        return unwrapped;
                    }
                }
            }
        }

        return null;
    }

    stripOmmlNamespaces(omml) {
        if (!omml) {
            return '';
        }

        return omml
            .replace(/\sxmlns:m="[^"]*"/g, '')
            .replace(/\sxmlns:w="[^"]*"/g, '');
    }

    extractSingleOMath(omml) {
        if (!omml) {
            return '';
        }

        const trimmed = omml.trim();

        if (/^<m:oMathPara\b/.test(trimmed)) {
            const match = trimmed.match(/<m:oMath\b[\s\S]*?<\/m:oMath>/);
            if (match) {
                return match[0];
            }
        }

        return trimmed;
    }

    createInlineMathRun(math, ImportedXmlComponent) {
        if (!math || !ImportedXmlComponent) {
            return null;
        }

        let sanitizedOmml = this.stripOmmlNamespaces(math.omml);
        sanitizedOmml = this.extractSingleOMath(sanitizedOmml);
        if (!sanitizedOmml || !sanitizedOmml.trim()) {
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
            const imported = ImportedXmlComponent.fromXmlString(runXml);
            const component = this.unwrapImportedComponent(imported);
            return component || null;
        } catch (error) {
            console.warn('内联数学公式转换失败，使用文本回退。', error);
            return null;
        }
    }

    createDisplayMathComponent(math, ImportedXmlComponent) {
        if (!math || !ImportedXmlComponent) {
            return null;
        }

        let sanitizedOmml = this.stripOmmlNamespaces(math.omml || '').trim();
        if (!sanitizedOmml) {
            return null;
        }

        if (/^<m:oMathPara\b/.test(sanitizedOmml)) {
            sanitizedOmml = sanitizedOmml
                .replace(/^<m:oMathPara\b[^>]*>/, '')
                .replace(/<\/m:oMathPara>\s*$/, '')
                .trim();
        }

        if (!/^<m:oMath\b/.test(sanitizedOmml)) {
            sanitizedOmml = `<m:oMath>${sanitizedOmml}</m:oMath>`;
        }

        const mathParaXml = `<m:oMathPara xmlns:m="${MATH_NAMESPACE}" xmlns:w="${WORD_NAMESPACE}">${sanitizedOmml}</m:oMathPara>`;

        try {
            const imported = ImportedXmlComponent.fromXmlString(mathParaXml.trim());
            const component = this.unwrapImportedComponent(imported);
            return component || null;
        } catch (error) {
            console.warn('块级数学公式转换失败，使用文本回退。', error);
            return null;
        }
    }

    createDisplayMathParagraph(math, Paragraph, ImportedXmlComponent, AlignmentType) {
        const mathComponent = this.createDisplayMathComponent(math, ImportedXmlComponent);

        if (mathComponent) {
            return new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [mathComponent]
            });
        }

        return new Paragraph({
            alignment: AlignmentType.CENTER,
            text: math ? math.content : ''
        });
    }

    processMarkdownWithMath(markdown, mathExpressions) {
        const sections = [];
        const lines = markdown.split(/\r?\n/);
        const mathLookup = new Map();

        (mathExpressions || []).forEach((math) => {
            if (math && math.placeholder) {
                mathLookup.set(math.placeholder, math);
            }
        });

        let inCodeBlock = false;
        let codeBuffer = [];
        let paragraphBuffer = [];

        const flushParagraph = () => {
            if (paragraphBuffer.length === 0) {
                return;
            }

            const paragraphText = paragraphBuffer.join(' ').trim();
            paragraphBuffer = [];

            if (!paragraphText) {
                return;
            }

            if (/__MATH_INLINE_\d+__/.test(paragraphText)) {
                const inlineRegex = /__MATH_INLINE_\d+__/g;
                const segments = [];
                let lastIndex = 0;
                let match;

                while ((match = inlineRegex.exec(paragraphText)) !== null) {
                    if (match.index > lastIndex) {
                        segments.push({
                            type: 'text',
                            content: paragraphText.substring(lastIndex, match.index)
                        });
                    }

                    const placeholder = match[0];
                    const math = mathLookup.get(placeholder);

                    if (math) {
                        segments.push({ type: 'math', math });
                    } else {
                        segments.push({ type: 'text', content: placeholder });
                    }

                    lastIndex = inlineRegex.lastIndex;
                }

                if (lastIndex < paragraphText.length) {
                    segments.push({
                        type: 'text',
                        content: paragraphText.substring(lastIndex)
                    });
                }

                const hasMath = segments.some((segment) => segment.type === 'math');

                if (hasMath) {
                    sections.push({ type: 'paragraph', hasMath: true, segments });
                    return;
                }

                sections.push({ type: 'paragraph', content: paragraphText });
                return;
            }

            sections.push({ type: 'paragraph', content: paragraphText });
        };

        const flushCodeBlock = () => {
            if (!inCodeBlock) {
                return;
            }

            sections.push({
                type: 'code',
                content: codeBuffer.join('\n')
            });

            inCodeBlock = false;
            codeBuffer = [];
        };

        for (const rawLine of lines) {
            const line = rawLine || '';
            const trimmed = line.trim();

            if (/^```/.test(trimmed)) {
                if (inCodeBlock) {
                    flushCodeBlock();
                } else {
                    flushParagraph();
                    inCodeBlock = true;
                }
                continue;
            }

            if (inCodeBlock) {
                codeBuffer.push(line);
                continue;
            }

            if (!trimmed) {
                flushParagraph();
                continue;
            }

            const headingMatch = trimmed.match(/^(#{1,6})\s+(.*)$/);
            if (headingMatch) {
                flushParagraph();
                const level = headingMatch[1].length;
                const type = `heading${level}`;
                sections.push({ type, content: headingMatch[2].trim() });
                continue;
            }

            const displayMatch = trimmed.match(/^__MATH_DISPLAY_\d+__$/);
            if (displayMatch) {
                flushParagraph();
                const math = mathLookup.get(displayMatch[0]);
                if (math) {
                    sections.push({ type: 'display-math', math });
                } else {
                    sections.push({ type: 'paragraph', content: trimmed });
                }
                continue;
            }

            paragraphBuffer.push(line);
        }

        flushParagraph();
        flushCodeBlock();

        return sections;
    }

    async generateWordDocument(markdown, mathExpressions) {
        if (typeof docx === 'undefined') {
            throw new Error('docx.js 库未加载');
        }

        const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, ImportedXmlComponent } = docx;
        const sections = this.processMarkdownWithMath(markdown, mathExpressions);
        const children = [];

        sections.forEach((section) => {
            switch (section.type) {
                case 'heading1':
                    children.push(new Paragraph({
                        text: section.content,
                        heading: HeadingLevel.HEADING_1
                    }));
                    break;
                case 'heading2':
                    children.push(new Paragraph({
                        text: section.content,
                        heading: HeadingLevel.HEADING_2
                    }));
                    break;
                case 'heading3':
                    children.push(new Paragraph({
                        text: section.content,
                        heading: HeadingLevel.HEADING_3
                    }));
                    break;
                case 'heading4':
                    children.push(new Paragraph({
                        text: section.content,
                        heading: HeadingLevel.HEADING_4
                    }));
                    break;
                case 'heading5':
                    children.push(new Paragraph({
                        text: section.content,
                        heading: HeadingLevel.HEADING_5
                    }));
                    break;
                case 'heading6':
                    children.push(new Paragraph({
                        text: section.content,
                        heading: HeadingLevel.HEADING_6
                    }));
                    break;
                case 'display-math': {
                    const paragraph = this.createDisplayMathParagraph(
                        section.math,
                        Paragraph,
                        ImportedXmlComponent,
                        AlignmentType
                    );
                    children.push(paragraph);
                    break;
                }
                case 'code': {
                    const codeLines = String(section.content || '').split('\n');
                    const codeRuns = [];

                    codeLines.forEach((line, index) => {
                        codeRuns.push(
                            new TextRun({
                                text: line,
                                font: 'Courier New',
                                break: index === 0 ? undefined : 1
                            })
                        );
                    });

                    children.push(new Paragraph({ children: codeRuns }));
                    break;
                }
                case 'paragraph':
                default: {
                    if (section.hasMath && Array.isArray(section.segments)) {
                        const inlineChildren = [];
                        section.segments.forEach((segment) => {
                            if (segment.type === 'math') {
                                const mathComponent = this.createInlineMathRun(
                                    segment.math,
                                    ImportedXmlComponent
                                );
                                if (mathComponent) {
                                    inlineChildren.push(mathComponent);
                                } else if (segment.math) {
                                    inlineChildren.push(
                                        new TextRun({ text: segment.math.content })
                                    );
                                }
                            } else if (segment.type === 'text' && segment.content) {
                                inlineChildren.push(new TextRun({ text: segment.content }));
                            }
                        });

                        if (inlineChildren.length > 0) {
                            children.push(new Paragraph({ children: inlineChildren }));
                        }
                    } else if (section.content) {
                        children.push(new Paragraph({ text: section.content }));
                    }
                    break;
                }
            }
        });

        const doc = new Document({
            sections: [
                {
                    children
                }
            ]
        });

        return doc;
    }

    async saveDocumentToFile(doc, filename) {
        const { Packer } = docx;

        if (!Packer || typeof Packer.toBlob !== 'function') {
            throw new Error('当前 docx.js 版本不支持保存文档');
        }

        const blob = await Packer.toBlob(doc);
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');

        link.style.display = 'none';
        link.href = url;
        link.download = filename;

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
    }
}

class MarkdownToWordConverter {
    constructor() {
        this.parser = new MarkdownParser();
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
        if (this.markdownInput) {
            this.markdownInput.addEventListener('input', () => this.updatePreview());
        }

        if (this.convertBtn) {
            this.convertBtn.addEventListener('click', () => this.convertToWord());
        }

        this.inputMethodRadios.forEach((radio) => {
            radio.addEventListener('change', () => this.switchInputMethod(radio.value));
        });

        if (this.fileInputElement) {
            this.fileInputElement.addEventListener('change', (event) => this.handleFileSelect(event));
        }

        if (this.fileDropZone) {
            this.fileDropZone.addEventListener('dragover', (event) => this.handleDragOver(event));
            this.fileDropZone.addEventListener('dragleave', (event) => this.handleDragLeave(event));
            this.fileDropZone.addEventListener('drop', (event) => this.handleDrop(event));
        }

        this.updatePreview();
    }

    checkMathJaxAvailability() {
        if (typeof MathJax !== 'undefined' && MathJax.startup && MathJax.startup.promise) {
            MathJax.startup.promise
                .then(() => {
                    this.mathJaxLoaded = true;
                    this.renderMath();
                })
                .catch((error) => {
                    console.warn('MathJax 加载失败，将使用回退方案。', error);
                    this.mathJaxLoaded = false;
                });
        } else {
            setTimeout(() => this.checkMathJaxAvailability(), 500);
        }
    }

    async ensureMathJaxReady() {
        if (this.mathJaxLoaded && typeof MathJax !== 'undefined' && MathJax.startup) {
            return;
        }

        if (typeof MathJax !== 'undefined' && MathJax.startup && MathJax.startup.promise) {
            try {
                await MathJax.startup.promise;
                this.mathJaxLoaded = true;
            } catch (error) {
                console.warn('等待 MathJax 初始化失败，使用回退方案。', error);
            }
        }
    }

    convertTeXToMathML(tex, isDisplay) {
        if (!tex) {
            return '';
        }

        if (
            this.mathJaxLoaded &&
            typeof MathJax !== 'undefined' &&
            typeof MathJax.tex2mml === 'function'
        ) {
            try {
                return MathJax.tex2mml(tex, { display: isDisplay });
            } catch (error) {
                console.warn('MathJax 转换 MathML 失败，使用回退方案。', error);
            }
        }

        return this.createFallbackMathML(tex, isDisplay);
    }

    createFallbackMathML(tex, isDisplay) {
        const safeContent = escapeHtml(tex);
        const displayAttr = isDisplay ? ' display="block"' : '';
        return `<math xmlns="http://www.w3.org/1998/Math/MathML"${displayAttr}><mrow><mtext>${safeContent}</mtext></mrow></math>`;
    }

    convertMathMLToOMML(mathml) {
        if (!mathml || typeof window.mml2omml !== 'function') {
            return '';
        }

        try {
            return window.mml2omml(mathml);
        } catch (error) {
            console.warn('MathML 转换为 OMML 失败，将使用文本回退。', error);
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
            if (!math.original) {
                if (math.type === 'display') {
                    math.original = `$$${math.content}$$`;
                } else {
                    math.original = `$${math.content}$`;
                }
            }
        });
    }

    async prepareMathExpressions(mathExpressions) {
        if (!Array.isArray(mathExpressions) || mathExpressions.length === 0) {
            return;
        }

        await this.ensureMathJaxReady();
        this.populateMathExpressions(mathExpressions);

        mathExpressions.forEach((math) => {
            if (!math) {
                return;
            }

            if (!math.mathml) {
                math.mathml = this.createFallbackMathML(math.content, math.type === 'display');
            }

            math.omml = this.convertMathMLToOMML(math.mathml);
        });
    }

    switchInputMethod(method) {
        if (method === 'text') {
            this.textInputDiv?.classList.add('active');
            this.fileInputDiv?.classList.remove('active');
        } else {
            this.textInputDiv?.classList.remove('active');
            this.fileInputDiv?.classList.add('active');
        }
    }

    handleFileSelect(event) {
        const file = event.target.files && event.target.files[0];
        if (file) {
            this.readFile(file);
        }
    }

    handleDragOver(event) {
        event.preventDefault();
        this.fileDropZone?.classList.add('drag-over');
    }

    handleDragLeave(event) {
        event.preventDefault();
        this.fileDropZone?.classList.remove('drag-over');
    }

    handleDrop(event) {
        event.preventDefault();
        this.fileDropZone?.classList.remove('drag-over');

        const files = event.dataTransfer?.files;
        if (files && files.length > 0) {
            this.readFile(files[0]);
        }
    }

    readFile(file) {
        if (!file.name.match(/\.(md|txt)$/i)) {
            alert('请选择 .md 或 .txt 文件');
            return;
        }

        const reader = new FileReader();
        reader.onload = (event) => {
            const result = event.target?.result || '';
            if (typeof result === 'string') {
                this.markdownInput.value = result;
                const textRadio = document.querySelector('input[name="input-method"][value="text"]');
                if (textRadio) {
                    textRadio.checked = true;
                }
                this.switchInputMethod('text');
                this.updatePreview();
            }
        };

        reader.readAsText(file);
    }

    preprocessMath(text) {
        const mathExpressions = [];
        let workingText = text;
        let counter = 0;

        const storeMath = (type, content, original) => {
            const placeholder = type === 'display'
                ? `__MATH_DISPLAY_${counter}__`
                : `__MATH_INLINE_${counter}__`;

            mathExpressions.push({
                type,
                content: content.trim(),
                placeholder,
                original
            });

            counter += 1;
            return placeholder;
        };

        workingText = workingText.replace(/\$\$([\s\S]+?)\$\$/g, (match, content) => {
            return storeMath('display', content, match);
        });

        workingText = workingText.replace(/\\\[([\s\S]+?)\\\]/g, (match, content) => {
            return storeMath('display', content, match);
        });

        workingText = workingText.replace(/\\\(([\s\S]+?)\\\)/g, (match, content) => {
            return storeMath('inline', content, match);
        });

        workingText = workingText.replace(/(^|[^$\\])\$([^$\n]+?)\$(?!\$)/g, (fullMatch, prefix, content) => {
            return `${prefix}${storeMath('inline', content, `$${content}$`)}`;
        });

        return {
            text: workingText,
            mathExpressions
        };
    }

    restoreMath(html, mathExpressions) {
        if (!Array.isArray(mathExpressions) || mathExpressions.length === 0) {
            return html;
        }

        let result = html;

        mathExpressions.forEach((math) => {
            if (!math || !math.placeholder) {
                return;
            }

            const tex = math.original || (math.type === 'display' ? `$$${math.content}$$` : `$${math.content}$`);
            const isDisplay = math.type === 'display';
            const classes = isDisplay
                ? 'math-expression math-display'
                : 'math-expression math-inline';
            const extraAttr = isDisplay ? ' style="display:block; text-align:center;"' : '';
            const replacement = `<span class="${classes}" data-tex="${escapeHtml(math.content)}"${extraAttr}>${tex}</span>`;
            result = result.split(math.placeholder).join(replacement);
        });

        return result;
    }

    updatePreview() {
        if (!this.markdownInput || !this.previewContent) {
            return;
        }

        const markdownText = this.markdownInput.value || '';

        if (!markdownText.trim()) {
            this.previewContent.innerHTML = '<p class="placeholder">输入 Markdown 文本后，预览将显示在这里</p>';
            return;
        }

        try {
            const { text: processedText, mathExpressions } = this.preprocessMath(markdownText);
            const html = this.parser.parse(processedText);
            const htmlWithMath = this.restoreMath(html, mathExpressions);

            this.previewContent.innerHTML = htmlWithMath;
            this.renderMath();
        } catch (error) {
            console.error('预览渲染失败:', error);
            this.previewContent.innerHTML = '<p style="color: red;">预览出现错误，请检查 Markdown 内容</p>';
        }
    }

    renderMath() {
        if (!this.previewContent) {
            return;
        }

        if (
            this.mathJaxLoaded &&
            typeof MathJax !== 'undefined' &&
            typeof MathJax.typesetPromise === 'function'
        ) {
            MathJax.typesetClear?.([this.previewContent]);
            MathJax.typesetPromise([this.previewContent]).catch((error) => {
                console.warn('MathJax 渲染出错，使用原始文本显示。', error);
            });
        }
    }

    async convertToWord() {
        if (!this.markdownInput) {
            return;
        }

        const markdownText = this.markdownInput.value || '';

        if (!markdownText.trim()) {
            alert('请先输入一些 Markdown 文本');
            return;
        }

        if (this.convertBtn) {
            this.convertBtn.disabled = true;
            this.convertBtn.textContent = '转换中...';
        }

        try {
            const { text: processedText, mathExpressions } = this.preprocessMath(markdownText);
            await this.prepareMathExpressions(mathExpressions);
            const doc = await this.docxGenerator.generateWordDocument(processedText, mathExpressions);
            const fileName = `markdown-document-${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.docx`;
            await this.docxGenerator.saveDocumentToFile(doc, fileName);
        } catch (error) {
            console.error('转换为 Word 时发生错误:', error);
            alert(`转换过程中出现错误: ${error.message}`);
        } finally {
            if (this.convertBtn) {
                this.convertBtn.disabled = false;
                this.convertBtn.textContent = '转换为 Word';
            }
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new MarkdownToWordConverter();
});
