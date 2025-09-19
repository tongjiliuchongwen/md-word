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

// Word document generator with enhanced math support
class SimpleWordGenerator {
    constructor() {
        this.wordContent = '';
    }
    
    async generateWordContent(html) {
        // Create a Word-compatible HTML document with special formula namespace
        const wordHtml = `<!DOCTYPE html>
<html xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
      xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
    <meta charset="UTF-8">
    <title>Converted Document</title>
    <xml>
        <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
    </xml>
    <style>
        body { font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.5; margin: 1in; }
        h1 { font-size: 18pt; font-weight: bold; margin: 12pt 0; }
        h2 { font-size: 16pt; font-weight: bold; margin: 10pt 0; }
        h3 { font-size: 14pt; font-weight: bold; margin: 8pt 0; }
        p { margin: 6pt 0; text-align: justify; }
        code { font-family: 'Courier New', monospace; background-color: #f5f5f5; padding: 2px 4px; }
        pre { font-family: 'Courier New', monospace; background-color: #f5f5f5; padding: 10px; margin: 10px 0; }
        strong { font-weight: bold; }
        em { font-style: italic; }
        .math-display { text-align: center; margin: 10px 0; }
        /* Word specific formula styles */
        .MathJax_SVG { display: inline-block; }
        .MathJax_SVG_Display { display: block; margin: 1em 0; text-align: center; }
    </style>
</head>
<body>
${html}
</body>
</html>`;
        
        return wordHtml;
    }
    
    downloadAsWord(content, filename) {
        // 使用.doc格式，确保公式被正确处理
        const blob = new Blob([content], { type: 'application/msword' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename + '.doc';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
}

// Main application class
class MarkdownToWordConverter {
    constructor() {
        this.parser = new SimpleMarkdownParser();
        this.wordGenerator = new SimpleWordGenerator();
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

    // 创建Word专用的公式标记
    createWordEquation(latex, isDisplay) {
        // 为行内公式和块级公式使用不同的Word公式格式
        if (isDisplay) {
            // 块级公式 - 使用Word的居中公式格式
            return `<p style="text-align:center">
<span style="mso-element:equation;mso-element-wrap:around">
<![if !msEquation]>
<v:shape id="_x0000_i1025" type="#_x0000_t75">
</v:shape>
<![endif]>
<![if msEquation]>
<m:oMath>
<m:r><m:t>${latex}</m:t></m:r>
</m:oMath>
<![endif]>
</span>
</p>`;
        } else {
            // 行内公式 - 使用Word的行内公式格式
            return `<span style="mso-element:field-begin"></span>
<span style="mso-spacerun:yes">&nbsp;</span>EQ ${latex} 
<span style="mso-element:field-end"></span>`;
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
            // 预处理数学公式
            const { text: processedText, mathExpressions } = this.preprocessMath(markdownText);
            
            // 转换markdown为HTML
            let html = this.parser.parse(processedText);
            
            // 使用Word识别的公式格式替换数学公式
            for (let i = 0; i < mathExpressions.length; i++) {
                const math = mathExpressions[i];
                const placeholder = math.placeholder;
                
                // 使用Word专用公式标记
                if (math.type === 'display') {
                    // 块级公式
                    html = html.replace(
                        new RegExp(placeholder, 'g'),
                        `<!--[if gte msEquation 12]>
<m:oMathPara><m:oMath>${math.content}</m:oMath></m:oMathPara>
<![endif]-->`
                    );
                } else {
                    // 行内公式
                    html = html.replace(
                        new RegExp(placeholder, 'g'),
                        `<!--[if gte msEquation 12]>
<m:oMath>${math.content}</m:oMath>
<![endif]-->`
                    );
                }
            }
            
            // 生成Word文档
            const wordContent = await this.wordGenerator.generateWordContent(html);
            const fileName = `markdown-document-${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}`;
            this.wordGenerator.downloadAsWord(wordContent, fileName);
            
        } catch (error) {
            console.error('Error converting to Word:', error);
            alert('转换过程中出现错误，请稍后重试');
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
