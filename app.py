from flask import Flask, request, render_template_string, jsonify, send_file
import pdfplumber
import re
import os
import tempfile
import io
from openpyxl import Workbook
from datetime import datetime

app = Flask(__name__)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'

def extract_invoice_info(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            if not text:
                return {"error": "未提取到文本，请确认PDF不是扫描件"}
            
            full_text = ' '.join(text.split())
            full_text = full_text.replace('价 税 合 计', '价税合计')
            
            result = {}
            
            match = re.search(r'发票号码[：:]\s*(\d+)', full_text)
            result['发票号码'] = match.group(1) if match else ''
            
            match = re.search(r'开票日期[：:]\s*(\d{4}年\d{2}月\d{2}日)', full_text)
            result['开票日期'] = match.group(1) if match else ''
            
            match = re.search(r'购\s*名称[：:]\s*([^\s]+(?:有限公司|公司|厂|店|部))', full_text)
            if not match:
                match = re.search(r'买\s*名称[：:]\s*([^\s]+(?:有限公司|公司|厂|店|部))', full_text)
            result['购方名称'] = match.group(1) if match else ''
            
            match = re.search(r'销\s*名称[：:]\s*([^\s]+(?:有限公司|公司|厂|店|部))', full_text)
            if not match:
                match = re.search(r'售\s*名称[：:]\s*([^\s]+(?:有限公司|公司|厂|店|部))', full_text)
            result['销方名称'] = match.group(1) if match else ''
            
            match = re.search(r'合\s*计\s*[¥￥]?\s*([\d.]+)', full_text)
            if match:
                result['金额'] = match.group(1)
            else:
                match = re.search(r'合计\s*[¥￥]?\s*([\d.]+)', full_text)
                result['金额'] = match.group(1) if match else ''
            
            match = re.search(r'合\s*计\s*[¥￥]?\s*[\d.]+\s*[¥￥]?\s*([\d.]+)', full_text)
            if match:
                result['税额'] = match.group(1)
            else:
                match = re.search(r'税额[：:]\s*[¥￥]?\s*([\d.]+)', full_text)
                result['税额'] = match.group(1) if match else ''
            
            all_numbers = re.findall(r'[¥￥]\s*([\d.]+)', full_text)
            result['价税合计'] = all_numbers[-1] if all_numbers else ''
            
            return result
            
    except Exception as e:
        return {"error": f"解析失败: {str(e)}"}

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📄 发票批量解析工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: #f5f7fa;
            display: flex;
            justify-content: center;
            align-items: flex-start;
            min-height: 100vh;
            padding: 30px 20px;
        }
        .container {
            background: white;
            border-radius: 16px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.08);
            padding: 40px;
            max-width: 1000px;
            width: 100%;
        }
        h1 {
            font-size: 24px;
            color: #1a1a2e;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .subtitle {
            color: #666;
            font-size: 14px;
            margin-bottom: 24px;
        }
        .upload-area {
            border: 2px dashed #d0d7de;
            border-radius: 12px;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: border-color 0.2s;
            background: #fafbfc;
        }
        .upload-area:hover { border-color: #6c63ff; }
        .upload-area.dragover { border-color: #6c63ff; background: #f0f0ff; }
        .upload-area input { display: none; }
        .upload-icon { font-size: 48px; margin-bottom: 12px; }
        .upload-text { color: #333; font-weight: 500; }
        .upload-hint { color: #888; font-size: 13px; margin-top: 4px; }
        
        /* ===== 需求1：文件框固定大小 + 滚动 ===== */
        #file-list-wrapper {
            margin-top: 16px;
            max-height: 120px;
            overflow-y: auto;
            border: 1px solid #edf0f5;
            border-radius: 8px;
            padding: 8px 12px;
            background: #fafbfc;
        }
        #file-list {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }
        .file-tag {
            background: #eef2f7;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 13px;
            color: #333;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .file-tag .remove {
            cursor: pointer;
            color: #999;
            font-weight: bold;
            margin-left: 4px;
        }
        .file-tag .remove:hover { color: #c00; }
        .file-count {
            font-size: 13px;
            color: #888;
            margin-top: 6px;
        }
        .btn {
            background: #6c63ff;
            color: white;
            border: none;
            padding: 12px 32px;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.2s;
            margin-top: 16px;
            width: 100%;
        }
        .btn:hover { background: #5a52e0; }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; }
        .btn-success {
            background: #34a853;
        }
        .btn-success:hover { background: #2d9248; }
        .btn-success:disabled { opacity: 0.5; cursor: not-allowed; }
        .loading {
            display: none;
            text-align: center;
            padding: 20px;
            color: #6c63ff;
        }
        .loading.show { display: block; }
        .error {
            background: #fee;
            color: #c00;
            padding: 12px 16px;
            border-radius: 8px;
            margin-top: 16px;
            display: none;
        }
        .error.show { display: block; }
        .results-container {
            margin-top: 24px;
            overflow-x: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }
        th {
            background: #f0f2f5;
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
            color: #333;
            white-space: nowrap;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #edf0f5;
            max-width: 150px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        tr:hover td { background: #f8f9fc; }
        .empty { color: #ccc; }
        .status-success { color: #22c55e; font-weight: 500; }
        .status-error { color: #ef4444; font-weight: 500; }
        
        /* ===== 需求2：下载按钮在解析按钮下方 ===== */
        .action-bar {
            display: flex;
            gap: 12px;
            margin-top: 16px;
        }
        .action-bar .btn {
            width: 50%;
            margin-top: 0;
        }
        #downloadSection {
            display: none;
        }
        #downloadSection.show {
            display: block;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📄 发票批量解析工具</h1>
        <p class="subtitle">支持一次选择多张 PDF，批量提取关键信息，导出 Excel</p>
        
        <div class="upload-area" id="uploadArea">
            <div class="upload-icon">📎</div>
            <div class="upload-text">点击或拖拽上传 PDF 文件（可多选）</div>
            <div class="upload-hint">支持批量选择，一次上传多张</div>
            <input type="file" id="fileInput" accept=".pdf" multiple>
        </div>
        
        <!-- 文件列表固定区域 -->
        <div id="file-list-wrapper">
            <div id="file-list"></div>
        </div>
        <div class="file-count" id="fileCount">已选择 0 个文件</div>
        
        <!-- ===== 按钮组：解析 + 下载（并排） ===== -->
        <div class="action-bar">
            <button class="btn" id="parseBtn" disabled>🚀 开始批量解析</button>
            <button class="btn btn-success" id="downloadBtn" disabled>📥 下载 Excel</button>
        </div>
        
        <div class="loading" id="loading">⏳ 正在批量解析，请稍候...</div>
        <div class="error" id="error"></div>
        
        <div class="results-container" id="resultsContainer" style="display:none;">
            <table>
                <thead>
                    <tr>
                        <th>文件名</th>
                        <th>发票号码</th>
                        <th>开票日期</th>
                        <th>购方名称</th>
                        <th>销方名称</th>
                        <th>金额</th>
                        <th>税额</th>
                        <th>价税合计</th>
                        <th>状态</th>
                    </tr>
                </thead>
                <tbody id="resultsBody"></tbody>
            </table>
        </div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('file-list');
        const fileCount = document.getElementById('fileCount');
        const parseBtn = document.getElementById('parseBtn');
        const downloadBtn = document.getElementById('downloadBtn');
        const loadingDiv = document.getElementById('loading');
        const errorDiv = document.getElementById('error');
        const resultsContainer = document.getElementById('resultsContainer');
        const resultsBody = document.getElementById('resultsBody');
        
        let selectedFiles = [];
        let latestResults = [];
        
        function updateFileList() {
            fileList.innerHTML = '';
            selectedFiles.forEach((file, index) => {
                const tag = document.createElement('span');
                tag.className = 'file-tag';
                tag.innerHTML = `${file.name} <span class="remove" data-index="${index}" onclick="event.stopPropagation();">×</span>`;
                fileList.appendChild(tag);
            });
            
            document.querySelectorAll('.file-tag .remove').forEach(el => {
                el.addEventListener('click', function(e) {
                    e.stopPropagation();
                    const idx = parseInt(this.dataset.index);
                    selectedFiles.splice(idx, 1);
                    updateFileList();
                    parseBtn.disabled = selectedFiles.length === 0;
                    downloadBtn.disabled = true;
                    fileInput.value = '';
                    if (selectedFiles.length === 0) {
                        resultsContainer.style.display = 'none';
                    }
                });
            });
            
            fileCount.textContent = `已选择 ${selectedFiles.length} 个文件`;
            parseBtn.disabled = selectedFiles.length === 0;
        }
        
        function handleFiles(files) {
            let hasError = false;
            for (const file of files) {
                if (!file.name.toLowerCase().endsWith('.pdf')) {
                    errorDiv.textContent = '⚠️ 只支持 PDF 文件，请检查文件类型';
                    errorDiv.classList.add('show');
                    hasError = true;
                    continue;
                }
                if (!selectedFiles.some(f => f.name === file.name && f.size === file.size)) {
                    selectedFiles.push(file);
                }
            }
            if (selectedFiles.length > 0) {
                errorDiv.classList.remove('show');
            }
            updateFileList();
            fileInput.value = '';
        }
        
        uploadArea.addEventListener('click', (e) => {
            if (e.target.closest('.file-tag')) return;
            fileInput.click();
        });
        
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFiles(e.target.files);
            }
        });
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                handleFiles(e.dataTransfer.files);
            }
        });
        
        // 解析按钮
        parseBtn.addEventListener('click', async () => {
            if (selectedFiles.length === 0) return;
            
            parseBtn.disabled = true;
            downloadBtn.disabled = true;
            loadingDiv.classList.add('show');
            errorDiv.classList.remove('show');
            resultsContainer.style.display = 'none';
            resultsBody.innerHTML = '';
            latestResults = [];
            
            const results = [];
            
            for (let i = 0; i < selectedFiles.length; i++) {
                const file = selectedFiles[i];
                const formData = new FormData();
                formData.append('file', file);
                
                try {
                    const response = await fetch('/parse', {
                        method: 'POST',
                        body: formData
                    });
                    const data = await response.json();
                    
                    if (data.error) {
                        results.push({ filename: file.name, ...data, status: '失败' });
                    } else {
                        results.push({ filename: file.name, ...data, status: '成功' });
                    }
                } catch (err) {
                    results.push({ filename: file.name, error: '请求失败', status: '失败' });
                }
            }
            
            latestResults = results;
            
            loadingDiv.classList.remove('show');
            parseBtn.disabled = false;
            downloadBtn.disabled = false;
            
            let html = '';
            for (const r of results) {
                const isSuccess = r.status === '成功';
                html += `
                    <tr>
                        <td title="${r.filename}">${r.filename}</td>
                        <td>${r['发票号码'] || '<span class="empty">-</span>'}</td>
                        <td>${r['开票日期'] || '<span class="empty">-</span>'}</td>
                        <td>${r['购方名称'] || '<span class="empty">-</span>'}</td>
                        <td>${r['销方名称'] || '<span class="empty">-</span>'}</td>
                        <td>${r['金额'] || '<span class="empty">-</span>'}</td>
                        <td>${r['税额'] || '<span class="empty">-</span>'}</td>
                        <td>${r['价税合计'] || '<span class="empty">-</span>'}</td>
                        <td class="${isSuccess ? 'status-success' : 'status-error'}">${r.status}</td>
                    </tr>
                `;
            }
            resultsBody.innerHTML = html;
            resultsContainer.style.display = 'block';
            resultsContainer.scrollIntoView({ behavior: 'smooth' });
        });
        
        // 下载按钮
        downloadBtn.addEventListener('click', () => {
            if (latestResults.length === 0) return;
            
            downloadBtn.disabled = true;
            downloadBtn.textContent = '⏳ 生成中...';
            
            fetch('/download', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ results: latestResults })
            })
            .then(response => response.blob())
            .then(blob => {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                const now = new Date();
                const dateStr = now.getFullYear() + 
                    String(now.getMonth()+1).padStart(2,'0') + 
                    String(now.getDate()).padStart(2,'0');
                a.download = `发票解析结果_${dateStr}.xlsx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
                downloadBtn.disabled = false;
                downloadBtn.textContent = '📥 下载 Excel';
            })
            .catch(err => {
                errorDiv.textContent = '⚠️ 下载失败: ' + err.message;
                errorDiv.classList.add('show');
                downloadBtn.disabled = false;
                downloadBtn.textContent = '📥 下载 Excel';
            });
        });
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/parse', methods=['POST'])
def parse():
    if 'file' not in request.files:
        return jsonify({"error": "请选择文件"})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "请选择文件"})
    
    if not allowed_file(file.filename):
        return jsonify({"error": "只支持 PDF 文件"})
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        file.save(tmp_file.name)
        temp_path = tmp_file.name
    
    try:
        result = extract_invoice_info(temp_path)
        if "error" in result:
            return jsonify(result)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": f"解析失败: {str(e)}"})
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

@app.route('/download', methods=['POST'])
def download():
    data = request.get_json()
    results = data.get('results', [])
    
    if not results:
        return jsonify({"error": "没有数据可下载"}), 400
    
    wb = Workbook()
    ws = wb.active
    ws.title = "发票数据"
    
    headers = ['文件名', '发票号码', '开票日期', '购方名称', '销方名称', '金额', '税额', '价税合计', '状态']
    ws.append(headers)
    
    for r in results:
        row = [
            r.get('filename', ''),
            r.get('发票号码', ''),
            r.get('开票日期', ''),
            r.get('购方名称', ''),
            r.get('销方名称', ''),
            r.get('金额', ''),
            r.get('税额', ''),
            r.get('价税合计', ''),
            r.get('status', '')
        ]
        ws.append(row)
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column].width = adjusted_width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'发票解析结果_{datetime.now().strftime("%Y%m%d")}.xlsx'
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)