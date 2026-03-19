import React, { useState, useRef } from 'react';
import { Upload, FileCheck, Download, RefreshCcw } from 'lucide-react';
import { modifyDocx } from './modifier';
import { saveAs } from 'file-saver';

function App() {
  const [file, setFile] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [resultBlob, setResultBlob] = useState(null);
  const fileInputRef = useRef(null);

  const handleFile = (selectedFile) => {
    if (selectedFile && selectedFile.name.toLowerCase().endsWith('.docx')) {
      setFile(selectedFile);
      setResultBlob(null);
    } else {
      alert('請上傳有效的 .docx 檔案');
    }
  };

  const handleUpload = (e) => {
    handleFile(e.target.files[0]);
  };

  const onDragOver = (e) => {
    e.preventDefault();
  };

  const onDrop = (e) => {
    e.preventDefault();
    handleFile(e.dataTransfer.files[0]);
  };

  const processFile = async () => {
    if (!file) return;
    setIsProcessing(true);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const modifiedBlob = await modifyDocx(arrayBuffer);
      setResultBlob(modifiedBlob);
    } catch (err) {
      console.error(err);
      alert('處理失敗：' + err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const downloadFile = () => {
    if (resultBlob) {
      const fileName = file.name.replace(/\.docx$/i, '_已修正版.docx');
      saveAs(resultBlob, fileName);
    }
  };

  const reset = () => {
    setFile(null);
    setResultBlob(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  return (
    <div className="container">
      <header>
        <h1>考試卷修正英雄</h1>
        <p className="subtitle">AI 驅動的 Word 考卷自動修正工具 | 針對性比對與標紅</p>
      </header>

      {!file ? (
        <div 
          className="upload-zone" 
          onClick={() => fileInputRef.current.click()}
          onDragOver={onDragOver}
          onDrop={onDrop}
        >
          <Upload className="upload-icon" />
          <div className="upload-text">點擊或拖曳 Word 檔案此處</div>
          <p style={{ color: '#64748b', marginTop: '0.5rem' }}>僅支持 .docx 格式</p>
          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleUpload} 
            style={{ display: 'none' }} 
            accept=".docx"
          />
        </div>
      ) : (
        <div className="process-area">
          <div className="file-info">
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <FileCheck style={{ marginRight: '1rem', color: '#10b981' }} />
              <div style={{ textAlign: 'left' }}>
                <div style={{ fontWeight: 700 }}>{file.name}</div>
                <div style={{ fontSize: '0.8rem', color: '#94a3b8' }}>{(file.size / 1024).toFixed(1)} KB</div>
              </div>
            </div>
          </div>

          <div style={{ marginTop: '2rem' }}>
            {!resultBlob ? (
              <button className="btn" onClick={processFile} disabled={isProcessing} style={{ width: '100%', fontSize: '1.2rem' }}>
                {isProcessing ? (
                  <><RefreshCcw className="spinning" style={{ marginRight: '0.5rem' }} /> 處理中...</>
                ) : '開始自動修正'}
              </button>
            ) : (
              <button className="btn" onClick={downloadFile} style={{ width: '100%', fontSize: '1.2rem', background: '#10b981' }}>
                <Download style={{ marginRight: '0.5rem' }} /> 下載已修正檔案
              </button>
            )}
          </div>

          {resultBlob && (
            <div className="status-grid">
              <div className="status-card">
                <div className="status-label">是非題處理</div>
                <div className="status-value">✅ 已比對並標紅</div>
              </div>
              <div className="status-card">
                <div className="status-label">選擇題處理</div>
                <div className="status-value">✅ 已校正並標紅</div>
              </div>
            </div>
          )}

          <button onClick={reset} style={{ marginTop: '2rem', background: 'transparent', border: '1px solid #475569', color: '#94a3b8', padding: '0.5rem 1rem', borderRadius: '0.5rem', cursor: 'pointer' }}>
            重新選擇檔案
          </button>
        </div>
      )}
    </div>
  );
}

export default App;
