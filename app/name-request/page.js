'use client';

import { useState, useRef } from 'react';

export default function NameRequestPage() {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState(null);
  const [message, setMessage] = useState('');
  const [dragover, setDragover] = useState(false);
  const inputRef = useRef(null);

  function handleFile(f) {
    if (!f) return;
    if (!f.name.match(/\.xlsx?$/i)) {
      setStatus('error');
      setMessage('Зөвхөн .xlsx файл оруулна уу.');
      return;
    }
    setFile(f);
    setStatus(null);
    setMessage('');
  }

  async function handleConvert() {
    if (!file) return;
    setStatus('loading');
    setMessage('Боловсруулж байна...');
    try {
      const formData = new FormData();
      formData.append('file', file);

      const res = await fetch('/api/name-request', { method: 'POST', body: formData });

      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: 'Серверийн алдаа' }));
        throw new Error(err.error || 'Серверийн алдаа');
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'huseltiing_maygt.docx';
      a.click();
      URL.revokeObjectURL(url);

      setStatus('success');
      setMessage('Хүсэлтийн маягт амжилттай үүслээ.');
    } catch (err) {
      setStatus('error');
      setMessage('Алдаа: ' + err.message);
    }
  }

  return (
    <div className="page">
      <div className="card">
        <h1>Газар зүйн нэрийг шинээр өгөх хүсэлт</h1>
        <p className="desc">
          Excel файл оруулна уу — хүсэлтийн маягтны Word баримт татагдана.
        </p>

        <div
          className={`drop-zone${dragover ? ' dragover' : ''}`}
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setDragover(true); }}
          onDragLeave={() => setDragover(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragover(false);
            handleFile(e.dataTransfer.files[0]);
          }}
        >
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => handleFile(e.target.files[0])}
          />
          <div className="icon">📂</div>
          <div className="label">Excel файл сонгох эсвэл чирж оруулах</div>
          <div className="hint">.xlsx өргөтгөлтэй файл</div>
          {file && <div className="file-name">✓ {file.name}</div>}
        </div>

        <button
          className="btn btn-primary"
          onClick={handleConvert}
          disabled={!file || status === 'loading'}
        >
          {status === 'loading' ? (
            <><span className="spinner" /> Боловсруулж байна...</>
          ) : (
            '⬇ Word баримт татах'
          )}
        </button>

        {status && (
          <div className={`status ${status}`}>
            {status === 'loading' && <span className="spinner" />}
            {status === 'success' && '✓ '}
            {status === 'error' && '✕ '}
            {message}
          </div>
        )}

        <div className="info-box">
          <h3>Excel файлын баганын дараалал</h3>
          <ul>
            <li>D (4): Газар зүйн нэр</li>
            <li>F (6): Дэвсгэр нэр</li>
            <li>O (15): Өргөрөг 1, P (16): Уртраг 1</li>
            <li>Q (17): Өргөрөг 2, R (18): Уртраг 2</li>
            <li>S (19): Аймаг/сум/баг</li>
            <li>T (20): Алслагдах зай / байрлал</li>
          </ul>
        </div>
      </div>
    </div>
  );
}
