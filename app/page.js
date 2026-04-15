'use client';

import { useState, useRef } from 'react';

export default function HuviinHeregPage() {
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

      const res = await fetch('/api/huviin-hereg', { method: 'POST', body: formData });

      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: 'Серверийн алдаа' }));
        throw new Error(err.error || 'Серверийн алдаа');
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Huviin_hereg.docx';
      a.click();
      URL.revokeObjectURL(url);

      setStatus('success');
      setMessage('Хувийн хэрэг амжилттай үүслээ.');
    } catch (err) {
      setStatus('error');
      setMessage(err.message);
    }
  }

  return (
    <div className="page">
      <div className="card">
        <h1>Газар зүйн нэрийн хувийн хэрэг</h1>
        <p className="desc">
          Excel файл оруулна уу — хувийн хэрэгний Word баримт татагдана.
        </p>

        <a href="/api/template/huviin-hereg" className="btn btn-secondary">
          ⬇ Excel загвар татах
        </a>

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
          style={{ marginTop: 16 }}
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
            '⬇ Word баримт үүсгэх'
          )}
        </button>

        {status && status !== 'loading' && (
          <div className={`status ${status}`}>
            {status === 'success' && '✓ '}
            {status === 'error' && '✕ '}
            {message}
          </div>
        )}

        <div className="info-box">
          <h3>Шаардлагатай баганууд (*)</h3>
          <ul>
            <li>B: Дахин давтагдашгүй дугаар</li>
            <li>C: Нэрийн зургийн индекс</li>
            <li>D: Газар зүйн нэр</li>
            <li>E: Төрөл</li>
            <li>O: Өргөрөг 1 &nbsp; P: Уртраг 1</li>
            <li>S: Аймаг/сум/баг</li>
          </ul>
        </div>
      </div>
    </div>
  );
}
