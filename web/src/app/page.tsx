
'use client';

import { useState } from 'react';

export default function Home() {
  const [markdown, setMarkdown] = useState('');
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState<{ type: 'error' | 'success', message: string } | null>(null);

  const handleConvert = async () => {
    if (!markdown.trim()) {
      setStatus({ type: 'error', message: 'Please enter markdown content.' });
      return;
    }

    setLoading(true);
    setStatus(null);

    try {
      const response = await fetch('/api/convert', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ markdown }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Conversion failed');
      }

      // Handle download
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'document.hml';
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      setStatus({ type: 'success', message: 'Conversion successful! File downloaded.' });
    } catch (error: any) {
      console.error(error);
      setStatus({ type: 'error', message: error.message || 'An unexpected error occurred.' });
    } finally {
      setLoading(false);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const content = e.target?.result as string;
        setMarkdown(content);
      };
      reader.readAsText(file);
    }
  };

  return (
    <>
      <main className="container">
        <div className="hero">
          <h1 className="title">MD2HML Converter</h1>
          <p className="subtitle">Transform your Markdown into HWP-compatible HML documents instantly with our premium conversion engine.</p>
        </div>

        <div className="card">
          <div className="dropzone">
            <input
              type="file"
              accept=".md,.txt"
              onChange={handleFileUpload}
              style={{ display: 'none' }}
              id="fileInput"
            />
            <label htmlFor="fileInput" style={{ cursor: 'pointer', display: 'block' }}>
              <div className="drop-icon">📄</div>
              <p style={{ fontSize: '1.2rem', fontWeight: 500 }}>Click to Upload Markdown File</p>
              <p style={{ color: 'var(--text-secondary)', marginTop: '0.5rem' }}>or paste your content below</p>
            </label>
          </div>

          <textarea
            value={markdown}
            onChange={(e) => setMarkdown(e.target.value)}
            placeholder="# Your Markdown Here..."
          />

          <button
            className="btn"
            onClick={handleConvert}
            disabled={loading}
          >
            {loading ? 'Converting...' : 'Convert to HML'}
          </button>

          {status && (
            <div className={`status ${status.type}`}>
              {status.message}
            </div>
          )}
        </div>
      </main>

      <footer className="footer">
        © 2025 MD2HML Team. Powered by Antigravity.
      </footer>
    </>
  );
}
