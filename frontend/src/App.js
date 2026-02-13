import React, { useState } from 'react';
import './App.css';

function App() {
  const [files, setFiles] = useState([]);
  const [sessionId, setSessionId] = useState(null);
  const [extractedData, setExtractedData] = useState(null);
  const [reportDate, setReportDate] = useState('');
  const [adviserComments, setAdviserComments] = useState('');
  const [uploading, setUploading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [message, setMessage] = useState('');

  const handleFileChange = async (e) => {
    const selectedFiles = Array.from(e.target.files);
    setUploading(true);
    setMessage('Uploading files and extracting data...');

    try {
      const formData = new FormData();
      selectedFiles.forEach(file => {
        formData.append('files', file);
      });

      console.log('Uploading files to backend...');
      
      // CORRECTED: Define apiUrl FIRST
      const apiUrl = process.env.NODE_ENV === 'production' 
        ? '/api' 
        : 'http://localhost:3001/api';

      const response = await fetch(`${apiUrl}/upload`, {
        method: 'POST',
        body: formData
      });

      const data = await response.json();
      console.log('Upload response:', data);

      if (data.success) {
        setFiles(data.files);
        setSessionId(data.sessionId);
        setExtractedData(data.extractedData);
        setMessage('✅ Files uploaded and data extracted successfully!');
      } else {
        setMessage('❌ Upload failed: ' + data.error);
      }
    } catch (error) {
      console.error('Upload error:', error);
      setMessage('❌ Upload failed: ' + error.message);
    } finally {
      setUploading(false);
    }
  };

  const handleGenerate = async () => {
    if (!sessionId) {
      setMessage('❌ Please upload files first');
      return;
    }

    setGenerating(true);
    setMessage('Generating comprehensive report...');

    try {
      // CORRECTED: Define apiUrl FIRST
      const apiUrl = process.env.NODE_ENV === 'production' 
        ? '/api' 
        : 'http://localhost:3001/api';

      const response = await fetch(`${apiUrl}/generate-report`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          sessionId,
          clientName: extractedData?.clientName || 'Client Name',
          reportDate: reportDate || new Date().toLocaleDateString('en-GB'),
          adviserComments
        })
      });

      const data = await response.json();
      console.log('Generate response:', data);

      if (data.success) {
        setMessage('✅ Report generated successfully!');
        
        // CORRECTED: Use apiUrl for download too
        const downloadUrl = process.env.NODE_ENV === 'production'
          ? data.downloadUrl
          : `http://localhost:3001${data.downloadUrl}`;
        
        window.location.href = downloadUrl;
      } else {
        setMessage('❌ Generation failed: ' + data.error);
      }
    } catch (error) {
      console.error('Generate error:', error);
      setMessage('❌ Generation failed: ' + error.message);
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>📊 Financial Report Generator</h1>
        <p>Automated Multi-Provider Report System</p>
      </header>

      <main className="App-main">
        {/* Upload Section */}
        <section className="upload-section">
          <h2>1. Upload Documents</h2>
          <p>Upload pension PDFs (AJ Bell, Morningstar) and cashflow documents</p>
          
          <input
            type="file"
            multiple
            accept=".pdf,.docx"
            onChange={handleFileChange}
            disabled={uploading}
            className="file-input"
          />

          {uploading && <div className="spinner">⏳ Uploading and extracting data...</div>}

          {files.length > 0 && (
            <div className="file-list">
              <h3>Uploaded Files:</h3>
              <ul>
                {files.map((file, idx) => (
                  <li key={idx}>{file.name} ({file.size})</li>
                ))}
              </ul>
            </div>
          )}
        </section>

        {/* Extracted Data Section */}
        {extractedData && (
          <section className="extracted-section">
            <h2>2. Extracted Data</h2>
            
            {extractedData.clientName && (
              <div className="data-item">
                <strong>Client:</strong> {extractedData.clientName}
              </div>
            )}

            {extractedData.accounts && extractedData.accounts.length > 0 && (
              <div className="data-item">
                <strong>Accounts Found:</strong> {extractedData.accounts.length}
                <ul>
                  {extractedData.accounts.map((acc, idx) => (
                    <li key={idx}>
                      {acc.type} ({acc.provider}): £{acc.value?.toLocaleString() || '0'}
                      {acc.performance && ` - ${acc.performance}% return`}
                    </li>
                  ))}
                </ul>
              </div>
            )}

            {extractedData.totalValue > 0 && (
              <div className="data-item">
                <strong>Total Portfolio:</strong> £{extractedData.totalValue.toLocaleString()}
              </div>
            )}

            {extractedData.chartsExtracted && (
              <div className="data-item success">
                ✅ Cashflow charts extracted
              </div>
            )}
          </section>
        )}

        {/* Report Details Section */}
        {sessionId && (
          <section className="details-section">
            <h2>3. Report Details</h2>
            
            <div className="form-group">
              <label htmlFor="reportDate">Report Date:</label>
              <input
                id="reportDate"
                type="text"
                value={reportDate}
                onChange={(e) => setReportDate(e.target.value)}
                placeholder="e.g., 5 February 2026"
                className="text-input"
              />
            </div>

            <div className="form-group">
              <label htmlFor="adviserComments">Adviser Comments (optional):</label>
              <textarea
                id="adviserComments"
                value={adviserComments}
                onChange={(e) => setAdviserComments(e.target.value)}
                placeholder="Add any additional notes for the client..."
                rows={4}
                className="textarea-input"
              />
            </div>
          </section>
        )}

        {/* Generate Section */}
        {sessionId && (
          <section className="generate-section">
            <h2>4. Generate Report</h2>
            
            <button
              onClick={handleGenerate}
              disabled={generating}
              className="generate-button"
            >
              {generating ? '⏳ Generating...' : '📝 Generate Comprehensive Report'}
            </button>
          </section>
        )}

        {/* Status Messages */}
        {message && (
          <div className={`message ${message.includes('✅') ? 'success' : message.includes('❌') ? 'error' : 'info'}`}>
            {message}
          </div>
        )}
      </main>

      <footer className="App-footer">
        <p>Multi-Provider Support: AJ Bell • Morningstar • Cashflow Documents</p>
      </footer>
    </div>
  );
}

export default App;
