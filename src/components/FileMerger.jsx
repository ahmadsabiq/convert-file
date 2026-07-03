import React, { useState } from 'react';
import { PDFDocument } from 'pdf-lib';
import * as XLSX from 'xlsx';
import { FileText, FileSpreadsheet, ArrowUp, ArrowDown, Trash2, Combine, CheckCircle, Download } from 'lucide-react';
import UploadZone from './UploadZone';

export default function FileMerger() {
  const [activeTab, setActiveTab] = useState('pdf'); // 'pdf' or 'excel'

  // PDF states
  const [pdfFiles, setPdfFiles] = useState([]);
  const [isMergingPdf, setIsMergingPdf] = useState(false);
  const [pdfResult, setPdfResult] = useState(null);

  // Excel states
  const [excelFiles, setExcelFiles] = useState([]);
  const [mergeMethod, setMergeMethod] = useState('tabs'); // 'tabs' or 'combine'
  const [isMergingExcel, setIsMergingExcel] = useState(false);
  const [excelResult, setExcelResult] = useState(null);

  // PDF actions
  const handlePdfFilesSelected = (files) => {
    setPdfFiles(prev => [...prev, ...files]);
    setPdfResult(null);
  };

  const movePdf = (index, direction) => {
    const newIndex = index + direction;
    if (newIndex < 0 || newIndex >= pdfFiles.length) return;
    
    const updated = [...pdfFiles];
    const temp = updated[index];
    updated[index] = updated[newIndex];
    updated[newIndex] = temp;
    setPdfFiles(updated);
    setPdfResult(null);
  };

  const removePdf = (index) => {
    setPdfFiles(prev => prev.filter((_, i) => i !== index));
    setPdfResult(null);
  };

  const handleMergePdfs = async () => {
    if (pdfFiles.length < 2) {
      alert("Please upload at least 2 PDF files to merge.");
      return;
    }
    setIsMergingPdf(true);
    try {
      const mergedPdf = await PDFDocument.create();
      
      for (const file of pdfFiles) {
        const bytes = await file.arrayBuffer();
        const pdf = await PDFDocument.load(bytes);
        const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
        copiedPages.forEach((page) => mergedPdf.addPage(page));
      }

      const mergedPdfBytes = await mergedPdf.save();
      const blob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
      const name = `merged_document_${Date.now()}.pdf`;
      
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      setPdfResult({
        name,
        size: blob.size,
        fileCount: pdfFiles.length
      });
    } catch (err) {
      alert("PDF merge failed: " + err.message);
    } finally {
      setIsMergingPdf(false);
    }
  };

  // Excel / CSV actions
  const handleExcelFilesSelected = (files) => {
    setExcelFiles(prev => [...prev, ...files]);
    setExcelResult(null);
  };

  const removeExcel = (index) => {
    setExcelFiles(prev => prev.filter((_, i) => i !== index));
    setExcelResult(null);
  };

  const handleMergeExcel = async () => {
    if (excelFiles.length < 2) {
      alert("Please upload at least 2 files to merge.");
      return;
    }
    setIsMergingExcel(true);
    try {
      const mergedWb = XLSX.utils.book_new();

      if (mergeMethod === 'tabs') {
        // Option 1: Each file becomes a separate sheet in the workbook
        const nameCountMap = {};
        for (let i = 0; i < excelFiles.length; i++) {
          const file = excelFiles[i];
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          
          workbook.SheetNames.forEach((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            
            // Generate unique sheet name (SheetJS requires name length <= 31 and unique)
            let baseName = `${file.name.replace(/\.[^/.]+$/, "")}_${sheetName}`;
            baseName = baseName.substring(0, 27); // leave space for index if duplicate
            
            let finalName = baseName;
            if (nameCountMap[baseName] !== undefined) {
              nameCountMap[baseName]++;
              finalName = `${baseName}_${nameCountMap[baseName]}`;
            } else {
              nameCountMap[baseName] = 0;
            }
            
            XLSX.utils.book_append_sheet(mergedWb, worksheet, finalName);
          });
        }
      } else {
        // Option 2: Combine row-by-row into a single master sheet
        let combinedRows = [];
        let headerRow = null;

        for (let i = 0; i < excelFiles.length; i++) {
          const file = excelFiles[i];
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          
          // Read first sheet of each file
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // AOA format

          if (jsonData.length > 0) {
            if (!headerRow) {
              headerRow = jsonData[0];
              combinedRows.push(headerRow);
            }
            // Append data rows (skipping header of subsequent files)
            const dataRows = i === 0 ? jsonData.slice(1) : jsonData.slice(1); 
            // In case sheet headers differ, we just append directly.
            // For a robust layout, we append all rows.
            combinedRows = [...combinedRows, ...dataRows];
          }
        }

        const ws = XLSX.utils.aoa_to_sheet(combinedRows);
        XLSX.utils.book_append_sheet(mergedWb, ws, "Combined Data");
      }

      const name = `merged_sheets_${Date.now()}.xlsx`;
      XLSX.writeFile(mergedWb, name);

      setExcelResult({
        name,
        size: 0,
        fileCount: excelFiles.length
      });
    } catch (err) {
      alert("Sheet merge failed: " + err.message);
    } finally {
      setIsMergingExcel(false);
    }
  };

  const formatSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
      {/* Sub tabs */}
      <div className="tab-container">
        <button
          className={`tab-btn ${activeTab === 'pdf' ? 'active' : ''}`}
          onClick={() => setActiveTab('pdf')}
        >
          PDF Merger
        </button>
        <button
          className={`tab-btn ${activeTab === 'excel' ? 'active' : ''}`}
          onClick={() => setActiveTab('excel')}
        >
          Excel/CSV Merger
        </button>
      </div>

      {/* PDF Merger Tab */}
      {activeTab === 'pdf' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          <UploadZone
            onFilesSelected={handlePdfFilesSelected}
            multiple={true}
            accept=".pdf"
            label="Drop multiple PDF files here"
            sublabel="Files will be merged in the order listed below"
          />

          {pdfFiles.length > 0 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
              <div className="file-list">
                <div style={{ fontSize: '0.9rem', fontWeight: '700', color: 'var(--text-secondary)', paddingLeft: '0.25rem' }}>
                  Merge Queue ({pdfFiles.length} files)
                </div>
                {pdfFiles.map((f, i) => (
                  <div className="file-item" key={i}>
                    <div className="file-item-icon" style={{ color: 'var(--accent-rose)' }}>
                      <FileText size={22} />
                    </div>
                    <div className="file-item-info">
                      <div className="file-item-name">{f.name}</div>
                      <div className="file-item-size">{formatSize(f.size)}</div>
                    </div>
                    <div className="file-item-actions">
                      <button
                        className="icon-btn"
                        onClick={() => movePdf(i, -1)}
                        disabled={i === 0}
                        title="Move Up"
                      >
                        <ArrowUp size={16} />
                      </button>
                      <button
                        className="icon-btn"
                        onClick={() => movePdf(i, 1)}
                        disabled={i === pdfFiles.length - 1}
                        title="Move Down"
                      >
                        <ArrowDown size={16} />
                      </button>
                      <button
                        className="icon-btn delete"
                        onClick={() => removePdf(i)}
                        title="Remove"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>

              <button
                className="action-btn"
                onClick={handleMergePdfs}
                disabled={pdfFiles.length < 2 || isMergingPdf}
              >
                <Combine size={18} />
                {isMergingPdf ? 'Merging PDFs...' : 'Merge PDF Documents'}
              </button>
            </div>
          )}

          {/* PDF Merge Result */}
          {pdfResult && (
            <div className="result-card">
              <div className="result-title">
                <CheckCircle size={20} />
                <span>PDF Merge Successful!</span>
              </div>
              <div className="result-list">
                <div className="result-item">
                  <span style={{ fontSize: '0.9rem', color: 'var(--text-secondary)' }}>
                    Merged {pdfResult.fileCount} PDFs into
                  </span>
                  <span style={{ fontWeight: '600' }}>{pdfResult.name} ({formatSize(pdfResult.size)})</span>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Excel/CSV Merger Tab */}
      {activeTab === 'excel' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          <UploadZone
            onFilesSelected={handleExcelFilesSelected}
            multiple={true}
            accept=".xlsx,.xls,.csv"
            label="Drop Excel or CSV files here"
            sublabel="Combine sheets or append rows together"
          />

          {excelFiles.length > 0 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
              
              {/* Settings */}
              <div className="config-card">
                <div className="config-title">
                  <Combine size={20} />
                  <span>Merge Settings</span>
                </div>
                <div className="form-group" style={{ maxWidth: '400px' }}>
                  <label className="form-label">Merge Strategy</label>
                  <select
                    className="form-select"
                    value={mergeMethod}
                    onChange={(e) => setMergeMethod(e.target.value)}
                  >
                    <option value="tabs">Export as separate workbook tabs (sheets)</option>
                    <option value="combine">Combine row-by-row into a single master sheet</option>
                  </select>
                </div>
              </div>

              {/* Uploaded File List */}
              <div className="file-list">
                <div style={{ fontSize: '0.9rem', fontWeight: '700', color: 'var(--text-secondary)', paddingLeft: '0.25rem' }}>
                  Sheets to Merge ({excelFiles.length} files)
                </div>
                {excelFiles.map((f, i) => (
                  <div className="file-item" key={i}>
                    <div className="file-item-icon" style={{ color: 'var(--accent-emerald)' }}>
                      <FileSpreadsheet size={22} />
                    </div>
                    <div className="file-item-info">
                      <div className="file-item-name">{f.name}</div>
                      <div className="file-item-size">{formatSize(f.size)}</div>
                    </div>
                    <div className="file-item-actions">
                      <button className="icon-btn delete" onClick={() => removeExcel(i)}>
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>

              {/* Action */}
              <button
                className="action-btn"
                onClick={handleMergeExcel}
                disabled={excelFiles.length < 2 || isMergingExcel}
              >
                <Combine size={18} />
                {isMergingExcel ? 'Merging sheets...' : 'Merge Spreadsheets'}
              </button>
            </div>
          )}

          {/* Excel Merge Result */}
          {excelResult && (
            <div className="result-card">
              <div className="result-title">
                <CheckCircle size={20} />
                <span>Spreadsheets Merged Successfully!</span>
              </div>
              <div className="result-list">
                <div className="result-item">
                  <span style={{ fontSize: '0.9rem', color: 'var(--text-secondary)' }}>
                    Merged {excelResult.fileCount} sheets into
                  </span>
                  <span style={{ fontWeight: '600' }}>{excelResult.name} ({formatSize(excelResult.size)})</span>
                </div>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
