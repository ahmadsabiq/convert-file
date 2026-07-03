import React, { useState } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, BorderStyle } from 'docx';
import pptxgen from 'pptxgenjs';
import { FileSpreadsheet, FileText, Presentation, FileDown, CheckCircle, AlertTriangle } from 'lucide-react';
import UploadZone from './UploadZone';

export default function CsvConverter() {
  const [file, setFile] = useState(null);
  const [csvData, setCsvData] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [stats, setStats] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [conversionResult, setConversionResult] = useState(null);

  const handleFileSelected = (files) => {
    const selectedFile = files[0];
    setFile(selectedFile);
    setConversionResult(null);

    Papa.parse(selectedFile, {
      skipEmptyLines: true,
      complete: (results) => {
        const data = results.data;
        if (data.length > 0) {
          setHeaders(data[0]);
          setCsvData(data);
          setStats({
            rows: data.length,
            cols: data[0].length,
            size: selectedFile.size,
          });
        }
      },
      error: (err) => {
        alert("Error parsing CSV: " + err.message);
      }
    });
  };

  const handleReset = () => {
    setFile(null);
    setCsvData([]);
    setHeaders([]);
    setStats(null);
    setConversionResult(null);
  };

  const downloadBlob = (blob, fileName) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const convertToExcel = () => {
    setIsProcessing(true);
    try {
      const ws = XLSX.utils.aoa_to_sheet(csvData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet 1");
      const name = file.name.replace(/\.csv$/i, '') + '.xlsx';
      
      XLSX.writeFile(wb, name);
      setConversionResult({ format: 'Excel (.xlsx)', fileName: name });
    } catch (err) {
      alert("Excel conversion failed: " + err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const convertToPdf = () => {
    setIsProcessing(true);
    try {
      const doc = new jsPDF({
        orientation: stats.cols > 6 ? 'landscape' : 'portrait',
        unit: 'mm',
        format: 'a4'
      });

      doc.setFontSize(14);
      doc.text(file.name.replace(/\.csv$/i, ''), 14, 15);
      doc.setFontSize(8);
      doc.setTextColor(100);
      doc.text(`Rows: ${stats.rows} | Columns: ${stats.cols}`, 14, 20);

      const head = [headers];
      const body = csvData.slice(1);

      doc.autoTable({
        head: head,
        body: body,
        startY: 25,
        theme: 'grid',
        styles: { fontSize: 8, cellPadding: 2 },
        headStyles: { fillColor: [99, 102, 241] },
        margin: { top: 25, bottom: 15, left: 14, right: 14 }
      });

      const name = file.name.replace(/\.csv$/i, '') + '.pdf';
      doc.save(name);
      setConversionResult({ format: 'PDF (.pdf)', fileName: name });
    } catch (err) {
      alert("PDF conversion failed: " + err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const convertToWord = async () => {
    setIsProcessing(true);
    try {
      // Build a styled Word document table
      const headerRow = new TableRow({
        children: headers.map(header => new TableCell({
          children: [new Paragraph({ text: String(header), bold: true, color: 'FFFFFF' })],
          backgroundColor: '6366F1',
          width: { size: 100 / stats.cols, type: WidthType.PERCENTAGE }
        }))
      });

      // Limit rows in Word to prevent huge doc memory crashes, warning user if so
      const maxWordRows = Math.min(csvData.length, 1000);
      const dataRows = csvData.slice(1, maxWordRows).map(row => {
        return new TableRow({
          children: row.map(cell => new TableCell({
            children: [new Paragraph({ text: cell ? String(cell) : '' })],
            width: { size: 100 / stats.cols, type: WidthType.PERCENTAGE }
          }))
        });
      });

      const table = new Table({
        rows: [headerRow, ...dataRows],
        width: { size: 100, type: WidthType.PERCENTAGE },
      });

      const paragraphs = [
        new Paragraph({ text: file.name.replace(/\.csv$/i, ''), heading: 'Heading1' }),
        new Paragraph({ text: `Total Rows: ${stats.rows} | Columns: ${stats.cols} (Word preview limited to first 1000 rows)` }),
        new Paragraph({ text: '' }),
        table
      ];

      const doc = new Document({
        sections: [{
          properties: {
            page: {
              size: {
                orientation: stats.cols > 6 ? 'landscape' : 'portrait'
              }
            }
          },
          children: paragraphs,
        }]
      });

      const blob = await Packer.toBlob(doc);
      const name = file.name.replace(/\.csv$/i, '') + '.docx';
      downloadBlob(blob, name);
      setConversionResult({ format: 'Word (.docx)', fileName: name });
    } catch (err) {
      alert("Word conversion failed: " + err.message);
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };

  const convertToPptx = () => {
    setIsProcessing(true);
    try {
      const pptx = new pptxgen();
      pptx.layout = 'LAYOUT_16x9';

      // Title Slide
      const slide = pptx.addSlide();
      slide.background = { fill: '0f172a' };
      slide.addText(file.name.replace(/\.csv$/i, ''), {
        x: 0.5, y: 2.2, w: 9.0, h: 1.5,
        fontSize: 32, bold: true, color: 'ffffff', align: 'center'
      });
      slide.addText(`CSV Data Presentation\nRows: ${stats.rows} | Columns: ${stats.cols}`, {
        x: 0.5, y: 3.8, w: 9.0, h: 1.0,
        fontSize: 16, color: '94a3b8', align: 'center'
      });

      // Split data table across slides (approx 10 rows per slide)
      const rowsPerSlide = 10;
      const totalRows = csvData.length;
      let currentIndex = 1;
      let slideNum = 1;

      while (currentIndex < totalRows && slideNum <= 10) { // Limit to 10 slides max to avoid browser crash
        const nextSlide = pptx.addSlide();
        nextSlide.background = { fill: '070913' };
        
        // Header info on slide
        nextSlide.addText(`${file.name.replace(/\.csv$/i, '')} - Part ${slideNum}`, {
          x: 0.5, y: 0.3, w: 9.0, h: 0.5,
          fontSize: 18, bold: true, color: 'ffffff'
        });

        const slideRows = csvData.slice(currentIndex, currentIndex + rowsPerSlide);
        const formattedTableData = [
          headers.map(h => ({ text: String(h), options: { bold: true, color: 'ffffff', fill: '6366F1' } })),
          ...slideRows.map(row => row.map(cell => ({ text: cell ? String(cell) : '', options: { color: 'cbd5e1' } })))
        ];

        nextSlide.addTable(formattedTableData, {
          x: 0.5,
          y: 1.0,
          w: 9.0,
          h: 4.5,
          border: { pt: 1, color: '334155' },
          fill: { color: '0f172a' },
          color: 'cbd5e1',
          fontSize: 9
        });

        currentIndex += rowsPerSlide;
        slideNum++;
      }

      const name = file.name.replace(/\.csv$/i, '') + '.pptx';
      pptx.writeFile({ fileName: name }).then(() => {
        setConversionResult({ format: 'PowerPoint (.pptx)', fileName: name });
        setIsProcessing(false);
      }).catch(err => {
        alert("PowerPoint writing failed: " + err.message);
        setIsProcessing(false);
      });
    } catch (err) {
      alert("PowerPoint conversion failed: " + err.message);
      setIsProcessing(false);
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
      {!file ? (
        <UploadZone
          onFilesSelected={handleFileSelected}
          accept=".csv"
          label="Drop your CSV file here"
          sublabel="Only CSV formats are accepted for sheet conversions"
        />
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          <div className="file-item">
            <div className="file-item-icon">
              <FileSpreadsheet size={28} />
            </div>
            <div className="file-item-info">
              <div className="file-item-name">{file.name}</div>
              <div className="file-item-size">
                {formatSize(stats?.size)} • {stats?.rows} Rows • {stats?.cols} Columns
              </div>
            </div>
            <div className="file-item-actions">
              <button className="browse-btn" onClick={handleReset}>
                Change File
              </button>
            </div>
          </div>

          {stats && stats.cols > 8 && (
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '0.75rem',
              padding: '1rem',
              background: 'rgba(245, 158, 11, 0.05)',
              border: '1px solid rgba(245, 158, 11, 0.2)',
              borderRadius: '12px',
              color: 'var(--accent-amber)',
              fontSize: '0.85rem'
            }}>
              <AlertTriangle size={20} style={{ flexShrink: 0 }} />
              <div>
                This sheet has <strong>{stats.cols} columns</strong>. Conversions to PDF, Word, and PowerPoint will auto-adjust layout to landscape but may look crowded. Excel format is recommended for large datasets.
              </div>
            </div>
          )}

          {/* Actions & Exporters */}
          <div className="config-card">
            <div className="config-title">
              <FileDown size={20} />
              <span>Convert & Export Options</span>
            </div>
            <div className="grid-cols-2">
              <button className="action-btn" onClick={convertToExcel} disabled={isProcessing}>
                <FileSpreadsheet size={18} />
                Export to Excel (.xlsx)
              </button>
              <button className="action-btn" onClick={convertToPdf} disabled={isProcessing}>
                <FileText size={18} />
                Export to PDF (.pdf)
              </button>
              <button className="action-btn" onClick={convertToWord} disabled={isProcessing}>
                <FileText size={18} />
                Export to Word (.docx)
              </button>
              <button className="action-btn" onClick={convertToPptx} disabled={isProcessing}>
                <Presentation size={18} />
                Export to PPTX (.pptx)
              </button>
            </div>
          </div>

          {/* Loader */}
          {isProcessing && (
            <div className="status-indicator">
              <div className="spinner"></div>
              <span>Processing file conversion...</span>
            </div>
          )}

          {/* Success Result */}
          {conversionResult && (
            <div className="result-card">
              <div className="result-title">
                <CheckCircle size={20} />
                <span>Conversion Successful!</span>
              </div>
              <div className="result-list">
                <div className="result-item">
                  <span style={{ fontSize: '0.9rem', color: 'var(--text-secondary)' }}>
                    Exported to {conversionResult.format}
                  </span>
                  <span style={{ fontWeight: '600' }}>{conversionResult.fileName}</span>
                </div>
              </div>
            </div>
          )}

          {/* Table Preview */}
          <div className="config-card" style={{ overflowX: 'auto' }}>
            <div className="config-title">
              <span>CSV Data Preview (First 10 rows)</span>
            </div>
            <table style={{
              width: '100%',
              borderCollapse: 'collapse',
              fontSize: '0.8rem',
              color: 'var(--text-primary)',
              textAlign: 'left'
            }}>
              <thead>
                <tr style={{ borderBottom: '1px solid var(--border-light)' }}>
                  {headers.map((h, i) => (
                    <th key={i} style={{ padding: '0.75rem 1rem', background: 'rgba(255, 255, 255, 0.02)', fontWeight: '700' }}>
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {csvData.slice(1, 11).map((row, rIdx) => (
                  <tr key={rIdx} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.03)' }}>
                    {row.map((cell, cIdx) => (
                      <td key={cIdx} style={{ padding: '0.75rem 1rem', color: 'var(--text-secondary)' }}>
                        {cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
