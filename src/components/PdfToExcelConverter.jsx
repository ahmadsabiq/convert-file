import React, { useState, useRef, useEffect } from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import { createWorker } from 'tesseract.js';
import * as XLSX from 'xlsx';
import { 
  FileSpreadsheet, 
  FileText, 
  CheckCircle, 
  AlertTriangle, 
  Loader2, 
  Plus, 
  Trash2, 
  Download, 
  RefreshCw, 
  X,
  Languages,
  Edit2,
  ChevronLeft,
  ChevronRight
} from 'lucide-react';
import UploadZone from './UploadZone';

// Set up PDF.js worker using unpkg CDN
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

export default function PdfToExcelConverter() {
  const [file, setFile] = useState(null);
  const [mode, setMode] = useState('text'); // 'text' or 'ocr'
  const [ocrLang, setOcrLang] = useState('ind+eng'); // Tesseract language code
  const [columnGap, setColumnGap] = useState(25); // Default column gap threshold in pixels
  
  // Processing States
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(null); // { page, totalPages, phase, message, ocrProgress }
  const [error, setError] = useState(null);
  
  // Table Data States
  const [tableData, setTableData] = useState([]); // 2D array [row][col]
  const [headers, setHeaders] = useState([]); // Array of column headers
  const [numPages, setNumPages] = useState(0);
  
  // Grid Editor States
  const [editingCell, setEditingCell] = useState(null); // { rowIndex, colIndex }
  const [editValue, setEditValue] = useState('');
  const [exportFileName, setExportFileName] = useState('');
  const [editingHeaderIndex, setEditingHeaderIndex] = useState(null);
  const [editHeaderValue, setEditHeaderValue] = useState('');
  
  // Page Preview State
  const [previewPageNum, setPreviewPageNum] = useState(1);
  
  // Refs
  const workerRef = useRef(null);
  const cancelRequestedRef = useRef(false);

  // PDF Page Preview Render Hook
  useEffect(() => {
    if (!file || tableData.length === 0 || !numPages) return;
    
    let active = true;
    const renderPreview = async () => {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const page = await pdf.getPage(previewPageNum);
        
        // Wait briefly to make sure the canvas DOM element is available
        await new Promise(resolve => setTimeout(resolve, 50));
        
        const canvas = document.getElementById('pdf-preview-canvas');
        if (!canvas || !active) return;
        
        const context = canvas.getContext('2d');
        const viewport = page.getViewport({ scale: 1.0 }); // 1.0x scale for preview
        
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        
        // White background
        context.fillStyle = '#ffffff';
        context.fillRect(0, 0, canvas.width, canvas.height);
        
        await page.render({ canvasContext: context, viewport: viewport }).promise;
      } catch (e) {
        console.error("Error rendering page preview:", e);
      }
    };
    
    renderPreview();
    
    return () => {
      active = false;
    };
  }, [file, tableData.length, previewPageNum, numPages]);

  // Clean up Tesseract worker on unmount
  useEffect(() => {
    return () => {
      if (workerRef.current) {
        workerRef.current.terminate();
      }
    };
  }, []);

  const handleFileSelected = async (files) => {
    const selectedFile = files[0];
    setFile(selectedFile);
    setError(null);
    setTableData([]);
    setHeaders([]);
    setExportFileName(selectedFile.name.replace(/\.[^/.]+$/, "") + "_converted");

    // Quick pre-check for pages count
    try {
      const arrayBuffer = await selectedFile.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      setNumPages(pdf.numPages);
    } catch (err) {
      setError("Gagal membaca file PDF: " + err.message);
    }
  };

  const handleReset = () => {
    setFile(null);
    setTableData([]);
    setHeaders([]);
    setError(null);
    setProgress(null);
    setIsProcessing(false);
    setNumPages(0);
    cancelRequestedRef.current = false;
    if (workerRef.current) {
      workerRef.current.terminate();
      workerRef.current = null;
    }
  };

  const handleCancel = () => {
    cancelRequestedRef.current = true;
    if (workerRef.current) {
      workerRef.current.terminate();
      workerRef.current = null;
    }
    setIsProcessing(false);
    setProgress(null);
    setError("Proses konversi dibatalkan oleh pengguna.");
  };

  // Convert column index to Excel letters (A, B, C, ..., Z, AA, AB...)
  const getColLetter = (index) => {
    let letter = '';
    let temp = index;
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  };

  // Run PDF parsing
  const handleConvert = async () => {
    if (!file) return;
    
    setIsProcessing(true);
    setError(null);
    cancelRequestedRef.current = false;
    
    try {
      if (mode === 'text') {
        await runTextExtraction();
      } else {
        await runOcrExtraction();
      }
    } catch (err) {
      if (!cancelRequestedRef.current) {
        console.error(err);
        setError("Kesalahan saat mengekstrak data: " + err.message);
        setIsProcessing(false);
      }
    }
  };

  // 1. Text-based PDF extraction
  const runTextExtraction = async () => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const totalPages = pdf.numPages;
    let allRows = [];
    let maxCols = 0;

    for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
      if (cancelRequestedRef.current) return;

      setProgress({
        page: pageNum,
        totalPages,
        phase: 'reading',
        message: `Membaca halaman ${pageNum} dari ${totalPages}...`,
        ocrProgress: 0
      });

      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      const items = textContent.items;

      if (items.length === 0) continue;

      // Extract valid text items with coordinates
      const validItems = items
        .map(item => ({
          text: item.str,
          x: item.transform[4],
          y: item.transform[5],
          width: item.width,
          height: item.height || Math.abs(item.transform[3]) || 10
        }))
        .filter(item => item.text.trim() !== '');

      if (validItems.length === 0) continue;

      // Sort by Y-coordinate descending (top of page first in PDF coordinates)
      validItems.sort((a, b) => b.y - a.y);

      // Group into rows based on Y coordinate with tolerance
      const rows = [];
      let currentY = null;
      let currentRow = [];
      const yTolerance = 6; // pixels tolerance for row alignment

      for (const item of validItems) {
        if (currentY === null) {
          currentY = item.y;
          currentRow.push(item);
        } else if (Math.abs(item.y - currentY) < yTolerance) {
          currentRow.push(item);
        } else {
          rows.push(currentRow);
          currentRow = [item];
          currentY = item.y;
        }
      }
      if (currentRow.length > 0) {
        rows.push(currentRow);
      }

      // Sort items within each row from left to right (X ascending)
      rows.forEach(r => r.sort((a, b) => a.x - b.x));

      // Group row items into cells based on columnGap threshold
      const pageRowsWithCells = [];
      const gapThreshold = columnGap;

      for (const rowItems of rows) {
        const rowCells = [];
        if (rowItems.length === 0) continue;
        
        let currentText = rowItems[0].text;
        let xStart = rowItems[0].x;
        let xEnd = rowItems[0].x + rowItems[0].width;
        
        for (let i = 1; i < rowItems.length; i++) {
          const prevItem = rowItems[i - 1];
          const currItem = rowItems[i];
          
          const gap = currItem.x - (prevItem.x + prevItem.width);
          
          if (gap > gapThreshold) {
            rowCells.push({ text: currentText, xStart, xEnd });
            currentText = currItem.text;
            xStart = currItem.x;
            xEnd = currItem.x + currItem.width;
          } else {
            if (gap > 2) {
              currentText += ' ' + currItem.text;
            } else {
              currentText += currItem.text;
            }
            xEnd = currItem.x + currItem.width;
          }
        }
        rowCells.push({ text: currentText, xStart, xEnd });
        pageRowsWithCells.push(rowCells);
      }

      // Collect cell start X coordinates to cluster into global columns
      const cellXStarts = [];
      pageRowsWithCells.forEach(row => row.forEach(cell => cellXStarts.push(cell.xStart)));
      cellXStarts.sort((a, b) => a - b);

      const colClusters = [];
      const colTolerance = columnGap * 1.5;

      for (const x of cellXStarts) {
        let added = false;
        for (const cluster of colClusters) {
          const avg = cluster.reduce((sum, v) => sum + v, 0) / cluster.length;
          if (Math.abs(x - avg) < colTolerance) {
            cluster.push(x);
            added = true;
            break;
          }
        }
        if (!added) {
          colClusters.push([x]);
        }
      }

      colClusters.sort((a, b) => {
        const avgA = a.reduce((sum, v) => sum + v, 0) / a.length;
        const avgB = b.reduce((sum, v) => sum + v, 0) / b.length;
        return avgA - avgB;
      });

      const colXValues = colClusters.map(cluster => cluster.reduce((sum, v) => sum + v, 0) / cluster.length);
      const numCols = colXValues.length;
      if (numCols > maxCols) maxCols = numCols;

      // Align cells to global columns
      for (const rowCells of pageRowsWithCells) {
        const gridRow = Array(numCols).fill('');
        for (const cell of rowCells) {
          let closestIdx = 0;
          let minDiff = Infinity;
          for (let i = 0; i < numCols; i++) {
            const diff = Math.abs(cell.xStart - colXValues[i]);
            if (diff < minDiff) {
              minDiff = diff;
              closestIdx = i;
            }
          }
          gridRow[closestIdx] = cell.text;
        }
        allRows.push(gridRow);
      }
    }

    if (cancelRequestedRef.current) return;

    if (allRows.length === 0) {
      throw new Error("Tidak ada data teks yang ditemukan. PDF ini mungkin berupa file scan/gambar. Silakan gunakan mode OCR.");
    }

    // Normalize rows length
    const finalCols = maxCols || 1;
    const normalizedRows = allRows.map(row => {
      if (row.length < finalCols) {
        return [...row, ...Array(finalCols - row.length).fill('')];
      } else if (row.length > finalCols) {
        return row.slice(0, finalCols);
      }
      return row;
    });

    setTableData(normalizedRows);
    setHeaders(Array(finalCols).fill('').map((_, i) => `Kolom ${getColLetter(i)}`));
    setIsProcessing(false);
    setProgress(null);
  };

  // 2. OCR-based PDF extraction
  const runOcrExtraction = async () => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const totalPages = pdf.numPages;
    let allRows = [];
    let maxCols = 0;

    setProgress({
      page: 0,
      totalPages,
      phase: 'init_ocr',
      message: 'Menginisialisasi modul OCR (Tesseract)...',
      ocrProgress: 0
    });

    // Create worker with progress logging (with fallback if language fails)
    let worker;
    try {
      worker = await createWorker(ocrLang, 1, {
        logger: m => {
          if (m.status === 'recognizing text') {
            setProgress(prev => {
              if (!prev) return null;
              return {
                ...prev,
                phase: 'ocr',
                ocrProgress: m.progress,
                message: `Halaman ${prev.page}/${prev.totalPages}: Memindai teks (OCR) (${Math.round(m.progress * 100)}%)...`
              };
            });
          }
        }
      });
    } catch (err) {
      console.warn(`Gagal memuat bahasa OCR "${ocrLang}", mencoba fallback ke 'eng':`, err);
      worker = await createWorker('eng', 1, {
        logger: m => {
          if (m.status === 'recognizing text') {
            setProgress(prev => {
              if (!prev) return null;
              return {
                ...prev,
                phase: 'ocr',
                ocrProgress: m.progress,
                message: `Halaman ${prev.page}/${prev.totalPages}: Memindai teks (OCR) (${Math.round(m.progress * 100)}%)...`
              };
            });
          }
        }
      });
    }

    workerRef.current = worker;

    for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
      if (cancelRequestedRef.current) {
        await worker.terminate();
        return;
      }

      setProgress({
        page: pageNum,
        totalPages,
        phase: 'rendering',
        message: `Halaman ${pageNum}/${totalPages}: Memuat & menggambar halaman...`,
        ocrProgress: 0
      });

      const page = await pdf.getPage(pageNum);
      
      // Render page on temporary canvas at high scale for OCR quality
      const viewport = page.getViewport({ scale: 2.0 });
      const canvas = document.createElement('canvas');
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      const context = canvas.getContext('2d');

      // Make background white for OCR contrast
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, canvas.width, canvas.height);

      await page.render({ canvasContext: context, viewport: viewport }).promise;

      if (cancelRequestedRef.current) {
        await worker.terminate();
        return;
      }

      setProgress({
        page: pageNum,
        totalPages,
        phase: 'ocr',
        message: `Halaman ${pageNum}/${totalPages}: Memulai pemindaian teks (OCR)...`,
        ocrProgress: 0
      });

      // Compose onto a secondary canvas with solid white background to eliminate transparency CORS/serialization issues
      const finalCanvas = document.createElement('canvas');
      finalCanvas.width = canvas.width;
      finalCanvas.height = canvas.height;
      const finalCtx = finalCanvas.getContext('2d');
      
      finalCtx.fillStyle = '#ffffff';
      finalCtx.fillRect(0, 0, finalCanvas.width, finalCanvas.height);
      finalCtx.drawImage(canvas, 0, 0);

      // Convert canvas to image data URL to ensure compatibility with Web Worker
      const imageData = finalCanvas.toDataURL('image/png');

      // Run OCR
      const { data } = await worker.recognize(imageData);

      if (cancelRequestedRef.current) {
        await worker.terminate();
        return;
      }

      if (!data) continue;

      // Check if data and words exist, or fall back to line parsing if words is empty but text is not
      if (!data.words || data.words.length === 0) {
        if (data.text && data.text.trim() !== '') {
          // Fallback line parsing (useful if Tesseract layout analyzer detects text block but fails to segment words)
          const lines = data.text.split('\n').map(l => l.trim()).filter(l => l !== '');
          for (const line of lines) {
            const cols = line.split(/\s{2,}|\t/);
            if (cols.length > maxCols) maxCols = cols.length;
            allRows.push(cols);
          }
        }
        continue;
      }

      // Map words with canvas coordinates (Y starts at top, increases down)
      const validWords = data.words
        .map(w => ({
          text: w.text,
          x: w.bbox.x0,
          y: w.bbox.y0,
          width: w.bbox.x1 - w.bbox.x0,
          height: w.bbox.y1 - w.bbox.y0
        }))
        .filter(w => w.text.trim() !== '');

      if (validWords.length === 0) {
        // Fallback line parsing if words are filtered out but raw text exists
        if (data.text && data.text.trim() !== '') {
          const lines = data.text.split('\n').map(l => l.trim()).filter(l => l !== '');
          for (const line of lines) {
            const cols = line.split(/\s{2,}|\t/);
            if (cols.length > maxCols) maxCols = cols.length;
            allRows.push(cols);
          }
        }
        continue;
      }

      // Sort by Y ascending (top to bottom)
      validWords.sort((a, b) => a.y - b.y);

      // Group into rows based on Y coordinate with tolerance
      const rows = [];
      let currentY = null;
      let currentRow = [];
      
      const avgHeight = validWords.reduce((sum, w) => sum + w.height, 0) / validWords.length;
      const yTolerance = avgHeight * 0.7; // 70% of average word height

      for (const word of validWords) {
        if (currentY === null) {
          currentY = word.y;
          currentRow.push(word);
        } else if (Math.abs(word.y - currentY) < yTolerance) {
          currentRow.push(word);
        } else {
          rows.push(currentRow);
          currentRow = [word];
          currentY = word.y;
        }
      }
      if (currentRow.length > 0) {
        rows.push(currentRow);
      }

      // Sort row words from left to right (X ascending)
      rows.forEach(r => r.sort((a, b) => a.x - b.x));

      // Group words into cells based on the columnGap threshold in each row
      const pageRowsWithCells = [];
      const gapThreshold = columnGap;

      for (const rowWords of rows) {
        const rowCells = [];
        if (rowWords.length === 0) continue;
        
        let currentText = rowWords[0].text;
        let xStart = rowWords[0].x;
        let xEnd = rowWords[0].x + rowWords[0].width;
        
        for (let i = 1; i < rowWords.length; i++) {
          const prevWord = rowWords[i - 1];
          const currWord = rowWords[i];
          
          const gap = currWord.x - (prevWord.x + prevWord.width);
          
          if (gap > gapThreshold) {
            rowCells.push({ text: currentText, xStart, xEnd });
            currentText = currWord.text;
            xStart = currWord.x;
            xEnd = currWord.x + currWord.width;
          } else {
            currentText += ' ' + currWord.text;
            xEnd = currWord.x + currWord.width;
          }
        }
        rowCells.push({ text: currentText, xStart, xEnd });
        pageRowsWithCells.push(rowCells);
      }

      // Collect cell start X coordinates to cluster into global columns
      const cellXStarts = [];
      pageRowsWithCells.forEach(row => row.forEach(cell => cellXStarts.push(cell.xStart)));
      cellXStarts.sort((a, b) => a - b);

      const colClusters = [];
      const colTolerance = columnGap * 1.5;

      for (const x of cellXStarts) {
        let added = false;
        for (const cluster of colClusters) {
          const avg = cluster.reduce((sum, v) => sum + v, 0) / cluster.length;
          if (Math.abs(x - avg) < colTolerance) {
            cluster.push(x);
            added = true;
            break;
          }
        }
        if (!added) {
          colClusters.push([x]);
        }
      }

      colClusters.sort((a, b) => {
        const avgA = a.reduce((sum, v) => sum + v, 0) / a.length;
        const avgB = b.reduce((sum, v) => sum + v, 0) / b.length;
        return avgA - avgB;
      });

      const colXValues = colClusters.map(cluster => cluster.reduce((sum, v) => sum + v, 0) / cluster.length);
      const numCols = colXValues.length;
      if (numCols > maxCols) maxCols = numCols;

      // Align cells to global columns
      for (const rowCells of pageRowsWithCells) {
        const gridRow = Array(numCols).fill('');
        for (const cell of rowCells) {
          let closestIdx = 0;
          let minDiff = Infinity;
          for (let i = 0; i < numCols; i++) {
            const diff = Math.abs(cell.xStart - colXValues[i]);
            if (diff < minDiff) {
              minDiff = diff;
              closestIdx = i;
            }
          }
          gridRow[closestIdx] = cell.text;
        }
        allRows.push(gridRow);
      }
    }

    await worker.terminate();
    workerRef.current = null;

    if (cancelRequestedRef.current) return;

    if (allRows.length === 0) {
      throw new Error("Tidak ada data teks yang berhasil dideteksi oleh OCR.");
    }

    // Normalize rows
    const finalCols = maxCols || 1;
    const normalizedRows = allRows.map(row => {
      if (row.length < finalCols) {
        return [...row, ...Array(finalCols - row.length).fill('')];
      } else if (row.length > finalCols) {
        return row.slice(0, finalCols);
      }
      return row;
    });

    setTableData(normalizedRows);
    setHeaders(Array(finalCols).fill('').map((_, i) => `Kolom ${getColLetter(i)}`));
    setIsProcessing(false);
    setProgress(null);
  };

  // GRID EDIT ACTIONS
  const startEditCell = (rIdx, cIdx, val) => {
    setEditingCell({ rowIndex: rIdx, colIndex: cIdx });
    setEditValue(val);
  };

  const saveEditCell = () => {
    if (!editingCell) return;
    const { rowIndex, colIndex } = editingCell;
    const updated = [...tableData];
    updated[rowIndex][colIndex] = editValue;
    setTableData(updated);
    setEditingCell(null);
  };

  const startEditHeader = (index, val) => {
    setEditingHeaderIndex(index);
    setEditHeaderValue(val);
  };

  const saveEditHeader = () => {
    if (editingHeaderIndex === null) return;
    const updated = [...headers];
    updated[editingHeaderIndex] = editHeaderValue;
    setHeaders(updated);
    setEditingHeaderIndex(null);
  };

  // Row Manipulation
  const addRow = () => {
    const newRow = Array(headers.length).fill('');
    setTableData([...tableData, newRow]);
  };

  const deleteRow = (rIdx) => {
    const updated = tableData.filter((_, i) => i !== rIdx);
    setTableData(updated);
  };

  // Column Manipulation
  const addColumn = () => {
    const newHeaders = [...headers, `Kolom ${getColLetter(headers.length)}`];
    const newTableData = tableData.map(row => [...row, '']);
    setHeaders(newHeaders);
    setTableData(newTableData);
  };

  const deleteColumn = (cIdx) => {
    if (headers.length <= 1) {
      alert("Tabel harus memiliki minimal 1 kolom.");
      return;
    }
    const newHeaders = headers.filter((_, i) => i !== cIdx);
    const newTableData = tableData.map(row => row.filter((_, i) => i !== cIdx));
    setHeaders(newHeaders);
    setTableData(newTableData);
  };

  // EXPORT EXCEL
  const handleExport = () => {
    if (tableData.length === 0) return;
    
    // Combine headers and data
    const wsData = [headers, ...tableData];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    
    XLSX.utils.book_append_sheet(wb, ws, "PDF Data");
    
    const fileName = `${exportFileName || 'export'}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }} className="animate-fadeIn">
      {!file ? (
        <UploadZone
          accept=".pdf"
          onFilesSelected={handleFileSelected}
          label="Pilih file PDF"
          sublabel="Tarik & lepas file PDF di sini (Mendukung file ukuran besar)"
        />
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          <div className="file-item">
            <div className="file-item-icon">
              <FileText size={28} style={{ color: 'var(--accent-indigo)' }} />
            </div>
            <div className="file-item-info">
              <div className="file-item-name">{file.name}</div>
              <div className="file-item-size">
                {(file.size / 1024 / 1024).toFixed(2)} MB • {numPages ? `${numPages} Halaman` : 'Membaca halaman...'}
              </div>
            </div>
            <div className="file-item-actions">
              <button className="browse-btn" onClick={handleReset} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <RefreshCw size={14} /> Ganti File
              </button>
            </div>
          </div>

          {/* Config options */}
          {tableData.length === 0 && !isProcessing && (
            <div className="config-card animate-scaleIn" style={{ marginTop: '1.5rem' }}>
              <div className="config-title">
                <span>Konfigurasi Konversi</span>
              </div>
              
              <div className="grid-cols-2">
                <div className="form-group">
                  <label className="form-label">Metode Ekstraksi</label>
                  <select 
                    className="form-select" 
                    value={mode} 
                    onChange={(e) => setMode(e.target.value)}
                  >
                    <option value="text">Ekstraksi Teks (PDF Digital)</option>
                    <option value="ocr">OCR (PDF Hasil Scan / Gambar)</option>
                  </select>
                  <p style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', marginTop: '0.25rem' }}>
                    {mode === 'text' 
                      ? 'Mengekstrak teks native dari PDF digital dengan cepat. Sangat presisi.' 
                      : 'Memproses halaman PDF sebagai gambar dan memindai teks. Lebih lambat, cocok untuk dokumen cetak/scan.'}
                  </p>
                </div>

                {mode === 'ocr' ? (
                  <div className="form-group">
                    <label className="form-label">Bahasa OCR (Tesseract)</label>
                    <select 
                      className="form-select" 
                      value={ocrLang} 
                      onChange={(e) => setOcrLang(e.target.value)}
                    >
                      <option value="ind+eng">Indonesia & Inggris (Rekomendasi)</option>
                      <option value="ind">Bahasa Indonesia saja</option>
                      <option value="eng">English only</option>
                    </select>
                  </div>
                ) : (
                  <div style={{ display: 'none' }} />
                )}
              </div>

              {/* Column Gap Sensitivity Slider */}
              <div className="form-group" style={{ marginTop: '1.25rem' }}>
                <label className="form-label" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span>Sensitivitas Pemisahan Kolom: <strong style={{ color: 'var(--accent-indigo)' }}>{columnGap}px</strong></span>
                  <span style={{ fontSize: '0.75rem', fontWeight: 'normal', color: 'var(--text-muted)' }}>(Default: 25px)</span>
                </label>
                <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginTop: '0.5rem' }}>
                  <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>Rapat (Lebih Banyak Kolom)</span>
                  <input 
                    type="range" 
                    min="5" 
                    max="80" 
                    value={columnGap} 
                    onChange={(e) => setColumnGap(Number(e.target.value))}
                    style={{ 
                      flexGrow: 1, 
                      accentColor: 'var(--accent-indigo)',
                      cursor: 'pointer',
                      height: '6px',
                      background: 'rgba(255,255,255,0.1)',
                      borderRadius: '3px'
                    }}
                  />
                  <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>Lebar (Lebih Sedikit Kolom)</span>
                </div>
                <p style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', marginTop: '0.5rem', lineHeight: '1.4' }}>
                  💡 <strong>Petunjuk:</strong> Geser ke kiri jika ada kolom data yang menyatu / terlewat. Geser ke kanan jika teks dalam sel yang sama terpisah menjadi kolom baru.
                </p>
              </div>

              <button 
                className="action-btn" 
                onClick={handleConvert}
                disabled={isProcessing || !numPages}
                style={{ marginTop: '1rem' }}
              >
                Mulai Konversi
              </button>
            </div>
          )}

          {/* Progress Indicators */}
          {isProcessing && progress && (
            <div className="config-card animate-scaleIn" style={{ marginTop: '1.5rem', textAlign: 'center', padding: '2rem' }}>
              <Loader2 className="animate-spin" size={32} style={{ margin: '0 auto 1rem', color: 'var(--accent-indigo)' }} />
              <div className="progress-status-title" style={{ fontWeight: 700, marginBottom: '0.5rem' }}>
                {progress.message}
              </div>
              
              {/* Progress bar */}
              <div className="progress-bar-container" style={{
                background: 'rgba(255, 255, 255, 0.05)',
                borderRadius: '8px',
                height: '8px',
                width: '100%',
                overflow: 'hidden',
                marginTop: '1rem',
                marginBottom: '1rem'
              }}>
                <div className="progress-bar-fill" style={{
                  background: 'linear-gradient(to right, var(--accent-indigo), var(--accent-purple))',
                  height: '100%',
                  width: `${(progress.page / progress.totalPages) * 100}%`,
                  transition: 'width 0.3s ease'
                }} />
              </div>

              {progress.phase === 'ocr' && progress.ocrProgress > 0 && (
                <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
                  Memproses Halaman {progress.page} dari {progress.totalPages}
                </div>
              )}

              <button 
                className="action-btn secondary" 
                onClick={handleCancel}
                style={{ marginTop: '1.5rem', width: 'auto', display: 'inline-flex' }}
              >
                <X size={16} /> Batal / Hentikan
              </button>
            </div>
          )}

          {/* Error Alert */}
          {error && (
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '0.75rem',
              padding: '1rem',
              background: 'rgba(244, 63, 94, 0.05)',
              border: '1px solid rgba(244, 63, 94, 0.2)',
              borderRadius: '12px',
              color: 'var(--accent-rose)',
              fontSize: '0.85rem',
              marginTop: '1.5rem'
            }} className="animate-scaleIn">
              <AlertTriangle size={20} style={{ flexShrink: 0 }} />
              <div>{error}</div>
            </div>
          )}

          {/* Table Grid Editor with Page Preview */}
          {tableData.length > 0 && (
            <div className="editor-split-container animate-scaleIn" style={{ marginTop: '1.5rem' }}>
              
              {/* Left Panel: Table Grid Editor */}
              <div className="editor-left-panel">
                <div className="config-card" style={{ padding: '1.5rem 1rem' }}>
                  <div className="config-title" style={{ display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', gap: '1rem' }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                      <Edit2 size={16} style={{ color: 'var(--accent-indigo)' }} />
                      Grid Editor Data Konversi
                    </span>
                    
                    {/* Actions row */}
                    <div style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap' }}>
                      <button className="grid-action-btn" onClick={addRow} title="Tambah baris di bawah">
                        <Plus size={14} /> Baris
                      </button>
                      <button className="grid-action-btn" onClick={addColumn} title="Tambah kolom di kanan">
                        <Plus size={14} /> Kolom
                      </button>
                    </div>
                  </div>

                  {/* Re-process / Adjust sensitivity controls */}
                  <div style={{ 
                    background: 'rgba(99, 102, 241, 0.04)', 
                    border: '1px solid rgba(99, 102, 241, 0.15)', 
                    borderRadius: '10px', 
                    padding: '0.8rem 1rem', 
                    marginBottom: '1rem',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '0.6rem'
                  }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '0.5rem' }}>
                      <span style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--text-primary)' }}>
                        Atur Batas Kolom (Jika hasil kurang pas):
                      </span>
                      <button 
                        onClick={handleConvert}
                        className="preview-btn"
                        style={{ 
                          background: 'var(--accent-indigo)', 
                          color: '#fff', 
                          border: 'none', 
                          padding: '0.25rem 0.75rem',
                          fontSize: '0.75rem',
                          fontWeight: 600,
                          borderRadius: '6px',
                          cursor: 'pointer'
                        }}
                      >
                        <RefreshCw size={12} style={{ marginRight: '4px' }} /> Proses Ulang
                      </button>
                    </div>
                    
                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                      <span style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>Lebih Rapat</span>
                      <input 
                        type="range" 
                        min="5" 
                        max="80" 
                        value={columnGap} 
                        onChange={(e) => setColumnGap(Number(e.target.value))}
                        style={{ 
                          flexGrow: 1, 
                          accentColor: 'var(--accent-indigo)',
                          cursor: 'pointer',
                          height: '5px'
                        }}
                      />
                      <span style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>Lebih Lebar ({columnGap}px)</span>
                    </div>
                  </div>

                  {/* Instructions */}
                  <div style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', margin: '0 0 1rem 0' }}>
                    💡 Double-click (atau klik) sel mana saja atau nama kolom untuk mulai mengedit teks langsung.
                  </div>

                  {/* Scrollable table container */}
                  <div className="grid-table-container" style={{
                    maxHeight: '450px',
                    overflow: 'auto',
                    border: '1px solid var(--border-light)',
                    borderRadius: '8px',
                    background: 'rgba(5, 7, 16, 0.4)'
                  }}>
                    <table style={{
                      width: '100%',
                      borderCollapse: 'collapse',
                      fontSize: '0.8rem',
                      color: 'var(--text-primary)',
                      textAlign: 'left'
                    }}>
                      <thead>
                        <tr style={{ borderBottom: '1px solid var(--border-light)', background: 'rgba(15, 23, 42, 0.8)', position: 'sticky', top: 0, zIndex: 10 }}>
                          <th style={{ padding: '0.75rem', width: '50px', textAlign: 'center', borderRight: '1px solid var(--border-light)', color: 'var(--text-muted)' }}>#</th>
                          {headers.map((h, colIdx) => (
                            <th key={colIdx} style={{ padding: '0.75rem', minWidth: '120px', borderRight: '1px solid var(--border-light)' }}>
                              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                                {editingHeaderIndex === colIdx ? (
                                  <input
                                    type="text"
                                    className="grid-header-input"
                                    value={editHeaderValue}
                                    onChange={(e) => setEditHeaderValue(e.target.value)}
                                    onBlur={saveEditHeader}
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter') saveEditHeader();
                                      if (e.key === 'Escape') setEditingHeaderIndex(null);
                                    }}
                                    autoFocus
                                    style={{
                                      background: 'var(--bg-base)',
                                      border: '1px solid var(--accent-indigo)',
                                      color: 'var(--text-primary)',
                                      padding: '0.1rem 0.3rem',
                                      fontSize: '0.8rem',
                                      borderRadius: '4px',
                                      width: '90%'
                                    }}
                                  />
                                ) : (
                                  <span 
                                    onClick={() => startEditHeader(colIdx, h)} 
                                    style={{ cursor: 'pointer', flexGrow: 1, borderBottom: '1px dashed rgba(255,255,255,0.2)' }}
                                    title="Klik untuk edit nama kolom"
                                  >
                                    {h}
                                  </span>
                                )}
                                <button 
                                  onClick={() => deleteColumn(colIdx)} 
                                  style={{ border: 'none', background: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '0 2px' }}
                                  title="Hapus kolom ini"
                                >
                                  <Trash2 size={12} className="hover-red" />
                                </button>
                              </div>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {tableData.map((row, rIdx) => (
                          <tr key={rIdx} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.03)' }}>
                            <td style={{ 
                              padding: '0.5rem', 
                              textAlign: 'center', 
                              borderRight: '1px solid var(--border-light)', 
                              background: 'rgba(255, 255, 255, 0.01)', 
                              color: 'var(--text-muted)',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              gap: '4px'
                            }}>
                              <button 
                                onClick={() => deleteRow(rIdx)}
                                className="row-delete-btn"
                                title="Hapus baris"
                              >
                                <Trash2 size={10} />
                              </button>
                              {rIdx + 1}
                            </td>
                            {row.map((cell, cIdx) => (
                              <td key={cIdx} style={{ padding: '0.35rem 0.5rem', borderRight: '1px solid var(--border-light)', minWidth: '120px' }}>
                                {editingCell && editingCell.rowIndex === rIdx && editingCell.colIndex === cIdx ? (
                                  <input
                                    type="text"
                                    className="grid-cell-input"
                                    value={editValue}
                                    onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={saveEditCell}
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter') saveEditCell();
                                      if (e.key === 'Escape') setEditingCell(null);
                                    }}
                                    autoFocus
                                    style={{
                                      width: '100%',
                                      background: 'rgba(15, 23, 42, 0.95)',
                                      border: '1px solid var(--accent-indigo)',
                                      color: 'var(--text-primary)',
                                      padding: '0.2rem 0.4rem',
                                      fontSize: '0.8rem',
                                      outline: 'none',
                                      borderRadius: '4px'
                                    }}
                                  />
                                ) : (
                                  <div 
                                    onClick={() => startEditCell(rIdx, cIdx, cell)} 
                                    style={{ 
                                      minHeight: '1.2rem', 
                                      cursor: 'pointer',
                                      color: cell ? 'var(--text-primary)' : 'var(--text-muted)'
                                    }}
                                    title="Klik untuk mengedit"
                                  >
                                    {cell || <span style={{ opacity: 0.2 }}>-</span>}
                                  </div>
                                )}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {/* Excel Export Config Area */}
                  <div style={{ 
                    marginTop: '1.5rem', 
                    paddingTop: '1.5rem', 
                    borderTop: '1px solid rgba(255, 255, 255, 0.05)',
                    display: 'flex', 
                    gap: '1rem', 
                    alignItems: 'flex-end',
                    flexWrap: 'wrap'
                  }}>
                    <div className="form-group" style={{ flexGrow: 1, minWidth: '200px' }}>
                      <label className="form-label">Nama File Excel</label>
                      <div style={{ display: 'flex', alignItems: 'center', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid var(--border-light)', borderRadius: '10px', overflow: 'hidden' }}>
                        <input 
                          type="text" 
                          className="form-input" 
                          value={exportFileName} 
                          onChange={(e) => setExportFileName(e.target.value)}
                          style={{ border: 'none', background: 'none', outline: 'none' }}
                          placeholder="Nama file"
                        />
                        <span style={{ paddingRight: '1rem', fontSize: '0.85rem', color: 'var(--text-secondary)', fontWeight: 600 }}>.xlsx</span>
                      </div>
                    </div>
                    
                    <button 
                      className="action-btn" 
                      onClick={handleExport}
                      style={{ width: 'auto', height: '44px', display: 'flex', alignItems: 'center', gap: '0.5rem' }}
                    >
                      <Download size={18} /> Ekspor ke Excel (.xlsx)
                    </button>
                  </div>
                </div>
              </div>

              {/* Right Panel: PDF Page Visual Preview */}
              <div className="editor-right-panel">
                <div className="config-card" style={{ padding: '1.25rem' }}>
                  <div className="config-title">
                    <FileText size={18} style={{ color: 'var(--accent-indigo)' }} />
                    <span>Pratinjau PDF Asli</span>
                  </div>

                  {/* Page Navigator */}
                  <div className="preview-controls">
                    <button
                      className="preview-btn"
                      onClick={() => setPreviewPageNum(prev => Math.max(1, prev - 1))}
                      disabled={previewPageNum <= 1}
                      title="Halaman sebelumnya"
                    >
                      <ChevronLeft size={16} /> Prev
                    </button>
                    <span className="page-num-indicator">
                      Halaman {previewPageNum} / {numPages}
                    </span>
                    <button
                      className="preview-btn"
                      onClick={() => setPreviewPageNum(prev => Math.min(numPages, prev + 1))}
                      disabled={previewPageNum >= numPages}
                      title="Halaman berikutnya"
                    >
                      Next <ChevronRight size={16} />
                    </button>
                  </div>

                  {/* Visual canvas */}
                  <div className="preview-canvas-container">
                    <canvas id="pdf-preview-canvas" style={{ display: 'block', maxWidth: '100%', height: 'auto' }} />
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
