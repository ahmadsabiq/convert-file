import React, { useState } from 'react';
import CsvConverter from './components/CsvConverter';
import FileCompressor from './components/FileCompressor';
import FileMerger from './components/FileMerger';
import PdfToExcelConverter from './components/PdfToExcelConverter';
import { 
  FileDown, 
  Archive, 
  Combine, 
  ShieldCheck, 
  Zap, 
  ServerCrash,
  FileSpreadsheet
} from 'lucide-react';

function App() {
  const [activeTab, setActiveTab] = useState('converter'); // 'converter', 'compressor', 'merger'

  const navigationItems = [
    {
      id: 'converter',
      label: 'CSV Converter',
      icon: <FileDown size={18} />,
      title: 'CSV Converter Suite',
      subtitle: 'Convert CSV spreadsheets to Excel (.xlsx), PDF, Word (.docx), or PowerPoint (.pptx) instantly.'
    },
    {
      id: 'pdf-to-excel',
      label: 'PDF to Excel',
      icon: <FileSpreadsheet size={18} />,
      title: 'PDF to Excel Converter',
      subtitle: 'Convert PDF files to editable Excel spreadsheets using native text parsing or OCR for scans.'
    },
    {
      id: 'compressor',
      label: 'File Compressor',
      icon: <Archive size={18} />,
      title: 'Compress & Optimize Suite',
      subtitle: 'Pack files into ZIP archives or compress image sizes directly using local GPU/canvas drawing.'
    },
    {
      id: 'merger',
      label: 'File Merger Suite',
      icon: <Combine size={18} />,
      title: 'File Merging Suite',
      subtitle: 'Merge multiple PDF documents or combine CSV and Excel sheets into a single workbook.'
    }
  ];

  const currentNav = navigationItems.find(item => item.id === activeTab);

  return (
    <div className="app-container">
      {/* Sidebar navigation */}
      <aside className="sidebar">
        <div className="logo-container">
          <div className="logo-icon">
            <Archive size={22} />
          </div>
          <span className="logo-text">OmniFile</span>
        </div>

        <nav style={{ flexGrow: 1 }}>
          <ul className="nav-list">
            {navigationItems.map(item => (
              <li key={item.id} style={{ display: 'contents' }}>
                <a
                  className={`nav-item ${activeTab === item.id ? 'active' : ''}`}
                  onClick={() => setActiveTab(item.id)}
                >
                  {item.icon}
                  <span>{item.label}</span>
                </a>
              </li>
            ))}
          </ul>
        </nav>

        <div className="sidebar-footer">
          <div>OmniFile Client-Side Engine</div>
          <div style={{ marginTop: '0.25rem', opacity: 0.7 }}>v1.0.0 • Local Only</div>
        </div>
      </aside>

      {/* Main dashboard content */}
      <main className="main-content">
        <header className="content-header">
          <div className="title-badge">
            <Zap size={12} />
            <span>100% Client-Side Engine</span>
          </div>
          <h1 className="main-title">{currentNav.title}</h1>
          <p className="main-subtitle">{currentNav.subtitle}</p>
        </header>

        {/* Dynamic Tool Content */}
        <section className="tool-container">
          {activeTab === 'converter' && <CsvConverter />}
          {activeTab === 'pdf-to-excel' && <PdfToExcelConverter />}
          {activeTab === 'compressor' && <FileCompressor />}
          {activeTab === 'merger' && <FileMerger />}
        </section>

        {/* Privacy & Speed Info Grid */}
        <section className="feature-info-grid" style={{ marginTop: '3rem' }}>
          <div className="feature-info-card">
            <ShieldCheck size={24} />
            <h3 className="feature-info-title">Privacy Guaranteed</h3>
            <p className="feature-info-desc">
              Your files never leave your computer. All processing, parsing, compression, and rendering are done locally in your browser.
            </p>
          </div>
          <div className="feature-info-card">
            <Zap size={24} />
            <h3 className="feature-info-title">Blazing Fast</h3>
            <p className="feature-info-desc">
              Eliminate upload and download times. Local operations run at the speed of your computer's CPU and memory.
            </p>
          </div>
          <div className="feature-info-card">
            <ServerCrash size={24} />
            <h3 className="feature-info-title">Offline Capable</h3>
            <p className="feature-info-desc">
              No internet connection required. Once loaded, you can disconnect from the network and continue converting files.
            </p>
          </div>
        </section>
      </main>
    </div>
  );
}

export default App;
