import React, { useState, useEffect } from 'react';
import JSZip from 'jszip';
import { Archive, Image as ImageIcon, File, Trash2, Sliders, CheckCircle, Download } from 'lucide-react';
import UploadZone from './UploadZone';

export default function FileCompressor() {
  const [activeTab, setActiveTab] = useState('zip'); // 'zip' or 'image'

  // ZIP states
  const [zipFiles, setZipFiles] = useState([]);
  const [zipName, setZipName] = useState('archive');
  const [isZipping, setIsZipping] = useState(false);
  const [zipResult, setZipResult] = useState(null);

  // Image states
  const [imageFile, setImageFile] = useState(null);
  const [imageSrc, setImageSrc] = useState('');
  const [imageDimensions, setImageDimensions] = useState({ w: 0, h: 0 });
  const [quality, setQuality] = useState(80);
  const [maxWidth, setMaxWidth] = useState(1200);
  const [optimizedBlob, setOptimizedBlob] = useState(null);
  const [optimizedSize, setOptimizedSize] = useState(null);
  const [isOptimizing, setIsOptimizing] = useState(false);

  // Clean image resources on unmount
  useEffect(() => {
    return () => {
      if (imageSrc) URL.revokeObjectURL(imageSrc);
    };
  }, [imageSrc]);

  // ZIP logic
  const handleZipFilesSelected = (files) => {
    setZipFiles(prev => [...prev, ...files]);
    setZipResult(null);
  };

  const removeZipFile = (index) => {
    setZipFiles(prev => prev.filter((_, i) => i !== index));
    setZipResult(null);
  };

  const handleCreateZip = async () => {
    if (zipFiles.length === 0) return;
    setIsZipping(true);
    try {
      const zip = new JSZip();
      
      // Prevent duplicate file names inside ZIP by appending a index if needed
      const fileNamesMap = {};
      zipFiles.forEach(file => {
        let name = file.name;
        if (fileNamesMap[name] !== undefined) {
          fileNamesMap[name]++;
          const parts = name.split('.');
          const ext = parts.pop();
          name = `${parts.join('.')}_(${fileNamesMap[name]}).${ext}`;
        } else {
          fileNamesMap[name] = 0;
        }
        zip.file(name, file);
      });

      const content = await zip.generateAsync({ type: "blob" });
      const finalName = `${zipName || 'archive'}.zip`;
      
      const url = URL.createObjectURL(content);
      const link = document.createElement('a');
      link.href = url;
      link.download = finalName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      setZipResult({
        name: finalName,
        size: content.size,
        fileCount: zipFiles.length
      });
    } catch (err) {
      alert("ZIP compression failed: " + err.message);
    } finally {
      setIsZipping(false);
    }
  };

  // Image Optimization logic
  const handleImageSelected = (files) => {
    const file = files[0];
    setImageFile(file);
    setOptimizedBlob(null);
    setOptimizedSize(null);

    if (imageSrc) URL.revokeObjectURL(imageSrc);

    const url = URL.createObjectURL(file);
    setImageSrc(url);

    const img = new Image();
    img.onload = () => {
      setImageDimensions({ w: img.width, h: img.height });
      setMaxWidth(Math.min(img.width, 1920)); // cap default max width at 1920 or original
    };
    img.src = url;
  };

  // Trigger compression when settings change
  useEffect(() => {
    if (!imageSrc || !imageFile) return;
    
    const delayDebounce = setTimeout(() => {
      optimizeImage();
    }, 400);

    return () => clearTimeout(delayDebounce);
  }, [quality, maxWidth, imageSrc]);

  const optimizeImage = () => {
    if (!imageSrc) return;
    setIsOptimizing(true);

    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let targetWidth = img.width;
      let targetHeight = img.height;

      if (img.width > maxWidth) {
        targetWidth = maxWidth;
        targetHeight = Math.round((img.height * maxWidth) / img.width);
      }

      canvas.width = targetWidth;
      canvas.height = targetHeight;

      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, targetWidth, targetHeight);

      // Determine output format
      let format = imageFile.type;
      if (format !== 'image/jpeg' && format !== 'image/png' && format !== 'image/webp') {
        format = 'image/jpeg'; // Fallback
      }

      canvas.toBlob((blob) => {
        if (blob) {
          setOptimizedBlob(blob);
          setOptimizedSize(blob.size);
        }
        setIsOptimizing(false);
      }, format, quality / 100);
    };
    img.src = imageSrc;
  };

  const handleDownloadOptimizedImage = () => {
    if (!optimizedBlob || !imageFile) return;

    const parts = imageFile.name.split('.');
    const ext = parts.pop();
    const finalName = `${parts.join('.')}_optimized.${ext}`;

    const url = URL.createObjectURL(optimizedBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = finalName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
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
          className={`tab-btn ${activeTab === 'zip' ? 'active' : ''}`}
          onClick={() => setActiveTab('zip')}
        >
          ZIP Compressor
        </button>
        <button
          className={`tab-btn ${activeTab === 'image' ? 'active' : ''}`}
          onClick={() => setActiveTab('image')}
        >
          Image Optimizer
        </button>
      </div>

      {/* ZIP Tab */}
      {activeTab === 'zip' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          <UploadZone
            onFilesSelected={handleZipFilesSelected}
            multiple={true}
            label="Drop files here to ZIP"
            sublabel="Supports all formats (Images, Documents, CSV, PDF, etc.)"
          />

          {zipFiles.length > 0 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
              <div className="config-card">
                <div className="config-title">
                  <Archive size={20} />
                  <span>ZIP File Details</span>
                </div>
                <div className="form-group" style={{ maxWidth: '400px' }}>
                  <label className="form-label">ZIP Archive Name</label>
                  <div style={{ display: 'flex', alignItems: 'center' }}>
                    <input
                      type="text"
                      className="form-input"
                      value={zipName}
                      onChange={(e) => setZipName(e.target.value.replace(/[^a-zA-Z0-9_\-]/g, ''))}
                      placeholder="archive"
                      style={{ borderTopRightRadius: 0, borderBottomRightRadius: 0 }}
                    />
                    <span style={{
                      background: 'rgba(255,255,255,0.05)',
                      border: '1px solid var(--border-light)',
                      borderLeft: 'none',
                      padding: '0.75rem 1rem',
                      borderRadius: '0 10px 10px 0',
                      fontSize: '0.9rem',
                      color: 'var(--text-secondary)'
                    }}>.zip</span>
                  </div>
                </div>
              </div>

              {/* Uploaded File List */}
              <div className="file-list">
                <div style={{ fontSize: '0.9rem', fontWeight: '700', color: 'var(--text-secondary)', paddingLeft: '0.25rem' }}>
                  Files inside archive ({zipFiles.length})
                </div>
                {zipFiles.map((f, i) => (
                  <div className="file-item" key={i}>
                    <div className="file-item-icon">
                      <File size={22} />
                    </div>
                    <div className="file-item-info">
                      <div className="file-item-name">{f.name}</div>
                      <div className="file-item-size">{formatSize(f.size)}</div>
                    </div>
                    <div className="file-item-actions">
                      <button className="icon-btn delete" onClick={() => removeZipFile(i)}>
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>

              {/* Create ZIP Action */}
              <button className="action-btn" onClick={handleCreateZip} disabled={isZipping}>
                <Archive size={18} />
                {isZipping ? 'Compressing files...' : 'Compress into ZIP'}
              </button>
            </div>
          )}

          {/* Success Info */}
          {zipResult && (
            <div className="result-card">
              <div className="result-title">
                <CheckCircle size={20} />
                <span>ZIP Created Successfully!</span>
              </div>
              <div className="result-list">
                <div className="result-item">
                  <span style={{ fontSize: '0.9rem', color: 'var(--text-secondary)' }}>
                    {zipResult.fileCount} files compressed into
                  </span>
                  <span style={{ fontWeight: '600' }}>{zipResult.name} ({formatSize(zipResult.size)})</span>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Image Tab */}
      {activeTab === 'image' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          {!imageFile ? (
            <UploadZone
              onFilesSelected={handleImageSelected}
              accept="image/*"
              label="Drop your image here"
              sublabel="Supports JPG, PNG, and WebP images"
            />
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
              <div className="file-item">
                <div className="file-item-icon">
                  <ImageIcon size={28} />
                </div>
                <div className="file-item-info">
                  <div className="file-item-name">{imageFile.name}</div>
                  <div className="file-item-size">
                    Original Size: {formatSize(imageFile.size)} • Dimensions: {imageDimensions.w} x {imageDimensions.h}px
                  </div>
                </div>
                <div className="file-item-actions">
                  <button className="browse-btn" onClick={() => setImageFile(null)}>
                    Change Image
                  </button>
                </div>
              </div>

              {/* Sliders config */}
              <div className="config-card">
                <div className="config-title">
                  <Sliders size={20} />
                  <span>Compression Settings</span>
                </div>
                <div className="grid-cols-2">
                  <div className="form-group">
                    <label className="form-label">Quality (JPEG/WebP)</label>
                    <div className="range-slider-container">
                      <input
                        type="range"
                        min="10"
                        max="100"
                        className="range-slider"
                        value={quality}
                        onChange={(e) => setQuality(Number(e.target.value))}
                      />
                      <span className="slider-val">{quality}%</span>
                    </div>
                  </div>
                  <div className="form-group">
                    <label className="form-label">Max Width (Resize)</label>
                    <div className="range-slider-container">
                      <input
                        type="range"
                        min="200"
                        max={imageDimensions.w || 2000}
                        className="range-slider"
                        value={maxWidth}
                        onChange={(e) => setMaxWidth(Number(e.target.value))}
                      />
                      <span className="slider-val">{maxWidth}px</span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Compression Progress or Info */}
              {isOptimizing ? (
                <div className="status-indicator">
                  <div className="spinner"></div>
                  <span>Estimating compression parameters...</span>
                </div>
              ) : (
                optimizedBlob && (
                  <div className="result-card" style={{ background: 'rgba(99, 102, 241, 0.05)', borderColor: 'rgba(99, 102, 241, 0.2)' }}>
                    <div className="result-title" style={{ color: 'var(--accent-indigo)' }}>
                      <CheckCircle size={20} />
                      <span>Compression Calculated</span>
                    </div>
                    <div className="result-list">
                      <div className="result-item" style={{ flexDirection: 'column', alignItems: 'stretch', gap: '0.5rem' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                          <span style={{ color: 'var(--text-secondary)' }}>Original Size:</span>
                          <span style={{ fontWeight: '600' }}>{formatSize(imageFile.size)}</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                          <span style={{ color: 'var(--text-secondary)' }}>Optimized Size:</span>
                          <span style={{ fontWeight: '700', color: 'var(--accent-emerald)' }}>
                            {formatSize(optimizedSize)}
                          </span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', borderTop: '1px solid rgba(255,255,255,0.05)', paddingTop: '0.5rem' }}>
                          <span style={{ color: 'var(--text-secondary)' }}>Size Saved:</span>
                          <span style={{ fontWeight: '800', color: 'var(--accent-cyan)' }}>
                            {Math.max(0, Math.round(((imageFile.size - optimizedSize) / imageFile.size) * 100))}% reduction
                          </span>
                        </div>
                      </div>
                    </div>
                    <button className="action-btn" onClick={handleDownloadOptimizedImage}>
                      <Download size={18} />
                      Download Optimized Image
                    </button>
                  </div>
                )
              )}

              {/* Side by side Preview */}
              <div className="grid-cols-2" style={{ gap: '2rem' }}>
                <div className="config-card" style={{ alignItems: 'center' }}>
                  <span className="form-label" style={{ marginBottom: '0.5rem' }}>Original Preview</span>
                  {imageSrc && (
                    <img
                      src={imageSrc}
                      alt="Original"
                      style={{ maxWidth: '100%', maxHeight: '300px', objectFit: 'contain', borderRadius: '8px', border: '1px solid var(--border-light)' }}
                    />
                  )}
                </div>
                <div className="config-card" style={{ alignItems: 'center' }}>
                  <span className="form-label" style={{ marginBottom: '0.5rem' }}>Optimized Preview</span>
                  {optimizedBlob && (
                    <img
                      src={URL.createObjectURL(optimizedBlob)}
                      alt="Optimized"
                      style={{ maxWidth: '100%', maxHeight: '300px', objectFit: 'contain', borderRadius: '8px', border: '1px solid var(--border-light)' }}
                    />
                  )}
                </div>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
