import React, { useState, useRef } from 'react';
import { Upload } from 'lucide-react';

export default function UploadZone({
  onFilesSelected,
  accept,
  multiple = false,
  label = "Drag & drop files here",
  sublabel = "Supports CSV, PDF, Excel, ZIP, etc.",
}) {
  const [isDragActive, setIsDragActive] = useState(false);
  const fileInputRef = useRef(null);

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setIsDragActive(true);
    } else if (e.type === "dragleave") {
      setIsDragActive(false);
    }
  };

  const validateFiles = (files) => {
    const validFiles = [];
    const acceptList = accept ? accept.split(',').map(ext => ext.trim().toLowerCase()) : [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const ext = '.' + file.name.split('.').pop().toLowerCase();
      
      if (acceptList.length === 0 || acceptList.includes(ext) || acceptList.includes(file.type)) {
        validFiles.push(file);
      } else {
        alert(`File format ${file.name} is not supported. Supported formats: ${accept}`);
      }
    }
    return validFiles;
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFiles = Array.from(e.dataTransfer.files);
      const valid = validateFiles(droppedFiles);
      if (valid.length > 0) {
        onFilesSelected(multiple ? valid : [valid[0]]);
      }
    }
  };

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      const selectedFiles = Array.from(e.target.files);
      const valid = validateFiles(selectedFiles);
      if (valid.length > 0) {
        onFilesSelected(multiple ? valid : [valid[0]]);
      }
    }
  };

  const onButtonClick = () => {
    fileInputRef.current.click();
  };

  return (
    <div
      className={`dropzone ${isDragActive ? 'drag-active' : ''}`}
      onDragEnter={handleDrag}
      onDragOver={handleDrag}
      onDragLeave={handleDrag}
      onDrop={handleDrop}
      onClick={onButtonClick}
    >
      <input
        ref={fileInputRef}
        type="file"
        style={{ display: 'none' }}
        multiple={multiple}
        accept={accept}
        onChange={handleFileChange}
      />
      <div className="dropzone-icon">
        <Upload size={32} />
      </div>
      <div className="dropzone-text">{label}</div>
      <div className="dropzone-subtext">{sublabel}</div>
      <button type="button" className="browse-btn">
        Browse Files
      </button>
    </div>
  );
}
