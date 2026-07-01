import React, { useState, useRef, DragEvent, ChangeEvent } from "react";
import { motion, AnimatePresence } from "motion/react";
import {
  Upload,
  FileText,
  FileImage,
  Loader2,
  Download,
  Trash2,
  CheckCircle,
  XCircle,
  FileSpreadsheet,
  Layers,
  ArrowRight,
  RefreshCw,
  Sparkles,
  Info,
  ChevronRight,
  Plus,
  Settings,
  ShieldAlert,
  Sliders,
} from "lucide-react";

interface DocxElement {
  type:
    | "title"
    | "heading1"
    | "heading2"
    | "paragraph"
    | "bullet_list"
    | "numbered_list"
    | "table"
    | "callout";
  text?: string;
  listItems?: string[];
  tableData?: {
    headers: string[];
    rows: string[][];
  };
}

interface DocumentStructure {
  title: string;
  elements: DocxElement[];
}

interface FileItem {
  id: string;
  file: File;
  name: string;
  size: number;
  type: string;
  preview: string; // Base64 thumbnail for images, icon for PDF
  status: "pending" | "processing" | "success" | "error";
  error?: string;
  docxBase64?: string;
  docxName?: string;
  docStructure?: DocumentStructure;
}

export default function App() {
  const [files, setFiles] = useState<FileItem[]>([]);
  const [dragActive, setDragActive] = useState(false);
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
  const [globalProcessing, setGlobalProcessing] = useState(false);
  const [selectedLanguage, setSelectedLanguage] = useState("Auto-Detect");
  const [selectedFormat, setSelectedFormat] = useState("Microsoft Word (.docx)");
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Helper: Convert File to Base64
  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = (error) => reject(error);
    });
  };

  // Helper: Trigger Browser Download for Base64 File
  const downloadBase64File = (base64: string, fileName: string) => {
    try {
      const binaryString = window.atob(base64);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (e) {
      console.error("Error creating download blob:", e);
    }
  };

  // Helper: Format Bytes to human readable size
  const formatBytes = (bytes: number): string => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  };

  // Process selected file additions
  const handleFilesAdded = async (newFileList: FileList | null) => {
    if (!newFileList) return;

    const newItems: FileItem[] = [];

    for (let i = 0; i < newFileList.length; i++) {
      const file = newFileList[i];
      const isPdf = file.type === "application/pdf" || file.name.endsWith(".pdf");
      const isImage = file.type.startsWith("image/");

      // Only accept images and PDFs
      if (!isPdf && !isImage) continue;

      let preview = "";
      if (isImage) {
        try {
          preview = await fileToBase64(file);
        } catch (e) {
          console.error("Failed to create image preview", e);
        }
      }

      newItems.push({
        id: Math.random().toString(36).substring(2, 9) + Date.now(),
        file,
        name: file.name,
        size: file.size,
        type: isPdf ? "application/pdf" : file.type,
        preview,
        status: "pending",
      });
    }

    if (newItems.length > 0) {
      setFiles((prev) => [...prev, ...newItems]);
      // Auto-select the first newly added item for preview if none selected
      if (!selectedFileId) {
        setSelectedFileId(newItems[0].id);
      }
    }
  };

  // Drag and drop event handlers
  const handleDrag = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFilesAdded(e.dataTransfer.files);
    }
  };

  const handleFileInputChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFilesAdded(e.target.files);
    }
  };

  const handleRemoveFile = (id: string, e?: React.MouseEvent) => {
    if (e) {
      e.stopPropagation();
    }
    setFiles((prev) => prev.filter((item) => item.id !== id));
    if (selectedFileId === id) {
      const remaining = files.filter((item) => item.id !== id);
      setSelectedFileId(remaining.length > 0 ? remaining[0].id : null);
    }
  };

  const handleClearAll = () => {
    setFiles([]);
    setSelectedFileId(null);
  };

  // Perform single conversion
  const handleConvertFile = async (id: string, e?: React.MouseEvent) => {
    if (e) {
      e.stopPropagation();
    }

    const itemIndex = files.findIndex((item) => item.id === id);
    if (itemIndex === -1) return;

    const item = files[itemIndex];
    if (item.status === "processing") return;

    // Update state to processing
    setFiles((prev) =>
      prev.map((f) => (f.id === id ? { ...f, status: "processing", error: undefined } : f))
    );

    try {
      const base64 = await fileToBase64(item.file);

      const response = await fetch("/api/convert", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          fileBase64: base64,
          fileName: item.name,
          mimeType: item.type,
        }),
      });

      const data = await response.json();

      if (!response.ok || data.error) {
        throw new Error(data.error || "Failed to convert document.");
      }

      // Update state with success and download info
      setFiles((prev) =>
        prev.map((f) =>
          f.id === id
            ? {
                ...f,
                status: "success",
                docxBase64: data.base64,
                docxName: data.fileName,
                docStructure: data.docStructure,
              }
            : f
        )
      );

      // Trigger automatic browser download immediately
      downloadBase64File(data.base64, data.fileName);
    } catch (err: any) {
      console.error(err);
      setFiles((prev) =>
        prev.map((f) =>
          f.id === id
            ? {
                ...f,
                status: "error",
                error: err.message || "An error occurred during conversion.",
              }
            : f
        )
      );
    }
  };

  // Convert and download all files
  const handleConvertAll = async () => {
    const pendingFiles = files.filter((f) => f.status !== "success" && f.status !== "processing");
    if (pendingFiles.length === 0) return;

    setGlobalProcessing(true);

    for (const pending of pendingFiles) {
      const currentFiles = [...files];
      const isStillAvailable = currentFiles.some((f) => f.id === pending.id);
      if (!isStillAvailable) continue;

      await new Promise<void>(async (resolve) => {
        setFiles((prev) =>
          prev.map((f) =>
            f.id === pending.id ? { ...f, status: "processing", error: undefined } : f
          )
        );

        try {
          const base64 = await fileToBase64(pending.file);
          const response = await fetch("/api/convert", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              fileBase64: base64,
              fileName: pending.name,
              mimeType: pending.type,
            }),
          });

          const data = await response.json();

          if (!response.ok || data.error) {
            throw new Error(data.error || "Failed to convert.");
          }

          setFiles((prev) =>
            prev.map((f) =>
              f.id === pending.id
                ? {
                    ...f,
                    status: "success",
                    docxBase64: data.base64,
                    docxName: data.fileName,
                    docStructure: data.docStructure,
                  }
                : f
            )
          );

          downloadBase64File(data.base64, data.fileName);
        } catch (err: any) {
          console.error(err);
          setFiles((prev) =>
            prev.map((f) =>
              f.id === pending.id
                ? {
                    ...f,
                    status: "error",
                    error: err.message || "An error occurred.",
                  }
                : f
            )
          );
        } finally {
          resolve();
        }
      });
    }

    setGlobalProcessing(false);
  };

  // Dynamic status/progress calculations
  const totalCount = files.length;
  const processedCount = files.filter((f) => f.status === "success").length;
  const progressPercent = totalCount > 0 ? Math.round((processedCount / totalCount) * 100) : 0;
  const isAnyConvertible = files.some((f) => f.status !== "success" && f.status !== "processing");

  const selectedFile = files.find((f) => f.id === selectedFileId);

  return (
    <div className="min-h-screen bg-[#F8FAFC] font-sans text-slate-800 antialiased selection:bg-blue-100 flex flex-col">
      {/* App Header (Matches the Polish Design perfectly) */}
      <header className="h-16 flex items-center justify-between px-6 sm:px-8 bg-white border-b border-slate-200 shadow-xs z-10">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-blue-600 rounded-xl flex items-center justify-center text-white shadow-md shadow-blue-500/10">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          <div>
            <h1 className="text-lg font-bold text-slate-800 font-display leading-tight">DocuBatch AI</h1>
            <p className="text-[11px] text-slate-500 font-medium">Enterprise Image-to-DOCX Engine</p>
          </div>
        </div>
        <div className="flex items-center gap-6">
          <div className="hidden sm:flex flex-col items-end">
            <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">System Status</span>
            <span className="text-xs font-semibold text-emerald-600 flex items-center gap-1.5">
              <span className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></span> OCR Engine Ready
            </span>
          </div>
          <div className="w-9 h-9 bg-slate-100 rounded-full border border-slate-200 flex items-center justify-center text-slate-600 shadow-inner">
            <Settings className="h-4 w-4" />
          </div>
        </div>
      </header>

      {/* Main Content Workspace Grid */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        
        {/* LEFT COLUMN: Upload & Queue List (Takes 7 columns) */}
        <section className="lg:col-span-7 flex flex-col gap-6">
          
          {/* Dropzone Container */}
          <div
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`h-44 border-2 border-dashed rounded-2xl flex flex-col items-center justify-center group cursor-pointer transition-all duration-300 ${
              dragActive
                ? "border-blue-500 bg-blue-50/70 scale-[0.99] shadow-sm"
                : "border-blue-200 bg-blue-50/20 hover:border-blue-400 hover:bg-blue-50/50"
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              multiple
              accept="image/*,application/pdf"
              onChange={handleFileInputChange}
              className="hidden"
            />
            
            <div className="w-12 h-12 bg-white rounded-full shadow-md flex items-center justify-center text-blue-600 mb-3 group-hover:scale-105 transition-transform duration-200">
              <Upload className="h-5 w-5" />
            </div>
            
            <p className="text-slate-700 font-bold font-display text-sm sm:text-base">Drop images or PDF files here</p>
            <p className="text-slate-400 text-xs mt-1 text-center px-4">
              Supports JPG, PNG, WEBP, and multi-page PDF • No file limit
            </p>
          </div>

          {/* File Queue List Container */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-xs flex flex-col overflow-hidden min-h-[350px]">
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
              <h3 className="font-bold text-slate-700 flex items-center gap-2 text-sm sm:text-base">
                Conversion Queue{" "}
                <span className="px-2.5 py-0.5 bg-slate-200/80 text-slate-600 rounded-md text-xs font-semibold">
                  {totalCount} {totalCount === 1 ? "item" : "items"}
                </span>
              </h3>
              {totalCount > 0 && (
                <button
                  onClick={handleClearAll}
                  className="text-xs font-bold text-rose-500 hover:text-rose-600 uppercase tracking-wider transition-colors"
                >
                  Clear All
                </button>
              )}
            </div>

            <div className="flex-1 overflow-y-auto p-4 space-y-3 max-h-[500px]">
              {files.map((file) => {
                const isSelected = selectedFileId === file.id;
                const isPdf = file.type === "application/pdf";

                return (
                  <div
                    key={file.id}
                    onClick={() => setSelectedFileId(file.id)}
                    className={`flex items-center gap-4 p-3 border rounded-xl transition-all duration-200 cursor-pointer ${
                      isSelected
                        ? "bg-blue-50/30 border-blue-200 ring-1 ring-blue-100"
                        : "border-slate-100 hover:bg-slate-50 hover:border-slate-200"
                    }`}
                  >
                    {/* Icon / Thumbnail Box */}
                    <div className="w-10 h-10 rounded-lg flex items-center justify-center shrink-0 overflow-hidden border border-slate-100 bg-slate-50">
                      {file.preview ? (
                        <img src={file.preview} alt="Thumb" className="w-full h-full object-cover" />
                      ) : isPdf ? (
                        <div className="text-rose-600 font-bold text-[10px] uppercase">PDF</div>
                      ) : (
                        <FileImage className="h-4 w-4 text-slate-400" />
                      )}
                    </div>

                    {/* Meta info details */}
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-semibold text-slate-700 truncate leading-snug">
                        {file.name}
                      </div>
                      
                      <div className="flex items-center gap-2 mt-0.5">
                        <span className="text-xs text-slate-400 font-medium">
                          {formatBytes(file.size)}
                        </span>
                        
                        {file.status === "processing" && (
                          <span className="text-xs text-blue-500 font-medium flex items-center gap-1.5 animate-pulse">
                            • Converting layout...
                          </span>
                        )}

                        {file.status === "success" && (
                          <span className="text-xs text-emerald-600 font-medium flex items-center gap-1">
                            • Converted Successfully
                          </span>
                        )}

                        {file.status === "error" && (
                          <span className="text-xs text-rose-500 font-medium flex items-center gap-1 truncate max-w-[150px]">
                            • {file.error || "Failed"}
                          </span>
                        )}

                        {file.status === "pending" && (
                          <span className="text-xs text-slate-400 font-medium">
                            • Queued
                          </span>
                        )}
                      </div>
                    </div>

                    {/* Progress indicators / Action trigger buttons */}
                    <div className="flex items-center shrink-0">
                      {file.status === "processing" && (
                        <div className="w-24 sm:w-32 h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div className="h-full bg-blue-500 rounded-full animate-progress-loading w-3/4"></div>
                        </div>
                      )}

                      {file.status === "success" && file.docxBase64 && (
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            downloadBase64File(file.docxBase64!, file.docxName!);
                          }}
                          className="px-3 py-1.5 text-xs font-bold bg-white text-emerald-600 border border-emerald-200 rounded-lg shadow-xs hover:bg-emerald-50 transition-all flex items-center gap-1"
                        >
                          <Download className="h-3 w-3" />
                          Download
                        </button>
                      )}

                      {file.status === "error" && (
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            handleConvertFile(file.id);
                          }}
                          className="px-2.5 py-1 text-xs font-bold bg-rose-50 text-rose-600 border border-rose-200 rounded-md hover:bg-rose-100 transition"
                        >
                          Retry
                        </button>
                      )}

                      {file.status === "pending" && (
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            handleConvertFile(file.id);
                          }}
                          className="px-3 py-1.5 text-xs font-bold bg-slate-50 text-slate-600 border border-slate-200 hover:bg-blue-600 hover:text-white hover:border-blue-600 rounded-lg transition-all"
                        >
                          Convert
                        </button>
                      )}

                      <button
                        onClick={(e) => handleRemoveFile(file.id, e)}
                        className="ml-2 text-slate-300 hover:text-rose-500 p-1 rounded-md hover:bg-slate-50 transition"
                      >
                        <XCircle className="h-4 w-4" />
                      </button>
                    </div>
                  </div>
                );
              })}

              {files.length === 0 && (
                <div className="flex flex-col items-center justify-center py-16 text-center text-slate-400">
                  <Layers className="h-8 w-8 text-slate-300 mb-2" />
                  <p className="text-sm font-semibold text-slate-600 font-display">No files in queue</p>
                  <p className="text-xs text-slate-400 mt-1 max-w-xs">
                    Upload image sheets or multi-page PDF documents to start your batch conversion workflow.
                  </p>
                </div>
              )}
            </div>
          </div>
        </section>

        {/* RIGHT COLUMN: Settings execution hub, document explorer, security instructions (Takes 5 columns) */}
        <section className="lg:col-span-5 flex flex-col gap-6">
          
          {/* Settings & Execution Panel (Solid Slate block matching Polish Design) */}
          <div className="bg-slate-800 text-white p-6 rounded-2xl shadow-lg flex flex-col min-h-[350px]">
            <h3 className="text-base sm:text-lg font-bold font-display mb-4 flex items-center gap-2">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-blue-400" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M11.3 1.046A1 1 0 0112 2v5h4a1 1 0 01.82 1.573l-7 10A1 1 0 018 18v-5H4a1 1 0 01-.82-1.573l7-10a1 1 0 011.12-.38z" clip-rule="evenodd" />
              </svg>
              Execution Hub
            </h3>

            <div className="space-y-5 flex-1">
              {/* Output format picker */}
              <div>
                <label className="block text-[9px] uppercase font-bold text-slate-400 mb-2 tracking-widest">
                  Batch Output Format
                </label>
                <select
                  value={selectedFormat}
                  onChange={(e) => setSelectedFormat(e.target.value)}
                  className="w-full bg-slate-700 border-none text-white rounded-lg px-3 py-2 text-xs focus:ring-2 focus:ring-blue-500 font-medium outline-hidden"
                >
                  <option>Microsoft Word (.docx)</option>
                  <option>Rich Text (.rtf)</option>
                  <option>Plain Text (.txt)</option>
                </select>
              </div>

              {/* Language tags list selector */}
              <div>
                <label className="block text-[9px] uppercase font-bold text-slate-400 mb-2 tracking-widest">
                  OCR Language Detection
                </label>
                <div className="flex flex-wrap gap-2">
                  {["Auto-Detect", "English", "Spanish", "German"].map((lang) => {
                    const isSel = selectedLanguage === lang;
                    return (
                      <button
                        key={lang}
                        onClick={() => setSelectedLanguage(lang)}
                        className={`px-2.5 py-1 rounded text-[10px] font-bold transition-all ${
                          isSel
                            ? "bg-blue-600 text-white shadow-xs"
                            : "bg-slate-700 hover:bg-slate-650 text-slate-300"
                        }`}
                      >
                        {lang}
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Visual Batch conversion status tracker */}
              <div className="p-4 bg-slate-700/50 rounded-xl">
                <div className="flex justify-between text-xs mb-2">
                  <span className="text-slate-400 font-semibold uppercase tracking-wider text-[9px]">Total Progress</span>
                  <span className="font-bold">{progressPercent}%</span>
                </div>
                <div className="w-full h-1.5 bg-slate-600 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-blue-400 transition-all duration-300"
                    style={{ width: `${progressPercent}%` }}
                  ></div>
                </div>
                <div className="flex justify-between mt-2 text-[10px] text-slate-400">
                  <span>Processed: {processedCount}</span>
                  <span>Total: {totalCount}</span>
                </div>
              </div>
            </div>

            <div className="mt-6 pt-4 border-t border-slate-700/60">
              <button
                onClick={handleConvertAll}
                disabled={globalProcessing || !isAnyConvertible}
                className="w-full bg-blue-600 hover:bg-blue-500 disabled:bg-slate-700 disabled:opacity-40 py-3 rounded-xl font-bold text-sm sm:text-base tracking-wide transition-all duration-200 cursor-pointer shadow-md flex items-center justify-center gap-2"
              >
                {globalProcessing ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin" />
                    Converting Batch...
                  </>
                ) : (
                  <>
                    <Sparkles className="h-4 w-4" />
                    Convert All Files
                  </>
                )}
              </button>
              <p className="text-center text-[10px] text-slate-400 mt-2.5 italic">
                One click converts all queued items to separate .docx files
              </p>
            </div>
          </div>

          {/* Interactive Document Explorer Visualizer */}
          <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-xs flex flex-col min-h-[350px]">
            <div className="border-b border-slate-100 pb-3.5 mb-3.5 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-blue-50 text-blue-600 rounded-lg">
                  <Layers className="h-4 w-4" />
                </div>
                <div>
                  <h3 className="font-display text-sm font-bold text-slate-900 leading-none">
                    Document Explorer
                  </h3>
                  <p className="text-[10px] text-slate-400 mt-0.5">
                    Preview logical parsed typography grid
                  </p>
                </div>
              </div>
            </div>

            {selectedFile ? (
              <div className="flex-1 flex flex-col min-h-0">
                {/* Active queue file identifier */}
                <div className="mb-3.5 bg-slate-50 border border-slate-200/50 rounded-xl p-2.5 flex items-center justify-between gap-3">
                  <div className="min-w-0">
                    <p className="text-xs font-semibold text-slate-700 truncate leading-snug">
                      {selectedFile.name}
                    </p>
                    <p className="text-[9px] text-slate-400">
                      Status:{" "}
                      <span className="font-semibold uppercase text-slate-600">
                        {selectedFile.status}
                      </span>
                    </p>
                  </div>
                  {selectedFile.status === "success" && selectedFile.docxBase64 && (
                    <button
                      onClick={() =>
                        downloadBase64File(selectedFile.docxBase64!, selectedFile.docxName!)
                      }
                      className="p-1.5 rounded-lg bg-emerald-50 text-emerald-700 border border-emerald-100 hover:bg-emerald-600 hover:text-white hover:border-emerald-600 transition duration-150 shadow-inner"
                    >
                      <Download className="h-3.5 w-3.5" />
                    </button>
                  )}
                </div>

                {/* Switch view logic on status */}
                {selectedFile.status === "success" && selectedFile.docStructure ? (
                  <div className="flex-1 overflow-y-auto space-y-4 max-h-[300px] pr-1">
                    {/* Header Docx Title */}
                    <div className="border-l-4 border-blue-500 pl-3 py-0.5">
                      <span className="text-[8px] uppercase tracking-wider font-extrabold text-blue-500 block">
                        Docx Document Title
                      </span>
                      <h4 className="font-display font-bold text-slate-800 text-sm leading-snug">
                        {selectedFile.docStructure.title}
                      </h4>
                    </div>

                    {/* Extracted structures list mapping */}
                    <div className="space-y-3.5 mt-2">
                      {selectedFile.docStructure.elements.map((el, index) => {
                        switch (el.type) {
                          case "title":
                            return (
                              <div key={index} className="border-b border-slate-100 pb-1 mt-1">
                                <span className="text-[7px] bg-indigo-50 text-indigo-600 font-extrabold px-1.5 py-0.5 rounded-xs uppercase block w-max mb-1">
                                  Doc Title
                                </span>
                                <h5 className="font-display font-bold text-slate-850 text-sm">
                                  {el.text}
                                </h5>
                              </div>
                            );
                          case "heading1":
                            return (
                              <div key={index} className="border-b border-slate-100 pb-1 mt-2">
                                <span className="text-[7px] bg-blue-50 text-blue-600 font-extrabold px-1.5 py-0.5 rounded-xs uppercase block w-max mb-1">
                                  Heading 1
                                </span>
                                <h5 className="font-display font-bold text-slate-800 text-xs">
                                  {el.text}
                                </h5>
                              </div>
                            );
                          case "heading2":
                            return (
                              <div key={index} className="mt-1">
                                <span className="text-[7px] bg-slate-100 text-slate-600 font-extrabold px-1.5 py-0.5 rounded-xs uppercase block w-max mb-1">
                                  Heading 2
                                </span>
                                <h6 className="font-display font-bold text-slate-700 text-xs">
                                  {el.text}
                                </h6>
                              </div>
                            );
                          case "paragraph":
                            return (
                              <div key={index} className="bg-slate-50/70 p-2 rounded-lg border border-slate-100/60">
                                <span className="text-[7px] text-slate-400 font-bold uppercase tracking-wider block mb-0.5">
                                  Paragraph Block
                                </span>
                                <p className="text-xs text-slate-600 leading-normal font-sans">
                                  {el.text}
                                </p>
                              </div>
                            );
                          case "bullet_list":
                            return (
                              <div key={index} className="bg-slate-50/70 p-2 rounded-lg border border-slate-100/60 space-y-1">
                                <span className="text-[7px] text-slate-400 font-bold uppercase block mb-0.5">
                                  Bulleted List
                                </span>
                                <ul className="space-y-0.5 pl-1.5">
                                  {el.listItems?.map((li, j) => (
                                    <li key={j} className="text-xs text-slate-600 flex items-start gap-1">
                                      <span className="text-blue-500 font-bold">•</span>
                                      <span>{li}</span>
                                    </li>
                                  ))}
                                </ul>
                              </div>
                            );
                          case "numbered_list":
                            return (
                              <div key={index} className="bg-slate-50/70 p-2 rounded-lg border border-slate-100/60 space-y-1">
                                <span className="text-[7px] text-slate-400 font-bold uppercase block mb-0.5">
                                  Numbered List
                                </span>
                                <ol className="space-y-0.5 pl-1.5">
                                  {el.listItems?.map((li, j) => (
                                    <li key={j} className="text-xs text-slate-600 flex items-start gap-1">
                                      <span className="text-blue-600 font-bold">{j + 1}.</span>
                                      <span>{li}</span>
                                    </li>
                                  ))}
                                </ol>
                              </div>
                            );
                          case "callout":
                            return (
                              <div key={index} className="border-l-4 border-amber-400 bg-amber-50/30 p-2.5 rounded-r-lg border border-slate-100">
                                <span className="text-[7px] text-amber-700 font-extrabold uppercase block mb-0.5">
                                  Highlighted Text Box
                                </span>
                                <p className="text-xs text-slate-700 italic font-sans leading-normal">
                                  {el.text}
                                </p>
                              </div>
                            );
                          case "table":
                            return (
                              <div key={index} className="border border-slate-200 rounded-lg overflow-hidden">
                                <div className="bg-slate-50 px-2 py-1 border-b border-slate-200 flex items-center justify-between">
                                  <span className="text-[7px] text-slate-500 font-bold uppercase flex items-center gap-1">
                                    <FileSpreadsheet className="h-3 w-3 text-emerald-500" />
                                    Parsed Table
                                  </span>
                                </div>
                                <div className="overflow-x-auto">
                                  <table className="min-w-full text-[9px] divide-y divide-slate-150">
                                    <thead className="bg-slate-50">
                                      <tr>
                                        {el.tableData?.headers.map((h, k) => (
                                          <th
                                            key={k}
                                            className="px-2 py-1 text-left font-semibold text-slate-500 uppercase"
                                          >
                                            {h}
                                          </th>
                                        ))}
                                      </tr>
                                    </thead>
                                    <tbody className="bg-white divide-y divide-slate-100">
                                      {el.tableData?.rows.map((row, k) => (
                                        <tr key={k}>
                                          {row.map((cell, m) => (
                                            <td key={m} className="px-2 py-1 text-slate-500 truncate max-w-[120px]">
                                              {cell}
                                            </td>
                                          ))}
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>
                            );
                          default:
                            return null;
                        }
                      })}
                    </div>
                  </div>
                ) : selectedFile.status === "processing" ? (
                  <div className="flex-1 flex flex-col items-center justify-center text-center p-6 space-y-3.5">
                    <Loader2 className="h-8 w-8 text-blue-500 animate-spin" />
                    <div>
                      <h4 className="text-xs font-semibold text-slate-700">Analyzing Layout...</h4>
                      <p className="text-[10px] text-slate-400 max-w-xs mt-0.5">
                        Mapping fonts, bullet lists, structural boundaries, and cell grids.
                      </p>
                    </div>
                  </div>
                ) : selectedFile.status === "error" ? (
                  <div className="flex-1 flex flex-col items-center justify-center text-center p-6 bg-rose-50/20 border border-rose-100 rounded-xl space-y-2">
                    <XCircle className="h-8 w-8 text-rose-500" />
                    <div>
                      <p className="text-xs font-semibold text-rose-800">Preview Unavailable</p>
                      <p className="text-[10px] text-rose-500 max-w-xs mt-0.5">
                        Please try converting this file again.
                      </p>
                    </div>
                  </div>
                ) : (
                  <div className="flex-1 flex flex-col items-center justify-center text-center py-10 space-y-2">
                    <Sparkles className="h-6 w-6 text-slate-350" />
                    <div>
                      <p className="text-xs font-semibold text-slate-500">Awaiting Analysis</p>
                      <p className="text-[10px] text-slate-400 max-w-xs mt-0.5">
                        Click the Convert button on this file card to view the parsed OCR node-tree structures.
                      </p>
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <div className="flex-1 flex flex-col items-center justify-center text-center py-12 text-slate-400">
                <Layers className="h-8 w-8 text-slate-200 mb-2" />
                <p className="text-xs font-semibold text-slate-500">No selection</p>
                <p className="text-[10px] text-slate-400 max-w-xs mt-0.5">
                  Select a document in your queue to inspect its extracted heading hierachy.
                </p>
              </div>
            )}
          </div>

          {/* Secure Storage Policy Card (White rounded card matching the bottom of Polish right side) */}
          <div className="bg-white border border-slate-200 p-5 rounded-2xl flex flex-col gap-3 shadow-xs">
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 rounded-full bg-amber-50 flex items-center justify-center text-amber-500">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              </div>
              <span className="text-xs font-bold text-slate-700 font-display">Storage Policy</span>
            </div>
            <p className="text-[11px] text-slate-400 leading-relaxed font-medium">
              All files are processed using local secure memory buffers. Converted assets are automatically purged from system memory 60 minutes after execution.
            </p>
          </div>

        </section>

      </main>

      {/* Polish Design Bottom Footer Action Bar */}
      <footer className="h-14 bg-slate-150 border-t border-slate-200 px-6 sm:px-8 flex items-center justify-between text-slate-400 text-xs shrink-0 bg-white">
        <div className="flex gap-4 sm:gap-6">
          <div className="text-[11px] flex items-center gap-1.5 font-medium">
            <span className="w-1.5 h-1.5 bg-slate-400 rounded-full"></span>
            API v2.4.1
          </div>
          <div className="text-[11px] flex items-center gap-1.5 font-medium">
            <span className="w-1.5 h-1.5 bg-slate-400 rounded-full"></span>
            Secured with TLS 1.3
          </div>
        </div>
        
        {processedCount > 0 ? (
          <button
            onClick={() => {
              files.forEach((file) => {
                if (file.status === "success" && file.docxBase64) {
                  downloadBase64File(file.docxBase64, file.docxName!);
                }
              });
            }}
            className="flex items-center gap-2 text-xs font-bold text-slate-600 hover:text-slate-900 transition-colors cursor-pointer"
          >
            Download Processed ({processedCount})
            <Download className="h-3.5 w-3.5" />
          </button>
        ) : (
          <div className="text-slate-350 text-[11px] font-semibold flex items-center gap-1">
            Reconstruct to download
          </div>
        )}
      </footer>
    </div>
  );
}
