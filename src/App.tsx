import React, { useState, useRef } from "react";
import { 
  FileUp, 
  FileSpreadsheet, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Loader2,
  FileText,
  ArrowRight
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { cn } from "@/src/lib/utils";

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<"idle" | "uploading" | "success" | "error">("idle");
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile && selectedFile.type === "text/csv") {
      setFile(selectedFile);
      setStatus("idle");
      setErrorMessage("");
    } else if (selectedFile) {
      setErrorMessage("Please select a valid CSV file.");
      setStatus("error");
    }
  };

  const handleUpload = async () => {
    if (!file) return;

    setStatus("uploading");
    setErrorMessage("");

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to convert file");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      setStatus("success");
    } catch (error) {
      console.error("Upload error:", error);
      setErrorMessage(error instanceof Error ? error.message : "An unexpected error occurred.");
      setStatus("error");
    }
  };

  const reset = () => {
    setFile(null);
    setStatus("idle");
    setErrorMessage("");
    setDownloadUrl(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  return (
    <div className="min-h-screen bg-[#f5f5f5] font-sans text-[#1a1a1a]">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center gap-3">
          <div className="bg-[#007AFF] p-2 rounded-lg">
            <FileSpreadsheet className="text-white w-6 h-6" />
          </div>
          <h1 className="text-xl font-semibold tracking-tight">Tempo to Delivery</h1>
        </div>
        <div className="text-sm text-gray-500 font-medium uppercase tracking-wider">
          CSV to Excel Converter
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-6 py-12">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          {/* Left Column: Instructions */}
          <div className="md:col-span-1 space-y-8">
            <section>
              <h2 className="text-xs font-bold uppercase tracking-widest text-gray-400 mb-4">How it works</h2>
              <div className="space-y-6">
                <div className="flex gap-4">
                  <div className="flex-shrink-0 w-6 h-6 rounded-full bg-gray-200 flex items-center justify-center text-xs font-bold">1</div>
                  <p className="text-sm text-gray-600 leading-relaxed">
                    Export your worklogs from Tempo as a <strong>CSV</strong> file.
                  </p>
                </div>
                <div className="flex gap-4">
                  <div className="flex-shrink-0 w-6 h-6 rounded-full bg-gray-200 flex items-center justify-center text-xs font-bold">2</div>
                  <p className="text-sm text-gray-600 leading-relaxed">
                    Upload the file here. We'll aggregate hours by task and epic.
                  </p>
                </div>
                <div className="flex gap-4">
                  <div className="flex-shrink-0 w-6 h-6 rounded-full bg-gray-200 flex items-center justify-center text-xs font-bold">3</div>
                  <p className="text-sm text-gray-600 leading-relaxed">
                    Download the generated <strong>Delivery SHAR</strong> Excel report.
                  </p>
                </div>
              </div>
            </section>

            <section className="bg-blue-50 p-6 rounded-2xl border border-blue-100">
              <h3 className="text-sm font-bold text-blue-900 mb-2 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4" /> Smart Processing
              </h3>
              <p className="text-xs text-blue-800 leading-relaxed opacity-80">
                Our system automatically skips sub-tasks to prevent double-counting of hours, ensuring your reports are 100% accurate.
              </p>
            </section>
          </div>

          {/* Right Column: Upload Area */}
          <div className="md:col-span-2">
            <div className="bg-white rounded-3xl shadow-sm border border-gray-200 overflow-hidden">
              <div className="p-8">
                <AnimatePresence mode="wait">
                  {status === "idle" || status === "error" ? (
                    <motion.div
                      key="upload"
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -10 }}
                      className="space-y-6"
                    >
                      <div 
                        onClick={() => fileInputRef.current?.click()}
                        className={cn(
                          "border-2 border-dashed rounded-2xl p-12 flex flex-col items-center justify-center transition-all cursor-pointer",
                          file ? "border-blue-400 bg-blue-50" : "border-gray-200 hover:border-blue-400 hover:bg-gray-50"
                        )}
                      >
                        <input 
                          type="file" 
                          ref={fileInputRef}
                          onChange={handleFileChange}
                          accept=".csv"
                          className="hidden"
                        />
                        {file ? (
                          <>
                            <FileText className="w-12 h-12 text-blue-500 mb-4" />
                            <p className="text-lg font-medium text-gray-900">{file.name}</p>
                            <p className="text-sm text-gray-500 mt-1">{(file.size / 1024).toFixed(1)} KB</p>
                          </>
                        ) : (
                          <>
                            <div className="w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mb-4">
                              <FileUp className="w-8 h-8 text-blue-500" />
                            </div>
                            <p className="text-lg font-medium text-gray-900">Drop your CSV here</p>
                            <p className="text-sm text-gray-500 mt-1">or click to browse files</p>
                          </>
                        )}
                      </div>

                      {status === "error" && (
                        <div className="bg-red-50 border border-red-100 p-4 rounded-xl flex items-start gap-3 text-red-800">
                          <AlertCircle className="w-5 h-5 flex-shrink-0 mt-0.5" />
                          <p className="text-sm font-medium">{errorMessage}</p>
                        </div>
                      )}

                      <button
                        onClick={handleUpload}
                        disabled={!file}
                        className={cn(
                          "w-full py-4 rounded-xl font-bold text-lg transition-all flex items-center justify-center gap-2",
                          file 
                            ? "bg-[#007AFF] text-white hover:bg-blue-600 shadow-lg shadow-blue-200" 
                            : "bg-gray-100 text-gray-400 cursor-not-allowed"
                        )}
                      >
                        Convert to Excel <ArrowRight className="w-5 h-5" />
                      </button>
                    </motion.div>
                  ) : status === "uploading" ? (
                    <motion.div
                      key="loading"
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      exit={{ opacity: 0 }}
                      className="py-20 flex flex-col items-center justify-center space-y-6"
                    >
                      <Loader2 className="w-16 h-16 text-blue-500 animate-spin" />
                      <div className="text-center">
                        <h3 className="text-xl font-bold text-gray-900">Processing your file</h3>
                        <p className="text-gray-500 mt-2">Aggregating tasks and generating report...</p>
                      </div>
                    </motion.div>
                  ) : (
                    <motion.div
                      key="success"
                      initial={{ opacity: 0, scale: 0.95 }}
                      animate={{ opacity: 1, scale: 1 }}
                      className="py-12 flex flex-col items-center justify-center space-y-8"
                    >
                      <div className="w-20 h-20 bg-green-50 rounded-full flex items-center justify-center">
                        <CheckCircle2 className="w-10 h-10 text-green-500" />
                      </div>
                      <div className="text-center">
                        <h3 className="text-2xl font-bold text-gray-900">Report Ready!</h3>
                        <p className="text-gray-500 mt-2">Your Delivery SHAR Excel file has been generated.</p>
                      </div>
                      
                      <div className="w-full space-y-4">
                        <a
                          href={downloadUrl!}
                          download="Delivery_Report.xlsx"
                          className="w-full bg-green-500 text-white py-4 rounded-xl font-bold text-lg hover:bg-green-600 transition-all flex items-center justify-center gap-2 shadow-lg shadow-green-100"
                        >
                          <Download className="w-5 h-5" /> Download Excel
                        </a>
                        <button
                          onClick={reset}
                          className="w-full bg-white border border-gray-200 text-gray-600 py-4 rounded-xl font-bold text-lg hover:bg-gray-50 transition-all"
                        >
                          Convert another file
                        </button>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            </div>

            {/* Stats / Info Footer */}
            <div className="mt-8 grid grid-cols-2 gap-4">
              <div className="bg-white p-6 rounded-2xl border border-gray-200">
                <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-1">Format</p>
                <p className="text-sm font-semibold text-gray-900">Delivery SHAR v1.0</p>
              </div>
              <div className="bg-white p-6 rounded-2xl border border-gray-200">
                <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-1">Compatibility</p>
                <p className="text-sm font-semibold text-gray-900">Tempo CSV Export</p>
              </div>
            </div>
          </div>
        </div>
      </main>

      <footer className="max-w-4xl mx-auto px-6 py-12 border-t border-gray-200 flex flex-col md:flex-row items-center justify-between gap-4">
        <p className="text-sm text-gray-400">© 2026 Tempo to Delivery Converter. All rights reserved.</p>
        <div className="flex gap-6">
          <a href="#" className="text-sm text-gray-400 hover:text-gray-600 transition-colors">Documentation</a>
          <a href="#" className="text-sm text-gray-400 hover:text-gray-600 transition-colors">Support</a>
        </div>
      </footer>
    </div>
  );
}
