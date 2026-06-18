/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as React from "react";
import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  UploadCloud,
  FileSpreadsheet,
  AlertTriangle,
  CheckCircle,
  RefreshCw,
  Search,
  Trash2,
  Copy,
  Check,
  Eye,
  EyeOff,
  Download,
  Sparkles,
  Filter,
  Info,
  HelpCircle,
  Plus
} from "lucide-react";
import { demoFile1Rows, demoFile2Rows } from "./demoData";
import { FileData, ComparisonResult, ColumnMapping, TransactionItem } from "./types";

export default function App() {
  // Loaded File States
  const [file1, setFile1] = useState<FileData | null>(null);
  const [file2, setFile2] = useState<FileData | null>(null);

  // Loading indicator states
  const [loading1, setLoading1] = useState(false);
  const [loading2, setLoading2] = useState(false);

  // Column Mapping State
  const [mapping, setMapping] = useState<ColumnMapping>({
    file1IdKey: "",
    file1MachineKey: "",
    file1PhoneKey: "",
    file2SessionIdKey: ""
  });

  // Search & Filter options
  const [searchTerm, setSearchTerm] = useState("");
  const [machineFilter, setMachineFilter] = useState<"ALL" | "TLV 1" | "TLV 2">("ALL");

  // Interaction feedbacks
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [copiedPhone, setCopiedPhone] = useState<string | null>(null);
  const [revealedPhones, setRevealedPhones] = useState<Record<string, boolean>>({});

  // Help Modal State
  const [showHelp, setShowHelp] = useState(false);

  // Sound play on alert
  const [playSound, setPlaySound] = useState(true);

  // Auto-detect columns whenever high level file state shifts
  useEffect(() => {
    if (file1) {
      const keys = file1.columns;
      // Look for id
      const idKey = keys.find(k => /^id$/i.test(k)) || 
                    keys.find(k => /mizahe|מזהה/i.test(k)) || 
                    keys[0] || "";

      // Look for machine/machineName
      const machineKey = keys.find(k => /machinename|machine_name|machine|מכונה|שם מכונה/i.test(k)) || 
                         keys.find(k => /name|שם/i.test(k)) || 
                         keys[3] || keys[1] || "";

      // Look for phone
      const phoneKey = keys.find(k => /phone|mobile|tel|טלפון|נייד|סלולרי/i.test(k)) || 
                       keys[9] || keys[2] || "";

      setMapping(prev => ({
        ...prev,
        file1IdKey: idKey,
        file1MachineKey: machineKey,
        file1PhoneKey: phoneKey
      }));
    }
  }, [file1]);

  useEffect(() => {
    if (file2) {
      const keys = file2.columns;
      // Look for Session ID
      const sessIdKey = keys.find(k => /session\s*id|session_id|sessionid|מזהה\s*סשן|סשן/i.test(k)) || 
                        keys.find(k => /^id$/i.test(k)) ||
                        keys[0] || "";

      setMapping(prev => ({
        ...prev,
        file2SessionIdKey: sessIdKey
      }));
    }
  }, [file2]);

  // Load Preset Demo Data
  const loadDemoData = () => {
    // File 1 Demo
    const file1Cols = Object.keys(demoFile1Rows[0]);
    setFile1({
      fileName: "TLV_Transactions_Daily.xlsx (נתוני דוגמה)",
      columns: file1Cols,
      rows: demoFile1Rows
    });

    // File 2 Demo
    const file2Cols = Object.keys(demoFile2Rows[0]);
    setFile2({
      fileName: "Session_Logs_Export.xlsx (נתוני דוגמה)",
      columns: file2Cols,
      rows: demoFile2Rows
    });
  };

  // Drag and Drop handlers for File A
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const processFile = (file: File, fileIndex: 1 | 2) => {
    const isFile1 = fileIndex === 1;
    if (isFile1) setLoading1(true);
    else setLoading2(true);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: "" });

        if (rows.length === 0) {
          alert(`הקובץ ${file.name} ריק או שאינו מכיל שורות נתונים.`);
          if (isFile1) setLoading1(false);
          else setLoading2(false);
          return;
        }

        const columns = Object.keys(rows[0]);
        const fileData: FileData = {
          fileName: file.name,
          columns,
          rows
        };

        if (isFile1) {
          setFile1(fileData);
          setLoading1(false);
        } else {
          setFile2(fileData);
          setLoading2(false);
        }
      } catch (err) {
        console.error(err);
        alert(`שגיאה בקריאת הקובץ: ${err instanceof Error ? err.message : String(err)}`);
        if (isFile1) setLoading1(false);
        else setLoading2(false);
      }
    };
    reader.onerror = () => {
      alert("שגיאה בטעינת הקובץ מהדיסק");
      if (isFile1) setLoading1(false);
      else setLoading2(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDropFile1 = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0], 1);
    }
  };

  const handleDropFile2 = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0], 2);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>, fileIndex: 1 | 2) => {
    if (e.target.files && e.target.files[0]) {
      processFile(e.target.files[0], fileIndex);
    }
  };

  // Phone Formatter function matching user specs
  const formatPhone = (phoneVal: any) => {
    if (phoneVal === undefined || phoneVal === null || phoneVal === "") {
      return {
        original: "לא הוזן",
        formatted: "אין טלפון",
        masked: "אין טלפון"
      };
    }
    const cleanStr = String(phoneVal).trim();
    // Keep numbers and optional plus
    const digitsOnly = cleanStr.replace(/[^\d+]/g, "");
    
    // Remove "+972" or "972" at start
    let baseDigits = digitsOnly.replace(/^\+?972/, "");
    
    // If it starts with 5, prepend 0 to make it 05...
    if (baseDigits.startsWith("5")) {
      baseDigits = "0" + baseDigits;
    }

    // Prepare masked phone
    let masked = baseDigits;
    if (baseDigits.startsWith("05")) {
      // replace middle with asterisks: e.g. 05 + ***** + remaining
      const rest = baseDigits.slice(7); // starts from index 7
      masked = "05*****" + rest;
    } else {
      // secondary fallback just in case
      masked = "05*****" + baseDigits.slice(baseDigits.length > 5 ? 5 : 2);
    }

    return {
      original: cleanStr,
      formatted: baseDigits,
      masked: masked
    };
  };

  // Perform Comparison Logic
  const comparison = useMemo((): ComparisonResult | null => {
    if (!file1) return null;

    const idKey = mapping.file1IdKey;
    const machineKey = mapping.file1MachineKey;
    const phoneKey = mapping.file1PhoneKey;
    const sessKey2 = mapping.file2SessionIdKey;

    if (!idKey || !machineKey) {
      return {
        totalFile1: file1.rows.length,
        totalFile2: file2 ? file2.rows.length : 0,
        matchedCount: 0,
        filteredCount: 0,
        skippedCount: file1.rows.length,
        missingItems: []
      };
    }

    // Step 1: Filter File 1 rows where machine is 'TLV 1' or 'TLV 2'
    let skippedCount = 0;
    const filteredRows = file1.rows.filter(row => {
      const machVal = String(row[machineKey] || "").trim().toUpperCase();
      const match = machVal === "TLV 1" || machVal === "TLV 2";
      if (!match) skippedCount++;
      return match;
    });

    // Step 2: Extract set of session IDs from File 2 (or empty set if file 2 not loaded)
    const file2SessionIds = new Set<string>();
    if (file2 && sessKey2) {
      file2.rows.forEach(row => {
        const val = String(row[sessKey2] || "").trim().toLowerCase();
        if (val) {
          file2SessionIds.add(val);
        }
      });
    }

    // Step 3: Identify the missing items
    const missing: TransactionItem[] = [];
    let matchedCount = 0;

    filteredRows.forEach(row => {
      const idVal = String(row[idKey] || "").trim();
      const idLower = idVal.toLowerCase();

      // Check if ID exists in File 2's session list
      if (file2 && file2SessionIds.has(idLower)) {
        matchedCount++;
      } else {
        // missing!
        missing.push({
          id: idVal,
          machineName: String(row[machineKey] || "").trim(),
          phone: String(row[phoneKey] || "").trim(),
          rawRow: row
        });
      }
    });

    return {
      totalFile1: file1.rows.length,
      totalFile2: file2 ? file2.rows.length : 0,
      matchedCount,
      filteredCount: filteredRows.length,
      skippedCount,
      missingItems: missing
    };
  }, [file1, file2, mapping]);

  // Audio effect toggle when new missing items appear
  useEffect(() => {
    if (comparison && comparison.missingItems.length > 0 && playSound) {
      try {
        // Subtle clean modern synthetic notification beep via AudioContext to be error-safe
        const ctx = new (window.AudioContext || (window as any).webkitAudioContext)();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = "sine";
        osc.frequency.setValueAtTime(587.33, ctx.currentTime); // D5 note
        gain.gain.setValueAtTime(0.08, ctx.currentTime);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        setTimeout(() => {
          osc.stop();
        }, 180);
      } catch (e) {
        // Clean catch for environments where audio blocks are active
        console.log("Audio not allowed yet or not supported.");
      }
    }
  }, [comparison?.missingItems.length]);

  // Clean formatted search & filter list
  const finalFilteredList = useMemo(() => {
    if (!comparison) return [];
    
    let list = comparison.missingItems;

    // Apply machine filter
    if (machineFilter !== "ALL") {
      list = list.filter(item => item.machineName.toUpperCase() === machineFilter);
    }

    // Apply search filter
    if (searchTerm.trim() !== "") {
      const term = searchTerm.toLowerCase();
      list = list.filter(item => {
        const idMatch = item.id.toLowerCase().includes(term);
        const machMatch = item.machineName.toLowerCase().includes(term);
        const parsed = formatPhone(item.phone);
        const phoneMatch = parsed.formatted.includes(term) || parsed.original.includes(term);
        return idMatch || machMatch || phoneMatch;
      });
    }

    return list;
  }, [comparison, machineFilter, searchTerm]);

  // Copy helpers
  const handleCopy = (text: string, type: "id" | "phone") => {
    navigator.clipboard.writeText(text);
    if (type === "id") {
      setCopiedId(text);
      setTimeout(() => setCopiedId(null), 1500);
    } else {
      setCopiedPhone(text);
      setTimeout(() => setCopiedPhone(null), 1500);
    }
  };

  const togglePhoneReveal = (id: string) => {
    setRevealedPhones(prev => ({
      ...prev,
      [id]: !prev[id]
    }));
  };

  // Reset App state
  const clearAll = () => {
    setFile1(null);
    setFile2(null);
    setSearchTerm("");
    setMachineFilter("ALL");
    setRevealedPhones({});
    setMapping({
      file1IdKey: "",
      file1MachineKey: "",
      file1PhoneKey: "",
      file2SessionIdKey: ""
    });
  };

  // Exports missing transactions to a fresh Excel file
  const exportMissingToExcel = () => {
    if (!comparison || comparison.missingItems.length === 0) return;

    // Build plain readable structure for the output Excel file
    const dataToExport = comparison.missingItems.map((item, index) => {
      const phoneDetails = formatPhone(item.phone);
      return {
        "מספר סידורי": index + 1,
        "מזהה עסקאות (ID)": item.id,
        "שם מכונה": item.machineName,
        "טלפון מקורי": phoneDetails.original,
        "טלפון מפורמט מוסתר": phoneDetails.masked,
        "טלפון מפורמט מלא": phoneDetails.formatted,
        "תאריך עסקה": item.rawRow.created || item.rawRow.createdDate || "לא צוין",
        "סטטוס מקורי": item.rawRow.status || "לא צוין",
        "סכום": item.rawRow.fiat ? `${item.rawRow.fiat} ${item.rawRow.fiatCode || ""}` : "לא צוין"
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    worksheet["!dir"] = "rtl"; // set RTL direction in excel metadata
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "TLV דוח מזהים חסרים");
    
    XLSX.writeFile(workbook, `TLV_Missing_IDs_Report_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col font-sans transition-colors antialiased">
      {/* Top Professional Global Header */}
      <header className="bg-gradient-to-r from-slate-900 via-slate-850 to-indigo-950 text-white shadow-md border-b border-indigo-900/40">
        <div className="max-w-7xl mx-auto px-4 py-5 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600/30 p-2.5 rounded-xl border border-indigo-500/20 text-indigo-400">
              <FileSpreadsheet className="w-8 h-8" />
            </div>
            <div className="text-right sm:text-right">
              <h1 className="text-xl sm:text-2xl font-bold tracking-tight bg-gradient-to-l from-white via-slate-100 to-indigo-200 bg-clip-text text-transparent">
                מערכת השוואת עסקאות TLV
              </h1>
              <p className="text-xs sm:text-sm text-slate-400 mt-1">
                כלי מקצועי לסיכון, סינון מכונות TLV 1 & TLV 2 והצלבת מזהי Session ID חסרים
              </p>
            </div>
          </div>
          
          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={loadDemoData}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 active:bg-indigo-700 text-white rounded-lg transition-all duration-200 border border-indigo-500/30 shadow-sm"
              id="load-demo-btn"
            >
              <Sparkles className="w-4 h-4 text-amber-300" />
              <span>טען קבצי דוגמה</span>
            </button>

            <button
              onClick={() => setShowHelp(!showHelp)}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-slate-850 hover:bg-slate-800 text-slate-300 rounded-lg transition-colors border border-slate-700"
              id="help-btn"
            >
              <HelpCircle className="w-4 h-4" />
              <span>עזרה והסבר</span>
            </button>

            <button
              onClick={clearAll}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-rose-950/40 text-rose-300 hover:bg-rose-900/40 rounded-lg transition-colors border border-rose-900/30"
              id="reset-btn"
              title="נקה הכל"
            >
              <Trash2 className="w-4 h-4" />
              <span>נקה קבצים</span>
            </button>
          </div>
        </div>
      </header>

      {/* Help Banner / Explanation Row if toggled */}
      {showHelp && (
        <div className="bg-indigo-50 border-b border-indigo-100 py-4 px-4 transition-all">
          <div className="max-w-7xl mx-auto flex items-start gap-3">
            <div className="bg-indigo-100 p-1.5 rounded-lg text-indigo-700 shrink-0 mt-0.5">
              <Info className="w-5 h-5" />
            </div>
            <div className="text-sm text-slate-700 space-y-2 leading-relaxed">
              <h3 className="font-bold text-indigo-900">איך משתמשים במערכת?</h3>
              <p>
                1. <strong>מעלים קובץ ראשון (קובץ העסקאות)</strong>: קובץ זה מכיל את רשימת העסקאות, מזהה העסקאות (עמודה שמזוהה אוטומטית כ-`id`), שם המכונה, ופרטי הקשר.
              </p>
              <p>
                2. <strong>מעלים קובץ שני (קובץ השוואה)</strong>: קובץ זה מכיל את המזהים שכבר קיימים במערכת (עמודה שמזוהה אוטומטית כ-`Session ID`).
              </p>
              <p>
                3. <strong>הסינון וההצלבה מתבצעים מיד</strong>: המערכת מסננת אוטומטית את שורות קובץ העסקאות ומשאירה רק עסקאות ששם המכונה שלהן הוא <code className="bg-indigo-50 px-1 py-0.5 rounded text-indigo-700 font-semibold font-mono">TLV 1</code> או <code className="bg-indigo-50 px-1 py-0.5 rounded text-indigo-700 font-semibold font-mono">TLV 2</code>.
              </p>
              <p>
                4. <strong>מזהים חסרים</strong>: מתוך העסקאות של TLV 1 ו-TLV 2 הנותרות, המערכת מזהה אילו עסקאות אינן מופיעות בקובץ השני, מתריעה עליהן ומציגה את פרטיהן.
              </p>
              <p>
                5. <strong>פרמוט והסתרה של מספרי טלפון</strong>: מספרי הטלפון המגיעים בפורמט בינלאומי עם קידומת <code className="font-mono bg-indigo-50 text-indigo-700 px-1 rounded">972</code> הופכים למספרים ישראליים מקומיים המתחילים ב-<code className="font-mono">0</code> וירונדרו בצורה מוסתרת כ-<code className="font-mono">05*****</code> עם אופציה לחשיפת המספר המלא בלחיצה.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 py-8 flex flex-col gap-8">
        
        {/* Step 1: Upload Zones */}
        <section className="grid grid-cols-1 lg:grid-cols-2 gap-6" id="upload-section">
          {/* File 1 Upload Area */}
          <div 
            onDragOver={handleDragOver}
            onDrop={handleDropFile1}
            className={`bg-white rounded-2xl border-2 border-dashed p-6 flex flex-col items-center justify-center transition-all ${
              file1 
                ? "border-emerald-200 bg-emerald-50/10 shadow-sm"
                : "border-slate-300 hover:border-indigo-400 bg-white"
            }`}
          >
            <div className="flex items-center justify-between w-full mb-4">
              <span className="bg-indigo-50 text-indigo-700 text-xs px-2.5 py-1 rounded-full font-bold">קובץ 1</span>
              {file1 && (
                <span className="flex items-center gap-1 text-emerald-600 text-xs font-semibold">
                  <CheckCircle className="w-4 h-4" /> קובץ נטען בהצלחה
                </span>
              )}
            </div>

            {loading1 ? (
              <div className="py-8 flex flex-col items-center gap-3">
                <RefreshCw className="w-10 h-10 text-indigo-600 animate-spin" />
                <p className="text-sm text-slate-500">קורא ומפענח את קובץ האקסל...</p>
              </div>
            ) : file1 ? (
              <div className="w-full text-right">
                <div className="flex items-start gap-3">
                  <div className="bg-emerald-100 text-emerald-700 p-3 rounded-xl">
                    <FileSpreadsheet className="w-8 h-8" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <h3 className="font-bold text-slate-900 truncate" title={file1.fileName}>
                      {file1.fileName}
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      {file1.rows.length.toLocaleString()} שורות נתונים שנטענו • {file1.columns.length} עמודות שנמצאו
                    </p>
                  </div>
                  <button 
                    onClick={() => setFile1(null)} 
                    className="p-1 text-slate-450 hover:text-rose-600 rounded hover:bg-rose-50 transition-colors"
                    title="הסר קובץ"
                  >
                    <Trash2 className="w-5 h-5" />
                  </button>
                </div>

                {/* Auto Column Mapping Viewer for File 1 */}
                <div className="mt-4 pt-4 border-t border-slate-100 grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">עמודת מזהה (ID):</label>
                    <select
                      value={mapping.file1IdKey}
                      onChange={(e) => setMapping(p => ({ ...p, file1IdKey: e.target.value }))}
                      className="w-full text-xs bg-slate-50 border border-slate-200 rounded px-2 py-1.5 font-mono text-slate-700 outline-none focus:border-indigo-500 focus:bg-white"
                    >
                      {file1.columns.map(col => (
                        <option key={col} value={col}>{col}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">עמודת מכונה:</label>
                    <select
                      value={mapping.file1MachineKey}
                      onChange={(e) => setMapping(p => ({ ...p, file1MachineKey: e.target.value }))}
                      className="w-full text-xs bg-slate-50 border border-slate-200 rounded px-2 py-1.5 font-mono text-slate-700 outline-none focus:border-indigo-500 focus:bg-white"
                    >
                      {file1.columns.map(col => (
                        <option key={col} value={col}>{col}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">עמודת טלפון:</label>
                    <select
                      value={mapping.file1PhoneKey}
                      onChange={(e) => setMapping(p => ({ ...p, file1PhoneKey: e.target.value }))}
                      className="w-full text-xs bg-slate-50 border border-slate-200 rounded px-2 py-1.5 font-mono text-slate-700 outline-none focus:border-indigo-500 focus:bg-white"
                    >
                      <option value="">-- ללא מזהה טלפון --</option>
                      {file1.columns.map(col => (
                        <option key={col} value={col}>{col}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            ) : (
              <div className="py-6 flex flex-col items-center">
                <UploadCloud className="w-12 h-12 text-slate-400 mb-3" />
                <p className="text-sm font-semibold text-slate-755 mb-1">גרור לכאן את קובץ העסקאות של TLV</p>
                <p className="text-xs text-slate-400 mb-4">תומך בפורמטים Excel (.xlsx, .xls) או דוח CSV</p>
                <label className="cursor-pointer bg-indigo-55 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-semibold px-4 py-2 rounded-lg text-xs transition-colors border border-indigo-200">
                  <span>בחר קובץ מהמחשב</span>
                  <input 
                    type="file" 
                    accept=".xlsx,.xls,.csv" 
                    className="hidden" 
                    onChange={(e) => handleFileChange(e, 1)} 
                  />
                </label>
              </div>
            )}
          </div>

          {/* File 2 Upload Area */}
          <div 
            onDragOver={handleDragOver}
            onDrop={handleDropFile2}
            className={`bg-white rounded-2xl border-2 border-dashed p-6 flex flex-col items-center justify-center transition-all ${
              file2 
                ? "border-emerald-200 bg-emerald-50/10 shadow-sm"
                : "border-slate-300 hover:border-indigo-400 bg-white"
            }`}
          >
            <div className="flex items-center justify-between w-full mb-4">
              <span className="bg-indigo-50 text-indigo-700 text-xs px-2.5 py-1 rounded-full font-bold">קובץ 2</span>
              {file2 && (
                <span className="flex items-center gap-1 text-emerald-600 text-xs font-semibold">
                  <CheckCircle className="w-4 h-4" /> קובץ נטען בהצלחה
                </span>
              )}
            </div>

            {loading2 ? (
              <div className="py-8 flex flex-col items-center gap-3">
                <RefreshCw className="w-10 h-10 text-indigo-600 animate-spin" />
                <p className="text-sm text-slate-500">קורא ומפענח את קובץ אקסל להשוואה...</p>
              </div>
            ) : file2 ? (
              <div className="w-full text-right">
                <div className="flex items-start gap-3">
                  <div className="bg-emerald-100 text-emerald-700 p-3 rounded-xl">
                    <FileSpreadsheet className="w-8 h-8" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <h3 className="font-bold text-slate-900 truncate" title={file2.fileName}>
                      {file2.fileName}
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      {file2.rows.length.toLocaleString()} מזהים ייחודיים שנטענו • {file2.columns.length} עמודות שנמצאו
                    </p>
                  </div>
                  <button 
                    onClick={() => setFile2(null)} 
                    className="p-1 text-slate-450 hover:text-rose-600 rounded hover:bg-rose-50 transition-colors"
                    title="הסר קובץ"
                  >
                    <Trash2 className="w-5 h-5" />
                  </button>
                </div>

                {/* Auto Column Mapping Viewer for File 2 */}
                <div className="mt-4 pt-4 border-t border-slate-100 grid grid-cols-1 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">עמודת השוואה (Session ID):</label>
                    <select
                      value={mapping.file2SessionIdKey}
                      onChange={(e) => setMapping(p => ({ ...p, file2SessionIdKey: e.target.value }))}
                      className="w-full text-xs bg-slate-50 border border-slate-200 rounded px-2 py-1.5 font-mono text-slate-700 outline-none focus:border-indigo-500 focus:bg-white"
                    >
                      {file2.columns.map(col => (
                        <option key={col} value={col}>{col}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            ) : (
              <div className="py-6 flex flex-col items-center">
                <UploadCloud className="w-12 h-12 text-slate-400 mb-3" />
                <p className="text-sm font-semibold text-slate-755 mb-1">גרור לכאן את קובץ ה-Session ID להשוואה</p>
                <p className="text-xs text-slate-400 mb-4">תומך בפורמטים Excel (.xlsx, .xls) או דוח CSV</p>
                <label className="cursor-pointer bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-semibold px-4 py-2 rounded-lg text-xs transition-colors border border-indigo-200">
                  <span>בחר קובץ מהמחשב</span>
                  <input 
                    type="file" 
                    accept=".xlsx,.xls,.csv" 
                    className="hidden" 
                    onChange={(e) => handleFileChange(e, 2)} 
                  />
                </label>
              </div>
            )}
          </div>
        </section>

        {/* Informational Guidelines Banner when no files are uploaded */}
        {!file1 && (
          <div className="bg-indigo-900/10 border border-indigo-500/20 text-indigo-950 p-6 rounded-2xl flex flex-col md:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-4 text-right">
              <div className="bg-indigo-600 text-white p-3 rounded-full shrink-0">
                <Sparkles className="w-6 h-6 text-amber-300" />
              </div>
              <div>
                <h4 className="font-bold text-base text-indigo-950">אין לך קבצים מוכנים להשוואה אצלך כרגע?</h4>
                <p className="text-sm text-slate-600 mt-1">
                  לחץ על הכפתור "טען קבצי דוגמה" בחלק העליון או משמאל כדי לחוות את חווית ההצלבה ולקבל הצצה לתוצאות באופן מיידי.
                </p>
              </div>
            </div>
            <button
              onClick={loadDemoData}
              className="bg-indigo-600 text-white hover:bg-indigo-500 font-semibold text-sm px-5 py-2.5 rounded-xl flex items-center gap-2 shadow transition-all shrink-0"
            >
              <Plus className="w-5 h-5 text-indigo-200" />
              <span>טען עסקאות TLV לדוגמה</span>
            </button>
          </div>
        )}

        {/* Step 2: Comparison Statistics & Alerts Dashboard */}
        {comparison && (
          <section className="flex flex-col gap-6" id="dashboard-section">
            <h2 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <Filter className="w-5 h-5 text-indigo-600" />
              <span>לוח נתוני ההצלבה והסינון</span>
            </h2>

            {/* Metrics Bento Boxes */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                <span className="text-xs font-bold text-slate-450">סה"כ עסקאות בקובץ 1</span>
                <div className="mt-2 flex items-baseline gap-2">
                  <span className="text-2xl sm:text-3xl font-extrabold text-slate-900">
                    {comparison.totalFile1.toLocaleString()}
                  </span>
                  <span className="text-xs text-slate-400">שורות</span>
                </div>
                <div className="mt-2 text-xs text-slate-400 border-t border-slate-100 pt-1">
                  נטענו מתוך קובץ העסקאות
                </div>
              </div>

              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                <span className="text-xs font-bold text-slate-450">שורות שעברו סינון מכונות TLV</span>
                <div className="mt-2 flex items-baseline gap-2">
                  <span className="text-2xl sm:text-3xl font-extrabold text-indigo-600">
                    {comparison.filteredCount.toLocaleString()}
                  </span>
                  <span className="text-xs text-slate-400">מתוך {comparison.totalFile1.toLocaleString()}</span>
                </div>
                <div className="mt-2 text-xs text-indigo-600/70 border-t border-slate-100 pt-1">
                  שורות ששם מכונה הוא TLV 1 (או 2)
                </div>
              </div>

              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                <span className="text-xs font-bold text-slate-450">סה"כ מזהים בקובץ השוואה</span>
                <div className="mt-2 flex items-baseline gap-2">
                  <span className="text-2xl sm:text-3xl font-extrabold text-slate-900">
                    {comparison.totalFile2.toLocaleString()}
                  </span>
                  <span className="text-xs text-slate-400">מזהי Session ID</span>
                </div>
                <div className="mt-2 text-xs text-slate-400 border-t border-slate-100 pt-1">
                  {file2 ? "קובץ השוואה 2 נטען" : "מחכה לטעינת קובץ 2..."}
                </div>
              </div>

              <div className={`p-5 rounded-2xl shadow-sm flex flex-col justify-between transition-colors ${
                comparison.missingItems.length > 0 
                  ? "bg-rose-50 border border-rose-200 text-rose-950" 
                  : "bg-emerald-50 border border-emerald-200 text-emerald-950"
              }`}>
                <span className="text-xs font-bold opacity-80">מזהים חסרים בקובץ 2</span>
                <div className="mt-2 flex items-baseline gap-2">
                  <span className={`text-2xl sm:text-3xl font-extrabold ${comparison.missingItems.length > 0 ? "text-rose-600" : "text-emerald-700"}`}>
                    {comparison.missingItems.length.toLocaleString()}
                  </span>
                  <span className="text-xs opacity-75">אי התאמות</span>
                </div>
                <div className="mt-2 text-xs opacity-70 border-t border-slate-200/50 pt-1 flex items-center justify-between">
                  <span>עסקאות ללא מזהה סשן תואם</span>
                  {comparison.missingItems.length > 0 && (
                    <span className="inline-block bg-rose-500 text-white rounded-full w-2 h-2 animate-ping" />
                  )}
                </div>
              </div>
            </div>

            {/* Big Status Warning Notification */}
            {comparison.missingItems.length > 0 ? (
              <div className="bg-amber-50/80 border border-amber-200 rounded-2xl p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 shadow-sm">
                <div className="flex gap-3">
                  <div className="bg-amber-100 text-amber-800 p-3 rounded-full mt-0.5 sm:mt-0 shrink-0">
                    <AlertTriangle className="w-6 h-6 animate-pulse" />
                  </div>
                  <div>
                    <h3 className="font-bold text-amber-900 text-base">התרעה: נמצאו מזהים חסרים!</h3>
                    <p className="text-sm text-amber-850 mt-1">
                      מתוך {comparison.filteredCount} עסקאות המסוננות עבור מכונות <span className="font-semibold">TLV 1 & TLV 2</span>, נמצאו <span className="font-bold">{comparison.missingItems.length}</span> מזהים שלא מופיעים (Session ID) בקובץ ההשוואה השני. עליך לבחון את פרטי העסקאות או לייצא אותן כדוח.
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-3 w-full sm:w-auto shrink-0 justify-end">
                  <label className="flex items-center gap-2 text-xs font-semibold text-slate-600 bg-white px-3 py-1.5 rounded-lg border border-slate-200 select-none">
                    <input 
                      type="checkbox" 
                      checked={playSound} 
                      onChange={(e) => setPlaySound(e.target.checked)}
                      className="rounded text-indigo-600 focus:ring-indigo-500" 
                    />
                    <span>השמע צליל התרעה</span>
                  </label>

                  <button
                    onClick={exportMissingToExcel}
                    className="flex items-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs px-4 py-2.5 rounded-xl transition shadow-sm w-full sm:w-auto justify-center"
                  >
                    <Download className="w-4 h-4" />
                    <span>ייצא רשימה לאקסל</span>
                  </button>
                </div>
              </div>
            ) : file2 ? (
              <div className="bg-emerald-50 border border-emerald-200 rounded-2xl p-5 flex items-center gap-3">
                <div className="bg-emerald-100 text-emerald-800 p-2.5 rounded-full">
                  <CheckCircle className="w-6 h-6" />
                </div>
                <div>
                  <h3 className="font-bold text-emerald-900">הכל תקין! אין מזהים חסרים</h3>
                  <p className="text-sm text-emerald-800 mt-0.5">
                    המערכת בדקה את כל עסקאות TLV 1 ו-TLV 2 הנטענות ומצאה התאמה מושלמת של 100% מול מזהי ה-Session ID בקובץ ההשוואה.
                  </p>
                </div>
              </div>
            ) : (
              <div className="bg-slate-100 border border-slate-200 rounded-2xl p-5 flex items-center gap-3">
                <div className="bg-slate-200 text-slate-700 p-2.5 rounded-full">
                  <Info className="w-5 h-5 animate-bounce" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-800">ממתין להעלאת קובץ ההשוואה (קובץ 2)</h3>
                  <p className="text-sm text-slate-500 mt-0.5">
                    נא להעלות את קובץ ה-Session ID כדי לבצע את ההצלבה ולקבל רשימת מזהים חסרים.
                  </p>
                </div>
              </div>
            )}
          </section>
        )}

        {/* Step 3: Interactive Filtered & Missing Items List Table */}
        {comparison && (
          <section className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden" id="results-table-section">
            {/* Table Search & Filter Toolbar Header */}
            <div className="bg-slate-50/80 px-6 py-4 border-b border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex items-center gap-2">
                <h3 className="font-bold text-slate-900 text-base">פירוט עסקאות מסוננות חסרות</h3>
                <span className="bg-rose-100 text-rose-700 font-bold font-mono text-xs px-2.5 py-0.5 rounded-full">
                  {finalFilteredList.length} שורות מוצגות
                </span>
              </div>

              {/* Filtering Controls */}
              <div className="flex flex-wrap items-center gap-3">
                {/* Search field */}
                <div className="relative w-full sm:w-64">
                  <Search className="w-4 h-4 text-slate-400 absolute right-3 top-2.5" />
                  <input
                    type="text"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    placeholder="חיפוש לפי מזהה, מכונה, טלפון..."
                    className="w-full text-xs bg-white border border-slate-200 rounded-lg pr-9 pl-3 py-2 outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500/20"
                  />
                  {searchTerm && (
                    <button 
                      onClick={() => setSearchTerm("")}
                      className="absolute left-2.5 top-2 text-slate-400 hover:text-slate-600 text-xs font-semibold px-1 rounded hover:bg-slate-100"
                    >
                      נקה
                    </button>
                  )}
                </div>

                {/* Machine Preset Quick Filters */}
                <div className="flex rounded-lg border border-slate-200 bg-white p-1 text-xs">
                  <button
                    onClick={() => setMachineFilter("ALL")}
                    className={`px-3 py-1 rounded-md font-semibold transition-colors ${
                      machineFilter === "ALL" 
                        ? "bg-indigo-600 text-white" 
                        : "text-slate-600 hover:bg-slate-55 hover:bg-slate-50"
                    }`}
                  >
                    הכל
                  </button>
                  <button
                    onClick={() => setMachineFilter("TLV 1")}
                    className={`px-3 py-1 rounded-md font-semibold transition-colors ${
                      machineFilter === "TLV 1" 
                        ? "bg-indigo-600 text-white" 
                        : "text-slate-600 hover:bg-slate-50"
                    }`}
                  >
                    TLV 1
                  </button>
                  <button
                    onClick={() => setMachineFilter("TLV 2")}
                    className={`px-3 py-1 rounded-md font-semibold transition-colors ${
                      machineFilter === "TLV 2" 
                        ? "bg-indigo-600 text-white" 
                        : "text-slate-600 hover:bg-slate-50"
                    }`}
                  >
                    TLV 2
                  </button>
                </div>

                {/* Discarded Filter Count Indicator tooltip */}
                {comparison.skippedCount > 0 && (
                  <span className="text-xs text-slate-450 bg-slate-100 rounded-lg px-2.5 py-1.5 flex items-center gap-1">
                    <Info className="w-3.5 h-3.5 text-slate-400" />
                    <span>התעלמנו מ-{comparison.skippedCount.toLocaleString()} שורות (לא TLV 1/2)</span>
                  </span>
                )}
              </div>
            </div>

            {/* Excel / Tabular display area */}
            <div className="overflow-x-auto w-full">
              {finalFilteredList.length > 0 ? (
                <table className="w-full text-right border-collapse">
                  <thead>
                    <tr className="bg-slate-50 text-slate-500 font-bold text-xs uppercase border-b border-slate-200">
                      <th className="py-3.5 px-6 text-center select-none w-14">#</th>
                      <th className="py-3.5 px-6 select-none font-semibold">מזהה עסקה (ID)</th>
                      <th className="py-3.5 px-6 select-none font-semibold">מכונה</th>
                      <th className="py-3.5 px-6 select-none font-semibold">מספר טלפון לפעולות קשר (054****)</th>
                      <th className="py-3.5 px-4 select-none font-semibold text-center w-28">חשוף / העתק</th>
                      <th className="py-3.5 px-6 select-none font-semibold">תאריך מקור</th>
                      <th className="py-3.5 px-6 select-none font-semibold">סכום / מטבע</th>
                      <th className="py-3.5 px-6 select-none font-semibold">סטטוס</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-sm">
                    {finalFilteredList.map((item, index) => {
                      const phoneDetails = formatPhone(item.phone);
                      const isRevealed = !!revealedPhones[item.id];

                      return (
                        <tr 
                          key={item.id} 
                          className="hover:bg-slate-50/50 transition-colors group"
                        >
                          {/* Row index */}
                          <td className="py-4 px-6 text-slate-400 text-xs font-mono text-center">
                            {index + 1}
                          </td>

                          {/* ID with instant hover copy button */}
                          <td className="py-4 px-6 font-mono font-medium text-slate-900">
                            <div className="flex items-center gap-2">
                              <span className="truncate max-w-xs block" title={item.id}>
                                {item.id}
                              </span>
                              <button
                                onClick={() => handleCopy(item.id, "id")}
                                className="p-1 rounded bg-slate-100 hover:bg-indigo-50 text-slate-500 hover:text-indigo-600 transition-colors opacity-0 group-hover:opacity-100"
                                title="העתק מזהה מלא"
                              >
                                {copiedId === item.id ? (
                                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                                ) : (
                                  <Copy className="w-3.5 h-3.5" />
                                )}
                              </button>
                            </div>
                          </td>

                          {/* Machine name pill */}
                          <td className="py-4 px-6">
                            <span className={`inline-flex items-center px-2 py-1 rounded-md text-xs font-bold font-mono ${
                              item.machineName.toUpperCase() === "TLV 1"
                                ? "bg-amber-100 text-amber-800 border border-amber-200/50"
                                : "bg-teal-100 text-teal-800 border border-teal-200/50"
                            }`}>
                              {item.machineName}
                            </span>
                          </td>

                          {/* Phone formatting per Hebrew user specs */}
                          <td className="py-4 px-6 font-mono">
                            <div className="flex items-center gap-2">
                              {/* Display output */}
                              <span className="text-slate-850 font-medium">
                                {isRevealed ? phoneDetails.formatted : phoneDetails.masked}
                              </span>

                              {/* Alert details tooltip if number starts with +972 */}
                              {String(item.phone).startsWith("+972") && (
                                <span className="text-slate-400 text-[10px] bg-slate-100 px-1 py-0.5 rounded leading-none">
                                  בוצע המרה מקידומת 972
                                </span>
                              )}
                            </div>
                          </td>

                          {/* Interactive action to reveal or copy phone number */}
                          <td className="py-4 px-4 text-center">
                            <div className="flex items-center justify-center gap-1.5">
                              {/* Toggle eye */}
                              <button
                                onClick={() => togglePhoneReveal(item.id)}
                                className="p-1.5 rounded-lg text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 transition-all"
                                title={isRevealed ? "הסתר מספר" : "הצג מספר מלא"}
                              >
                                {isRevealed ? (
                                  <EyeOff className="w-4 h-4" />
                                ) : (
                                  <Eye className="w-4 h-4" />
                                )}
                              </button>

                              {/* Copy number */}
                              <button
                                onClick={() => handleCopy(phoneDetails.formatted, "phone")}
                                className="p-1.5 rounded-lg text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 transition-all"
                                title="העתק טלפון שלם מיושר ל-0"
                              >
                                {copiedPhone === phoneDetails.formatted ? (
                                  <Check className="w-4 h-4 text-emerald-600" />
                                ) : (
                                  <Copy className="w-4 h-4" />
                                )}
                              </button>
                            </div>
                          </td>

                          {/* Original Transaction created date */}
                          <td className="py-4 px-6 text-slate-500 font-mono text-xs">
                            {item.rawRow.created || item.rawRow.createdDate || "אין נתון"}
                          </td>

                          {/* Original transaction fiat amount */}
                          <td className="py-4 px-6 font-medium text-slate-800">
                            {item.rawRow.fiat ? (
                              <span className="font-mono">
                                {Number(item.rawRow.fiat).toLocaleString()} {item.rawRow.fiatCode || "ILS"}
                              </span>
                            ) : (
                              <span className="text-slate-400">-</span>
                            )}
                          </td>

                          {/* Original transaction status */}
                          <td className="py-4 px-6">
                            <span className="text-xs text-slate-600 bg-slate-100 rounded-full px-2.5 py-0.5 inline-block">
                              {item.rawRow.status || "לא מוגדר"}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              ) : (
                <div className="py-16 text-center flex flex-col items-center justify-center">
                  <div className="bg-slate-100 p-4 rounded-full text-slate-400 mb-3">
                    <Search className="w-8 h-8" />
                  </div>
                  <h4 className="font-bold text-slate-800">לא נמצאו תוצאות העונות על החיפוש או הסינון</h4>
                  <p className="text-xs text-slate-550 mt-1 max-w-md mx-auto">
                    נסה לשנות את מונח החיפוש בתיבה או בחר להציג "הכל" במקום סינון למכונה ספציפית.
                  </p>
                </div>
              )}
            </div>

            {/* Table Footer Actions */}
            <div className="bg-slate-50 px-6 py-4 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500">
              <div className="flex items-center gap-2">
                <span>מציג {finalFilteredList.length} עסקאות חסרות.</span>
                <span>•</span>
                <span>מקור העסקאות: {file1?.fileName || "נתוני דוגמה"}</span>
              </div>
              <button
                onClick={exportMissingToExcel}
                className="flex items-center gap-1.5 text-indigo-700 hover:text-indigo-900 font-bold active:scale-95 transition-transform"
              >
                <Download className="w-4 h-4" />
                <span>שמור דוח מזהים חסרים כקובץ Excel (.xlsx)</span>
              </button>
            </div>
          </section>
        )}
      </main>

      {/* Modern Professional Footer */}
      <footer className="bg-white border-t border-slate-200 text-slate-500 text-xs py-6 mt-12">
        <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-slate-800">השוואת עסקאות TLV</span>
            <span>|</span>
            <span>נוצר לטובת ניהול ובקרת הפרשים בקובצי אקסל</span>
          </div>
          <div className="flex flex-col sm:items-end gap-1">
            <span>כל העיבוד מתבצע בצורה מאובטחת באופן מקומי בלבד בדפדפן</span>
            <span className="text-slate-650 font-medium text-[11px]">
              © כל הזכויות שמורות לרוברט טייגר 2026 @ Bitcoin Change TLV
            </span>
          </div>
        </div>
      </footer>
    </div>
  );
}
