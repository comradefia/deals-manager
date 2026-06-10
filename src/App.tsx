/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import { 
  Plus, 
  Trash2, 
  Eye, 
  EyeOff, 
  TrendingUp, 
  Coins, 
  ArrowDownLeft, 
  ArrowUpRight, 
  RotateCcw,
  Sparkles,
  HelpCircle,
  TrendingDown,
  Lock,
  Unlock,
  PackageOpen,
  DollarSign,
  Pencil,
  X,
  Check,
  UploadCloud,
  FileCheck,
  AlertCircle,
  Info,
  ListPlus
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { InboundRow, OutboundRow } from "./types";

// Unique ID helper
const generateId = () => Math.random().toString(36).substr(2, 9);

// Initial sample telecom swap rows to provide immediate utility
const INITIAL_INBOUND: InboundRow[] = [
  { id: "in-1", destination: "United Kingdom Mobiles O2 (447)", isMasked: true, plannedMinutes: 250000, rpm: 0.032 },
  { id: "in-2", destination: "Saudi Arabia STC Mobiles (966)", isMasked: true, plannedMinutes: 120000, rpm: 0.075 },
  { id: "in-3", destination: "Germany T-Mobile Trunks (491)", isMasked: true, plannedMinutes: 180000, rpm: 0.024 }
];

const INITIAL_OUTBOUND: OutboundRow[] = [
  { id: "out-1", destination: "United Kingdom Mobiles Vodafone (447)", isMasked: true, plannedMinutes: 200000, cpm: 0.029 },
  { id: "out-2", destination: "Saudi Arabia Mobily Trunks (966)", isMasked: true, plannedMinutes: 150000, cpm: 0.068 },
  { id: "out-3", destination: "Germany Vodafone Swaps (491)", isMasked: true, plannedMinutes: 140000, cpm: 0.021 }
];

// Aesthetic masking helper function
// Replaces mid-parts of words to shield actual destinations but preserve spatial shape
function maskDestination(val: string): string {
  if (!val) return "";
  return val.split(" ").map(word => {
    if (word.length <= 1) return "•";
    if (word.length === 2) return word[0] + "•";
    if (word.length === 3) return word[0] + "•" + word[2];
    const middleLength = word.length - 2;
    return word[0] + "•".repeat(middleLength) + word[word.length - 1];
  }).join(" ");
}

// Custom manual OR automated fallback selector helper
function getDisplayMaskedName(row: { destination: string; maskedName?: string }): string {
  if (row.maskedName && row.maskedName.trim() !== "") {
    return row.maskedName.trim();
  }
  return maskDestination(row.destination);
}

export default function App() {
  // State for rows
  const [inboundRows, setInboundRows] = useState<InboundRow[]>(INITIAL_INBOUND);
  const [outboundRows, setOutboundRows] = useState<OutboundRow[]>(INITIAL_OUTBOUND);

  // Modal Editing States
  const [editingInbound, setEditingInbound] = useState<InboundRow | null>(null);
  const [editingOutbound, setEditingOutbound] = useState<OutboundRow | null>(null);

  // Broad state configurations
  const [allInboundMasked, setAllInboundMasked] = useState<boolean>(true);
  const [allOutboundMasked, setAllOutboundMasked] = useState<boolean>(true);
  const [showInstructions, setShowInstructions] = useState<boolean>(false);

  // Excel File Upload states
  const [fileName, setFileName] = useState<string>("");
  const [parsedDestinations, setParsedDestinations] = useState<string[]>([]);
  const [parsingError, setParsingError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [importTarget, setImportTarget] = useState<"inbound" | "outbound">("inbound");
  const [defaultMinutes, setDefaultMinutes] = useState<number>(100000);
  const [defaultRate, setDefaultRate] = useState<number>(0.030);
  const [isImporting, setIsImporting] = useState<boolean>(false);
  const [importSuccessMsg, setImportSuccessMsg] = useState<string | null>(null);

  // NEW custom/manual masking states for imported spreadsheet values
  const [unmaskedIndices, setUnmaskedIndices] = useState<number[]>([]);
  const [customMaskedNames, setCustomMaskedNames] = useState<Record<number, string>>({});

  // Handle file select/drop
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFile(file);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      handleFile(file);
    }
  };

  const handleFile = (file: File) => {
    setParsingError(null);
    setImportSuccessMsg(null);
    setParsedDestinations([]);
    setFileName(file.name);
    setUnmaskedIndices([]);
    setCustomMaskedNames({});

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) throw new Error("Could not read file data.");
        
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) throw new Error("The Excel file doesn't contain any sheets.");
        
        const worksheet = workbook.Sheets[firstSheetName];
        const json: any[] = XLSX.utils.sheet_to_json(worksheet);
        
        if (json.length === 0) {
          throw new Error("The selected Excel sheet has no records.");
        }

        // Find standard "original destination" header (trim whitespace, lowercase, match spaces)
        const firstRow = json[0];
        const targetColumnKey = Object.keys(firstRow).find(
          key => key.toLowerCase().trim().replace(/[\s\-_]+/g, ' ') === 'original destination'
        );

        if (!targetColumnKey) {
          throw new Error("Missing required column: The Excel sheet MUST contain a column named exactly 'original destination'.");
        }

        const destinationsList: string[] = [];
        json.forEach((row: any) => {
          if (row[targetColumnKey] !== undefined && row[targetColumnKey] !== null) {
            const val = String(row[targetColumnKey]).trim();
            if (val) {
              destinationsList.push(val);
            }
          }
        });

        if (destinationsList.length === 0) {
          throw new Error("The 'original destination' column was found, but all records in it are blank.");
        }

        setParsedDestinations(destinationsList);
      } catch (err: any) {
        setParsingError(String(err.message || "Could not decode Excel file sheets. Ensure it is a valid Excel spreadsheet file (.xlsx, .xls, .xlsb, .xlsm, or .csv)."));
        setParsedDestinations([]);
      }
    };

    reader.onerror = () => {
      setParsingError("Error accessing file stream.");
    };

    reader.readAsArrayBuffer(file);
  };

  // Perform migration of destinations into the chosen table
  const handleCommitImport = () => {
    if (parsedDestinations.length === 0) return;
    setIsImporting(true);

    setTimeout(() => {
      if (importTarget === "inbound") {
        const newRows: InboundRow[] = parsedDestinations.map((dest, idx) => {
          const isUnmaskedSelected = unmaskedIndices.includes(idx);
          return {
            id: generateId(),
            destination: dest,
            isMasked: !isUnmaskedSelected,
            plannedMinutes: defaultMinutes,
            rpm: defaultRate,
            maskedName: customMaskedNames[idx] || ""
          };
        });
        setInboundRows(prev => [...prev, ...newRows]);
      } else {
        const newRows: OutboundRow[] = parsedDestinations.map((dest, idx) => {
          const isUnmaskedSelected = unmaskedIndices.includes(idx);
          return {
            id: generateId(),
            destination: dest,
            isMasked: !isUnmaskedSelected,
            plannedMinutes: defaultMinutes,
            cpm: defaultRate,
            maskedName: customMaskedNames[idx] || ""
          };
        });
        setOutboundRows(prev => [...prev, ...newRows]);
      }

      setImportSuccessMsg(`Successfully imported ${parsedDestinations.length} telecom trunks with customized planned parameters (${defaultMinutes.toLocaleString()} mins, $${defaultRate.toFixed(4)} rate)! Selected unmasked represents: ${unmaskedIndices.length} routes.`);
      // Reset state
      setParsedDestinations([]);
      setFileName("");
      setUnmaskedIndices([]);
      setCustomMaskedNames({});
      setIsImporting(false);
    }, 450);
  };

  const handleDownloadTemplate = () => {
    const sampleRows = [
      { "original destination": "United Kingdom Mobiles O2 (447)" },
      { "original destination": "Saudi Arabia STC Mobiles (966)" },
      { "original destination": "Germany T-Mobile Trunks (491)" },
      { "original destination": "United States Mobiles AT&T (12)" },
      { "original destination": "France Orange Mobiles (33)" }
    ];
    const ws = XLSX.utils.json_to_sheet(sampleRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Telecom Swap Sheet");
    XLSX.writeFile(wb, "telecom_swap_import_template.xlsx");
  };

  const handleClearExcelUpload = () => {
    setFileName("");
    setParsedDestinations([]);
    setParsingError(null);
    setImportSuccessMsg(null);
    setUnmaskedIndices([]);
    setCustomMaskedNames({});
  };

  // Auto-align single-row masks if global toggle shifts
  const toggleAllInboundMask = () => {
    const nextState = !allInboundMasked;
    setAllInboundMasked(nextState);
    setInboundRows(prev => prev.map(r => ({ ...r, isMasked: nextState })));
  };

  const toggleAllOutboundMask = () => {
    const nextState = !allOutboundMasked;
    setAllOutboundMasked(nextState);
    setOutboundRows(prev => prev.map(r => ({ ...r, isMasked: nextState })));
  };

  // Add empty or prefilled dynamic rows
  const addInboundRow = () => {
    const newRow: InboundRow = {
      id: generateId(),
      destination: `Voice Inbound Trunk (Prefix ${inboundRows.length + 1})`,
      isMasked: allInboundMasked,
      plannedMinutes: 100000,
      rpm: 0.035
    };
    setInboundRows(prev => [...prev, newRow]);
    // Immediately open edit modal for user convenience
    setEditingInbound(newRow);
  };

  const addOutboundRow = () => {
    const newRow: OutboundRow = {
      id: generateId(),
      destination: `Voice Outbound Trunk (Prefix ${outboundRows.length + 1})`,
      isMasked: allOutboundMasked,
      plannedMinutes: 100000,
      cpm: 0.025
    };
    setOutboundRows(prev => [...prev, newRow]);
    // Immediately open edit modal for user convenience
    setEditingOutbound(newRow);
  };

  // Delete row targets
  const deleteInboundRow = (id: string) => {
    setInboundRows(prev => prev.filter(r => r.id !== id));
    if (editingInbound?.id === id) setEditingInbound(null);
  };

  const deleteOutboundRow = (id: string) => {
    setOutboundRows(prev => prev.filter(r => r.id !== id));
    if (editingOutbound?.id === id) setEditingOutbound(null);
  };

  // Inline inputs update handler
  const updateInboundValue = (id: string, field: keyof InboundRow, value: any) => {
    setInboundRows(prev => prev.map(row => {
      if (row.id === id) {
        return { ...row, [field]: value };
      }
      return row;
    }));
  };

  const updateOutboundValue = (id: string, field: keyof OutboundRow, value: any) => {
    setOutboundRows(prev => prev.map(row => {
      if (row.id === id) {
        return { ...row, [field]: value };
      }
      return row;
    }));
  };

  // Edit Modal form submits
  const handleSaveInboundEdit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingInbound) return;
    setInboundRows(prev => prev.map(row => 
      row.id === editingInbound.id ? editingInbound : row
    ));
    setEditingInbound(null);
  };

  const handleSaveOutboundEdit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingOutbound) return;
    setOutboundRows(prev => prev.map(row => 
      row.id === editingOutbound.id ? editingOutbound : row
    ));
    setEditingOutbound(null);
  };

  // Restore defaults
  const handleResetData = () => {
    if (confirm("Are you sure you want to reset both tables to sample configurations?")) {
      setInboundRows(INITIAL_INBOUND);
      setOutboundRows(INITIAL_OUTBOUND);
      setAllInboundMasked(true);
      setAllOutboundMasked(true);
      setEditingInbound(null);
      setEditingOutbound(null);
    }
  };

  // Clear entirely
  const handleClearData = () => {
    if (confirm("Are you sure you want to delete all rows?")) {
      setInboundRows([]);
      setOutboundRows([]);
      setEditingInbound(null);
      setEditingOutbound(null);
    }
  };

  // Live summations and key stats calculations
  const totals = useMemo(() => {
    // Inbound calculations
    let totalInboundMinutes = 0;
    let totalPlannedRevenue = 0;

    inboundRows.forEach(row => {
      const mins = Math.max(0, Number(row.plannedMinutes) || 0);
      const rpm = Math.max(0, Number(row.rpm) || 0);
      totalInboundMinutes += mins;
      // Equation: Planned Revenue = Planned Minutes * RPM
      totalPlannedRevenue += (mins * rpm);
    });

    // Outbound calculations
    let totalOutboundMinutes = 0;
    let totalPlannedCost = 0;

    outboundRows.forEach(row => {
      const mins = Math.max(0, Number(row.plannedMinutes) || 0);
      const cpm = Math.max(0, Number(row.cpm) || 0);
      totalOutboundMinutes += mins;
      // Equation: Planned Cost = Planned Minutes * CPM
      totalPlannedCost += (mins * cpm);
    });

    const netMargin = totalPlannedRevenue - totalPlannedCost;
    const marginPercentage = totalPlannedRevenue > 0 
      ? (netMargin / totalPlannedRevenue) * 100 
      : 0;

    return {
      totalInboundMinutes,
      totalOutboundMinutes,
      totalPlannedRevenue,
      totalPlannedCost,
      netMargin,
      marginPercentage
    };
  }, [inboundRows, outboundRows]);

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 transition-colors duration-200">
      
      {/* Decorative Top Accent */}
      <div className="h-1.5 bg-gradient-to-r from-indigo-500 via-teal-500 to-emerald-600 w-full" />

      {/* Main Container */}
      <div className="max-w-7xl mx-auto px-4 py-8 sm:px-6 lg:px-8">
        
        {/* Header Section */}
        <header className="mb-8 flex flex-col md:flex-row md:items-center md:justify-between border-b border-slate-200 pb-6 gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1">
              <span className="p-1.5 bg-indigo-50 text-indigo-600 rounded-lg inline-flex">
                <Sparkles size={20} />
              </span>
              <span className="text-xs font-mono tracking-widest uppercase bg-indigo-50 px-2.5 py-1 text-indigo-750 font-semibold rounded">
                Carrier wholesale & SMS/Voice Swap
              </span>
            </div>
            <h1 className="text-3xl font-extrabold tracking-tight text-slate-900 sm:text-4xl" id="app-title">
              Telecom Bilateral Swap Deal Tracker
            </h1>
            <p className="mt-2 text-sm text-slate-500 max-w-2xl">
              Model, track, and balance telecom wholesale traffic swaps. Manage Inbound (terminating rate revenue) and Outbound (origination transit costs) easily.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2 self-start md:self-center">
            <button
              id="btn-help"
              onClick={() => setShowInstructions(!showInstructions)}
              className="px-3.5 py-2 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 active:bg-slate-100 rounded-lg border border-slate-200 transition-colors shadow-xs inline-flex items-center gap-1.5 cursor-pointer"
            >
              <HelpCircle size={15} />
              {showInstructions ? "Hide Reference" : "Formula Reference"}
            </button>
            <button
              id="btn-reset"
              onClick={handleResetData}
              className="px-3.5 py-2 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 active:bg-slate-100 rounded-lg border border-slate-200 transition-colors shadow-xs inline-flex items-center gap-1.5 cursor-pointer"
            >
              <RotateCcw size={15} />
              Reset Seeding
            </button>
            <button
              id="btn-clear"
              onClick={handleClearData}
              className="px-3.5 py-2 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 rounded-lg border border-rose-100 transition-colors inline-flex items-center gap-1.5 cursor-pointer"
            >
              <Trash2 size={15} />
              Clear Dataset
            </button>
          </div>
        </header>

        {/* Dynamic Formula Panel */}
        <AnimatePresence>
          {showInstructions && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: "auto", opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              className="overflow-hidden mb-8"
              id="doc-panel"
            >
              <div className="bg-indigo-950 text-indigo-100 p-6 rounded-xl border border-indigo-800 shadow-sm">
                <h3 className="text-md font-bold text-white mb-2 inline-flex items-center gap-2">
                  <Sparkles size={16} className="text-indigo-400" />
                  Formula & Automatic Calculations Reference
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 text-sm mt-3">
                  <div className="bg-indigo-900/50 p-4 rounded-lg border border-indigo-850">
                    <span className="block font-mono text-cyan-300 font-semibold mb-1">INBOUND SWAP REVENUE</span>
                    <p className="text-indigo-200">
                      Planned Revenue = Planned Minutes × RPM
                    </p>
                    <p className="text-xs text-indigo-300/80 mt-1">
                      (RPM = Revenue Per Minute. Incoming carrier voice termination terminating on your network)
                    </p>
                  </div>
                  <div className="bg-indigo-900/50 p-4 rounded-lg border border-indigo-850">
                    <span className="block font-mono text-cyan-300 font-semibold mb-1">OUTBOUND SWAP COST</span>
                    <p className="text-indigo-200">
                      Planned Cost = Planned Minutes × CPM
                    </p>
                    <p className="text-xs text-indigo-300/80 mt-1">
                      (CPM = Cost Per Minute. Cost of sending voice traffic to the partner's network)
                    </p>
                  </div>
                  <div className="bg-indigo-900/50 p-4 rounded-lg border border-indigo-850 md:col-span-2 lg:col-span-1">
                    <span className="block font-mono text-cyan-300 font-semibold mb-1">NET SETTLEMENT MARGIN</span>
                    <p className="text-indigo-200">
                      Margin = Total Planned Revenue − Total Planned Cost
                    </p>
                    <p className="text-xs text-indigo-300/80 mt-1">
                      Shows the net payout or receivable amount at the end of the bilateral cycle.
                    </p>
                  </div>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Global Financial Dashboard Overview */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-8" id="kpi-panel">
          {/* Card 1: Total Inbound Revenue */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Total Inbound Rev</span>
              <span className="p-1.5 bg-emerald-50 rounded-lg text-emerald-600">
                <TrendingUp size={18} />
              </span>
            </div>
            <div className="mt-3">
              <span className="text-2xl font-bold font-mono text-slate-900">
                ${totals.totalPlannedRevenue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </span>
              <span className="block text-xs text-slate-400 mt-1 font-mono">
                Across {totals.totalInboundMinutes} planned minutes
              </span>
            </div>
          </div>

          {/* Card 2: Total Outbound Cost */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Total Outbound Cost</span>
              <span className="p-1.5 bg-amber-50 rounded-lg text-amber-600">
                <TrendingDown size={18} />
              </span>
            </div>
            <div className="mt-3">
              <span className="text-2xl font-bold font-mono text-slate-900">
                ${totals.totalPlannedCost.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </span>
              <span className="block text-xs text-slate-400 mt-1 font-mono">
                Across {totals.totalOutboundMinutes} planned minutes
              </span>
            </div>
          </div>

          {/* Card 3: Net Deal Margin */}
          <div className={`bg-white rounded-xl border p-5 shadow-xs relative overflow-hidden ${totals.netMargin >= 0 ? 'border-emerald-200' : 'border-rose-200'}`}>
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Estimated Corporate Margin</span>
              <span className={`p-1.5 rounded-lg ${totals.netMargin >= 0 ? 'bg-emerald-50 text-emerald-600' : 'bg-rose-50 text-rose-600'}`}>
                <Coins size={18} />
              </span>
            </div>
            <div className="mt-3">
              <span className={`text-2xl font-bold font-mono ${totals.netMargin >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                {totals.netMargin < 0 ? "-" : ""}${Math.abs(totals.netMargin).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </span>
              <span className="block text-xs text-slate-400 mt-1 font-mono">
                Segment balance yield
              </span>
            </div>
          </div>

          {/* Card 4: Operating Yield % */}
          <div className={`bg-white rounded-xl border p-5 shadow-xs relative overflow-hidden ${totals.marginPercentage >= 0 ? 'border-indigo-200' : 'border-rose-200'}`}>
            <div className="flex items-center justify-between">
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Gross Operating Yield</span>
              <span className={`p-1.5 rounded-lg ${totals.marginPercentage >= 0 ? 'bg-indigo-50 text-indigo-650' : 'bg-rose-50 text-rose-600'}`}>
                <DollarSign size={18} />
              </span>
            </div>
            <div className="mt-3">
              <span className={`text-2xl font-bold font-mono ${totals.marginPercentage >= 0 ? 'text-indigo-700' : 'text-rose-700'}`}>
                {totals.marginPercentage.toFixed(1)}%
              </span>
              <div className="w-full bg-slate-100 rounded-full h-1.5 mt-2">
                <div 
                  className={`h-1.5 rounded-full ${totals.marginPercentage >= 0 ? 'bg-indigo-500' : 'bg-rose-500'}`} 
                  style={{ width: `${Math.min(Math.max(0, totals.marginPercentage), 100)}%` }} 
                />
              </div>
            </div>
          </div>
        </section>

        {/* XLS/spreadsheet Import Section */}
        <section className="mb-8" id="excel-import-section">
          <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
            <div className="p-5 border-b border-slate-100 bg-slate-50/50 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
              <div className="flex items-center gap-2.5">
                <span className="p-2 bg-indigo-50 text-indigo-600 rounded-xl">
                  <ListPlus size={20} />
                </span>
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Excel Spreadsheets Ingestion</h2>
                  <p className="text-xs text-slate-500">Bulk insert destinations from your spreadsheet. Supports all standard Excel formats (.xlsx, .xls, .xlsb, .xlsm, .csv). Columns must map to <span className="font-mono bg-slate-105 px-1 py-0.5 rounded text-indigo-700 font-semibold">'original destination'</span>.</p>
                </div>
              </div>
              <button
                type="button"
                onClick={handleDownloadTemplate}
                className="px-3.5 py-1.5 text-xs font-semibold text-indigo-600 bg-indigo-50 hover:bg-indigo-100 rounded-lg transition-colors inline-flex items-center gap-1.5 cursor-pointer self-start sm:self-center font-mono"
              >
                <RotateCcw size={13} className="rotate-180" />
                Download Template Excel
              </button>
            </div>

            <div className="p-6">
              {/* Drag/Drop Zone */}
              <div
                className={`border-2 border-dashed rounded-xl p-8 text-center transition-all ${
                  isDragging 
                    ? "border-indigo-500 bg-indigo-50/40" 
                    : fileName 
                      ? "border-emerald-300 bg-emerald-50/10"
                      : "border-slate-200 hover:border-slate-300 bg-slate-50/30 cursor-pointer"
                }`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => {
                  if (!fileName) {
                    document.getElementById("excel-input-element")?.click();
                  }
                }}
              >
                <input
                  type="file"
                  id="excel-input-element"
                  accept=".xlsx,.xls,.xlsm,.xlsb,.csv,.ods,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.ms-excel.sheet.macroEnabled.12,application/vnd.ms-excel.sheet.binary.macroEnabled.12,text/csv,application/vnd.oasis.opendocument.spreadsheet"
                  className="hidden"
                  onChange={handleFileChange}
                />

                <div className="flex flex-col items-center justify-center">
                  {fileName ? (
                    <>
                      <div className="p-3 bg-emerald-50 text-emerald-600 rounded-full mb-3 shrink-0 animate-bounce">
                        <FileCheck size={28} />
                      </div>
                      <h4 className="font-semibold text-slate-800 text-sm mb-1">
                        Loaded Sheet: <span className="font-mono text-emerald-700">{fileName}</span>
                      </h4>
                      <p className="text-xs text-slate-400 mb-2">
                        Successfully mapped {parsedDestinations.length} route entries!
                      </p>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleClearExcelUpload();
                        }}
                        className="text-xs text-rose-600 hover:text-rose-800 font-medium underline cursor-pointer"
                      >
                        Reset loaded file
                      </button>
                    </>
                  ) : (
                    <>
                      <div className="p-3 bg-slate-100 text-slate-500 rounded-full mb-3 shrink-0">
                        <UploadCloud size={28} />
                      </div>
                      <h4 className="font-semibold text-slate-700 text-sm mb-1">
                        Drag & Drop spreadsheet or <span className="text-indigo-600 underline font-medium">Browse files</span>
                      </h4>
                      <p className="text-xs text-slate-400 max-w-sm">
                        Select any Excel spreadsheet format (.xlsx, .xls, .xlsb, .xlsm, or .csv) containing a column labeled <strong className="font-mono bg-slate-105 px-1 py-0.5 rounded text-indigo-700 font-semibold uppercase">original destination</strong> to load.
                      </p>
                    </>
                  )}
                </div>
              </div>

              {/* Parsing Diagnostics Banner */}
              {parsingError && (
                <div className="mt-4 p-4 bg-rose-50 border border-rose-100 text-rose-705 rounded-xl flex items-start gap-2 text-sm">
                  <AlertCircle size={16} className="mt-0.5 shrink-0 text-rose-600" />
                  <div>
                    <h5 className="font-bold">Excel Parsing Failed</h5>
                    <p className="text-xs text-rose-600 mt-0.5">{parsingError}</p>
                  </div>
                </div>
              )}

              {/* Successful Integration Banner */}
              {importSuccessMsg && (
                <div className="mt-4 p-4 bg-emerald-50 border border-emerald-100 text-emerald-800 rounded-xl flex items-start gap-2 text-sm">
                  <FileCheck size={16} className="mt-0.5 text-emerald-600 shrink-0" />
                  <div>
                    <h5 className="font-bold">Telecom Swap Database Updated</h5>
                    <p className="text-xs text-emerald-700 mt-0.5">{importSuccessMsg}</p>
                  </div>
                </div>
              )}

              {/* Options panel for parsed destinations */}
              <AnimatePresence>
                {parsedDestinations.length > 0 && (
                  <motion.div
                    initial={{ height: 0, opacity: 0 }}
                    animate={{ height: "auto", opacity: 1 }}
                    exit={{ height: 0, opacity: 0 }}
                    className="overflow-hidden mt-6"
                  >
                    <div className="bg-slate-50 border border-slate-150 rounded-xl p-5 space-y-4">
                      <div className="flex items-center justify-between border-b border-slate-200 pb-3">
                        <span className="text-xs font-bold text-slate-700 uppercase tracking-widest flex items-center gap-1.5 font-mono">
                          <Info size={14} className="text-indigo-500" />
                          Import Configuration & settings
                        </span>
                        <span className="text-xs font-mono bg-indigo-55 text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded font-semibold">
                          {parsedDestinations.length} Discovered Trunks
                        </span>
                      </div>

                      {/* Fields configuration */}
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        {/* Target Table Dropdown Selector */}
                        <div>
                          <label className="block text-xs font-semibold text-slate-500 mb-1">
                            Target Deal Stream Table
                          </label>
                          <select
                            value={importTarget}
                            onChange={(e) => setImportTarget(e.target.value as "inbound" | "outbound")}
                            className="w-full bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-1.5 text-xs transition-all focus:outline-hidden"
                          >
                            <option value="inbound">Inbound Routes (Revenue Swap Stream)</option>
                            <option value="outbound">Outbound Routes (Cost Swap Stream)</option>
                          </select>
                        </div>

                        {/* Fallback Minutes Configuration */}
                        <div>
                          <label className="block text-xs font-semibold text-slate-500 mb-1">
                            Default Planned Minutes per Row
                          </label>
                          <input
                            type="number"
                            min="1"
                            value={defaultMinutes}
                            onChange={(e) => setDefaultMinutes(Math.max(1, parseInt(e.target.value) || 0))}
                            className="font-mono w-full bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-1.5 text-xs transition-all focus:outline-hidden"
                          />
                        </div>

                        {/* Fallback Rate Config (RPM/CPM) */}
                        <div>
                          <label className="block text-xs font-semibold text-slate-500 mb-1">
                            Default RPM/CPM Rate ($)
                          </label>
                          <input
                            type="number"
                            min="0"
                            step="0.01"
                            value={defaultRate}
                            onChange={(e) => setDefaultRate(Math.max(0, parseFloat(e.target.value) || 0))}
                            className="font-mono w-full bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-1.5 text-xs transition-all focus:outline-hidden"
                          />
                        </div>
                      </div>

                      {/* Render destinations preview block with custom selection & masking controllers */}
                      <div className="space-y-2 border border-slate-200/60 bg-white/50 p-4 rounded-xl">
                        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 pb-2 border-b border-slate-150">
                          <div>
                            <span className="block text-xs font-bold text-slate-705 text-slate-700 uppercase tracking-widest font-mono">
                              Configure Individual Route Masking Status
                            </span>
                            <span className="block text-[11px] text-slate-400 mt-0.5">
                              Check an item to include it in the <strong className="text-emerald-600">Unmasked List</strong>. Leave unchecked to keep <strong className="text-amber-600">Masked</strong>.
                            </span>
                          </div>
                          
                          {/* Bulk Actions */}
                          <div className="flex gap-1.5 shrink-0">
                            <button
                              type="button"
                              onClick={() => setUnmaskedIndices(parsedDestinations.map((_, i) => i))}
                              className="px-2.5 py-1 text-[10px] font-bold text-slate-600 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg transition-all cursor-pointer shadow-2xs"
                            >
                              Select All Unmasked
                            </button>
                            <button
                              type="button"
                              onClick={() => setUnmaskedIndices([])}
                              className="px-2.5 py-1 text-[10px] font-bold text-slate-600 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg transition-all cursor-pointer shadow-2xs"
                            >
                              Deselect (All Masked)
                            </button>
                          </div>
                        </div>

                        <div className="max-h-56 overflow-y-auto divide-y divide-slate-100 pr-1.5 scrollbar-thin">
                          {parsedDestinations.map((dest, i) => {
                            const isUnmasked = unmaskedIndices.includes(i);
                            return (
                              <div key={i} className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 py-2 text-xs">
                                {/* Left side: Index, Checkbox and raw name */}
                                <div className="flex items-center gap-2.5 min-w-0 flex-1">
                                  <span className="font-mono text-[10px] text-slate-400 w-11 shrink-0">Row {i + 1}</span>
                                  
                                  <label className="flex items-center gap-2 cursor-pointer select-none min-w-0">
                                    <input
                                      type="checkbox"
                                      className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 h-3.5 w-3.5 cursor-pointer focus:outline-hidden"
                                      checked={isUnmasked}
                                      onChange={(e) => {
                                        if (e.target.checked) {
                                          setUnmaskedIndices(prev => [...prev, i]);
                                        } else {
                                          setUnmaskedIndices(prev => prev.filter(idx => idx !== i));
                                        }
                                      }}
                                    />
                                    <span className={`font-semibold truncate text-slate-750 transition-colors ${isUnmasked ? 'text-slate-900 font-bold' : 'text-slate-400 font-medium'}`} title={dest}>
                                      {dest}
                                    </span>
                                  </label>
                                </div>

                                {/* Right side: Status and Manual Mask Input */}
                                <div className="flex items-center gap-2 shrink-0 sm:w-72 justify-end">
                                  {isUnmasked ? (
                                    <span className="px-2 py-0.5 text-[9px] font-bold font-mono tracking-wider uppercase rounded bg-emerald-50 text-emerald-700 border border-emerald-150 shrink-0">
                                      Unmasked
                                    </span>
                                  ) : (
                                    <div className="flex items-center gap-1.5 w-full">
                                      <span className="px-2 py-0.5 text-[9px] font-bold font-mono tracking-wider uppercase rounded bg-amber-50 text-amber-700 border border-amber-150 shrink-0">
                                        Masked
                                      </span>
                                      <input
                                        type="text"
                                        value={customMaskedNames[i] || ""}
                                        onChange={(e) => setCustomMaskedNames(prev => ({ ...prev, [i]: e.target.value }))}
                                        placeholder={maskDestination(dest) || "Manual masked label..."}
                                        className="w-full bg-slate-50 hover:bg-white focus:bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded px-2.5 py-0.5 text-[11px] font-mono transition-all text-slate-800 placeholder:text-slate-400 focus:outline-hidden"
                                        title="Manually key in masked name"
                                      />
                                    </div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>

                      {/* Commit Excel Data to React state */}
                      <div className="flex justify-end gap-2 pt-2">
                        <button
                          type="button"
                          onClick={handleClearExcelUpload}
                          className="px-4 py-1.5 text-xs font-semibold text-slate-600 bg-white border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors cursor-pointer"
                        >
                          Cancel
                        </button>
                        <button
                          type="button"
                          onClick={handleCommitImport}
                          disabled={isImporting}
                          className="px-5 py-1.5 text-xs font-bold text-white bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 rounded-lg transition-colors shadow-xs flex items-center gap-1.5 cursor-pointer"
                        >
                          {isImporting ? (
                            <>
                              <span className="animate-spin rounded-full h-3.5 w-3.5 border-b-2 border-white" />
                              Processing Ingestion...
                            </>
                          ) : (
                            <>
                              <Plus size={14} />
                              Confirm Import & Merge
                            </>
                          )}
                        </button>
                      </div>

                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </div>
        </section>

        {/* Dynamic Tables Area */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8" id="tables-container">
          
          {/* ========================================================= */}
          {/* TABLE ONE: INBOUND                                        */}
          {/* ========================================================= */}
          <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs flex flex-col" id="inbound-card">
            
            {/* Table Header Section */}
            <div className="p-5 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="p-2 bg-emerald-50 text-emerald-600 rounded-xl">
                  <ArrowDownLeft size={20} />
                </span>
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Inbound Voice Swaps</h2>
                  <p className="text-xs text-slate-500">Traffic terminated on our network (Revenue stream)</p>
                </div>
              </div>
              
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={toggleAllInboundMask}
                  title="Toggle destination privacy for Inbound"
                  className="px-2.5 py-1.5 text-xs font-medium border border-slate-200 bg-white rounded-lg hover:bg-slate-50 text-slate-600 transition-colors inline-flex items-center gap-1 cursor-pointer"
                >
                  {allInboundMasked ? (
                    <>
                      <Eye size={13} className="text-sky-600" />
                      <span className="hidden sm:inline">Show Original</span>
                    </>
                  ) : (
                    <>
                      <EyeOff size={13} className="text-slate-400" />
                      <span className="hidden sm:inline">Hide Original</span>
                    </>
                  )}
                </button>

                <button
                  type="button"
                  onClick={addInboundRow}
                  className="px-3 py-1.5 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition-colors inline-flex items-center gap-1 shadow-xs cursor-pointer"
                >
                  <Plus size={14} />
                  Add Row
                </button>
              </div>
            </div>

            {/* Table Core Container */}
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-150 text-xs text-slate-500 uppercase bg-slate-50/30">
                    <th className="py-3 px-3 font-semibold w-10 text-center">#</th>
                    <th className="py-3 px-3 font-semibold min-w-[200px]">Masked Destination</th>
                    <th className="py-3 px-2 font-semibold w-24 text-right">Planned Min</th>
                    <th className="py-3 px-2 font-semibold w-24 text-right">RPM ($)</th>
                    <th className="py-3 px-3 text-right font-semibold w-28">Planned Revenue</th>
                    <th className="py-3 px-3 font-semibold w-20 text-center">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-150">
                  <AnimatePresence initial={false}>
                    {inboundRows.length === 0 ? (
                      <motion.tr 
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        className="text-center"
                      >
                        <td colSpan={6} className="py-12 px-4 text-sm text-slate-400">
                          <PackageOpen className="mx-auto text-slate-300 mb-2" size={32} />
                          No Inbound routes added yet.
                        </td>
                      </motion.tr>
                    ) : (
                      inboundRows.map((row, idx) => {
                        const plannedRevenue = row.plannedMinutes * row.rpm;
                        return (
                          <motion.tr
                            key={row.id}
                            initial={{ opacity: 0, y: 15 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, x: -15, transition: { duration: 0.15 } }}
                            transition={{ duration: 0.15 }}
                            className="hover:bg-slate-50/80 text-sm align-middle group"
                          >
                            {/* Counter Index */}
                            <td className="py-2 px-3 text-center text-xs font-mono text-slate-400">
                              {idx + 1}
                            </td>

                            {/* Destination Input (Clearly demarcated form fields) */}
                            <td className="py-2 px-2">
                              <div className="flex flex-col gap-1">
                                <div className="flex items-center gap-1">
                                  <input
                                    type="text"
                                    value={row.destination}
                                    onChange={(e) => updateInboundValue(row.id, "destination", e.target.value)}
                                    placeholder="Destination location..."
                                    className="w-full bg-slate-50 hover:bg-white focus:bg-white text-slate-800 focus:ring-1 focus:ring-emerald-500 rounded border border-slate-200 hover:border-slate-300 px-2.5 py-1 text-xs transition-all focus:border-emerald-500 font-medium placeholder:text-slate-300 focus:outline-hidden"
                                    title="Edit destination location"
                                  />
                                  <button
                                    type="button"
                                    onClick={() => updateInboundValue(row.id, "isMasked", !row.isMasked)}
                                    title={row.isMasked ? "Click to Reveal Raw" : "Click to Mask Value"}
                                    className="p-1 px-1.5 text-slate-400 hover:text-slate-600 rounded border border-slate-200 bg-white hover:bg-slate-50 shrink-0 transition-all cursor-pointer"
                                  >
                                    {row.isMasked ? <Lock size={12} className="text-emerald-500" /> : <Unlock size={12} />}
                                  </button>
                                </div>
                                {row.isMasked && (
                                  <div className="mt-1 flex flex-col gap-1 pl-1">
                                    <div className="flex items-center gap-1.5">
                                      <span className="text-[10px] text-emerald-600 shrink-0 font-semibold font-mono">Masked Name:</span>
                                      <input
                                        type="text"
                                        value={row.maskedName || ""}
                                        onChange={(e) => updateInboundValue(row.id, "maskedName", e.target.value)}
                                        placeholder={maskDestination(row.destination) || "Custom masked alias..."}
                                        className="w-full bg-emerald-50/40 hover:bg-white focus:bg-white text-slate-750 focus:ring-1 focus:ring-emerald-500 rounded border border-emerald-100 px-2 py-0.5 text-[10px] font-mono transition-all focus:border-emerald-500 placeholder:text-slate-400 focus:outline-hidden"
                                        title="Manually enter custom masked name"
                                      />
                                    </div>
                                    <span className="block text-[9px] font-mono text-slate-450 pl-0.5" title="Final active representation">
                                      Active view: <strong className="text-emerald-700 font-bold font-mono">{getDisplayMaskedName(row)}</strong>
                                    </span>
                                  </div>
                                )}
                              </div>
                            </td>

                            {/* Planned Minutes Input */}
                            <td className="py-2 px-2">
                              <input
                                type="number"
                                min="0"
                                value={row.plannedMinutes === 0 ? "" : row.plannedMinutes}
                                onChange={(e) => updateInboundValue(row.id, "plannedMinutes", Math.max(0, parseInt(e.target.value) || 0))}
                                placeholder="0"
                                className="w-full font-mono text-xs text-right bg-slate-50 hover:bg-white focus:bg-white text-slate-800 focus:ring-1 focus:ring-emerald-500 rounded border border-slate-200 hover:border-slate-300 px-2 py-1 transition-all focus:border-emerald-500 focus:outline-hidden"
                                title="Edit Planned Minutes"
                              />
                            </td>

                            {/* RPM Input (Rate Per Minute) */}
                            <td className="py-2 px-2">
                              <div className="relative">
                                <span className="absolute left-1.5 top-1/2 -translate-y-1/2 text-[9px] text-slate-400 font-mono">$</span>
                                <input
                                  type="number"
                                  min="0"
                                  step="0.0001"
                                  value={row.rpm === 0 ? "" : row.rpm}
                                  onChange={(e) => updateInboundValue(row.id, "rpm", Math.max(0, parseFloat(e.target.value) || 0))}
                                  placeholder="0.0000"
                                  className="w-full font-mono text-xs text-right bg-slate-50 hover:bg-white focus:bg-white text-slate-800 focus:ring-1 focus:ring-emerald-500 rounded border border-slate-200 hover:border-slate-300 pl-4 pr-1.5 py-1 transition-all focus:border-emerald-500 focus:outline-hidden"
                                  title="Edit Revenue Per Minute (RPM)"
                                />
                              </div>
                            </td>

                            {/* Automatically Calculated Planned Revenue Result */}
                            <td className="py-2 px-3 text-right font-mono font-semibold text-slate-800 text-xs">
                              ${plannedRevenue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                            </td>

                            {/* Actions Column (Delete and Detailed Edit Modal trigger) */}
                            <td className="py-2 px-3 text-center">
                              <div className="flex items-center justify-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setEditingInbound(row)}
                                  className="p-1.5 text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 rounded border border-slate-100 bg-white transition-colors cursor-pointer"
                                  title="Edit in full interactive modal form"
                                >
                                  <Pencil size={12} />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => deleteInboundRow(row.id)}
                                  className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded border border-slate-100 bg-white transition-colors cursor-pointer"
                                  title="Delete Inbound row"
                                >
                                  <Trash2 size={12} />
                                </button>
                              </div>
                            </td>
                          </motion.tr>
                        );
                      })
                    )}
                  </AnimatePresence>
                </tbody>
              </table>
            </div>

            {/* Summation Footer (Requirement: automatically calculates summation) */}
            <div className="mt-auto border-t border-slate-200 bg-emerald-50/20 p-4 font-mono text-xs">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 text-slate-700">
                <div className="font-bold text-slate-900 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block animate-pulse" />
                  INBOUND ROW SUMS
                </div>
                
                <div className="flex items-center gap-6 justify-between sm:justify-end">
                  {/* Total Minutes Column Indicator */}
                  <div className="flex flex-col items-end">
                    <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Total Planned Mins</span>
                    <span className="text-sm font-bold text-slate-800">
                      {totals.totalInboundMinutes.toLocaleString()} mins
                    </span>
                  </div>

                  {/* Total Planned Revenue Indicator */}
                  <div className="flex flex-col items-end">
                    <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Total Planned Rev</span>
                    <span className="text-sm font-bold text-emerald-800">
                      ${totals.totalPlannedRevenue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                  </div>
                </div>
              </div>
            </div>

          </div>

          {/* ========================================================= */}
          {/* TABLE TWO: OUTBOUND                                        */}
          {/* ========================================================= */}
          <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs flex flex-col" id="outbound-card">
            
            {/* Table Header Section */}
            <div className="p-5 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="p-2 bg-amber-50 text-amber-600 rounded-xl">
                  <ArrowUpRight size={20} />
                </span>
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Outbound Voice Swaps</h2>
                  <p className="text-xs text-slate-500">Traffic sent to partners' networks (Cost stream)</p>
                </div>
              </div>
              
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={toggleAllOutboundMask}
                  title="Toggle destination privacy for Outbound"
                  className="px-2.5 py-1.5 text-xs font-medium border border-slate-200 bg-white rounded-lg hover:bg-slate-50 text-slate-600 transition-colors inline-flex items-center gap-1 cursor-pointer"
                >
                  {allOutboundMasked ? (
                    <>
                      <Eye size={13} className="text-sky-600" />
                      <span className="hidden sm:inline">Show Original</span>
                    </>
                  ) : (
                    <>
                      <EyeOff size={13} className="text-slate-400" />
                      <span className="hidden sm:inline">Hide Original</span>
                    </>
                  )}
                </button>

                <button
                  type="button"
                  onClick={addOutboundRow}
                  className="px-3 py-1.5 text-xs font-semibold text-white bg-amber-600 hover:bg-amber-700 rounded-lg transition-colors inline-flex items-center gap-1 shadow-xs cursor-pointer"
                >
                  <Plus size={14} />
                  Add Row
                </button>
              </div>
            </div>

            {/* Table Core Container */}
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-150 text-xs text-slate-500 uppercase bg-slate-50/30">
                    <th className="py-3 px-3 font-semibold w-10 text-center">#</th>
                    <th className="py-3 px-3 font-semibold min-w-[200px]">Masked Destination</th>
                    <th className="py-3 px-2 font-semibold w-24 text-right">Planned Min</th>
                    <th className="py-3 px-2 font-semibold w-24 text-right">cost (CPM $)</th>
                    <th className="py-3 px-3 text-right font-semibold w-28">planned cost</th>
                    <th className="py-3 px-3 font-semibold w-20 text-center">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-150">
                  <AnimatePresence initial={false}>
                    {outboundRows.length === 0 ? (
                      <motion.tr 
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        className="text-center"
                      >
                        <td colSpan={6} className="py-12 px-4 text-sm text-slate-400">
                          <PackageOpen className="mx-auto text-slate-300 mb-2" size={32} />
                          No Outbound routes added yet.
                        </td>
                      </motion.tr>
                    ) : (
                      outboundRows.map((row, idx) => {
                        const plannedCost = row.plannedMinutes * row.cpm;
                        return (
                          <motion.tr
                            key={row.id}
                            initial={{ opacity: 0, y: 15 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, x: 15, transition: { duration: 0.15 } }}
                            transition={{ duration: 0.15 }}
                            className="hover:bg-slate-50/80 text-sm align-middle group"
                          >
                            {/* Counter Index */}
                            <td className="py-2 px-3 text-center text-xs font-mono text-slate-400">
                              {idx + 1}
                            </td>

                            {/* Destination Input (Clearly demarcated form fields) */}
                            <td className="py-2 px-2">
                              <div className="flex flex-col gap-1">
                                <div className="flex items-center gap-1">
                                  <input
                                    type="text"
                                    value={row.destination}
                                    onChange={(e) => updateOutboundValue(row.id, "destination", e.target.value)}
                                    placeholder="Destination location..."
                                    className="w-full bg-slate-50 hover:bg-white focus:bg-white text-slate-800 focus:ring-1 focus:ring-amber-500 rounded border border-slate-200 hover:border-slate-300 px-2.5 py-1 text-xs transition-all focus:border-amber-500 font-medium placeholder:text-slate-300 focus:outline-hidden"
                                    title="Edit destination location"
                                  />
                                  <button
                                    type="button"
                                    onClick={() => updateOutboundValue(row.id, "isMasked", !row.isMasked)}
                                    title={row.isMasked ? "Click to Reveal Raw" : "Click to Mask Value"}
                                    className="p-1 px-1.5 text-slate-400 hover:text-slate-600 rounded border border-slate-200 bg-white hover:bg-slate-50 shrink-0 transition-all cursor-pointer"
                                  >
                                    {row.isMasked ? <Lock size={12} className="text-amber-500" /> : <Unlock size={12} />}
                                  </button>
                                </div>
                                {row.isMasked && (
                                  <div className="mt-1 flex flex-col gap-1 pl-1">
                                    <div className="flex items-center gap-1.5">
                                      <span className="text-[10px] text-amber-600 shrink-0 font-semibold font-mono">Masked Name:</span>
                                      <input
                                        type="text"
                                        value={row.maskedName || ""}
                                        onChange={(e) => updateOutboundValue(row.id, "maskedName", e.target.value)}
                                        placeholder={maskDestination(row.destination) || "Custom masked alias..."}
                                        className="w-full bg-amber-50/40 hover:bg-white focus:bg-white text-slate-755 focus:ring-1 focus:ring-amber-500 rounded border border-amber-100 px-2 py-0.5 text-[10px] font-mono transition-all focus:border-amber-500 placeholder:text-slate-400 focus:outline-hidden"
                                        title="Manually enter custom masked name"
                                      />
                                    </div>
                                    <span className="block text-[9px] font-mono text-slate-450 pl-0.5" title="Final active representation">
                                      Active view: <strong className="text-amber-700 font-bold font-mono">{getDisplayMaskedName(row)}</strong>
                                    </span>
                                  </div>
                                )}
                              </div>
                            </td>

                            {/* Planned Minutes Input */}
                            <td className="py-2 px-2">
                              <input
                                type="number"
                                min="0"
                                value={row.plannedMinutes === 0 ? "" : row.plannedMinutes}
                                onChange={(e) => updateOutboundValue(row.id, "plannedMinutes", Math.max(0, parseInt(e.target.value) || 0))}
                                placeholder="0"
                                className="w-full font-mono text-xs text-right bg-slate-50 hover:bg-white focus:bg-white text-slate-800 focus:ring-1 focus:ring-amber-500 rounded border border-slate-200 hover:border-slate-300 px-2 py-1 transition-all focus:border-amber-500 focus:outline-hidden"
                                title="Edit Planned Minutes"
                              />
                            </td>

                            {/* Cost Input (CPM Rate) */}
                            <td className="py-2 px-2">
                              <div className="relative">
                                <span className="absolute left-1.5 top-1/2 -translate-y-1/2 text-[9px] text-slate-400 font-mono">$</span>
                                <input
                                  type="number"
                                  min="0"
                                  step="0.0001"
                                  value={row.cpm === 0 ? "" : row.cpm}
                                  onChange={(e) => updateOutboundValue(row.id, "cpm", Math.max(0, parseFloat(e.target.value) || 0))}
                                  placeholder="0.0000"
                                  className="w-full font-mono text-xs text-right bg-slate-50 hover:bg-white focus:bg-white text-slate-800 focus:ring-1 focus:ring-amber-500 rounded border-slate-200 hover:border-slate-300 pl-4 pr-1.5 py-1 transition-all focus:border-amber-500 focus:outline-hidden"
                                  title="Edit Cost Rate (CPM)"
                                />
                              </div>
                            </td>

                            {/* Automated planned cost product */}
                            <td className="py-2 px-3 text-right font-mono font-semibold text-slate-800 text-xs">
                              ${plannedCost.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                            </td>

                            {/* Actions Column */}
                            <td className="py-2 px-3 text-center">
                              <div className="flex items-center justify-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setEditingOutbound(row)}
                                  className="p-1.5 text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 rounded border border-slate-100 bg-white transition-colors cursor-pointer"
                                  title="Edit in full interactive modal form"
                                >
                                  <Pencil size={12} />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => deleteOutboundRow(row.id)}
                                  className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded border border-slate-100 bg-white transition-colors cursor-pointer"
                                  title="Delete Outbound row"
                                >
                                  <Trash2 size={12} />
                                </button>
                              </div>
                            </td>
                          </motion.tr>
                        );
                      })
                    )}
                  </AnimatePresence>
                </tbody>
              </table>
            </div>

            {/* Summation Footer (Requirement: automatically calculates summation) */}
            <div className="mt-auto border-t border-slate-200 bg-amber-50/20 p-4 font-mono text-xs">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 text-slate-700">
                <div className="font-bold text-slate-900 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-amber-500 inline-block animate-pulse" />
                  OUTBOUND ROW SUMS
                </div>
                
                <div className="flex items-center gap-6 justify-between sm:justify-end">
                  {/* Total Minutes Column Indicator */}
                  <div className="flex flex-col items-end">
                    <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Total Planned Mins</span>
                    <span className="text-sm font-bold text-slate-800">
                      {totals.totalOutboundMinutes.toLocaleString()} mins
                    </span>
                  </div>

                  {/* Total Planned Cost Indicator */}
                  <div className="flex flex-col items-end">
                    <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Total Planned Cost</span>
                    <span className="text-sm font-bold text-amber-800">
                      ${totals.totalPlannedCost.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                  </div>
                </div>
              </div>
            </div>

          </div>

        </div>

        {/* Dynamic Joint Telecom Swaps Operations Analysis */}
        <section className="mt-8 bg-slate-900 text-slate-100 p-6 rounded-2xl border border-slate-800 shadow-xl" id="analytics-section">
          <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
            <Coins size={18} className="text-indigo-400" />
            Bilateral Swap Yield & Settlement Performance
          </h3>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            
            {/* Performance Indicators */}
            <div className="space-y-3">
              <div className="flex justify-between text-sm py-1 border-b border-slate-800">
                <span className="text-slate-400">Inbound Traffic Share:</span>
                <span className="font-mono text-emerald-400 font-semibold">
                  {totals.totalInboundMinutes + totals.totalOutboundMinutes > 0 
                    ? ((totals.totalInboundMinutes / (totals.totalInboundMinutes + totals.totalOutboundMinutes)) * 100).toFixed(0) 
                    : 0}%
                </span>
              </div>
              <div className="flex justify-between text-sm py-1 border-b border-slate-800">
                <span className="text-slate-400">Outbound Traffic Share:</span>
                <span className="font-mono text-amber-400 font-semibold">
                  {totals.totalInboundMinutes + totals.totalOutboundMinutes > 0 
                    ? ((totals.totalOutboundMinutes / (totals.totalInboundMinutes + totals.totalOutboundMinutes)) * 100).toFixed(0) 
                    : 0}%
                </span>
              </div>
              <div className="flex justify-between text-sm py-1">
                <span className="text-slate-400">Total Swapped Traffic:</span>
                <span className="font-mono text-slate-200 font-semibold">
                  {(totals.totalInboundMinutes + totals.totalOutboundMinutes).toLocaleString()} mins
                </span>
              </div>
            </div>

            {/* Visual breakdown progress bars */}
            <div className="flex flex-col justify-center space-y-4">
              <div>
                <div className="flex justify-between text-xs text-slate-400 mb-1">
                  <span>Average Inbound Rate (RPM)</span>
                  <span className="font-mono text-emerald-400">
                    ${totals.totalInboundMinutes > 0 ? (totals.totalPlannedRevenue / totals.totalInboundMinutes).toFixed(4) : "0.0000"}
                  </span>
                </div>
                <div className="w-full bg-slate-800 h-2 rounded-full overflow-hidden">
                  <div 
                    className="bg-emerald-500 h-2 rounded-full" 
                    style={{ width: `${Math.min(100, Math.max(0, (totals.totalInboundMinutes > 0 ? (totals.totalPlannedRevenue / totals.totalInboundMinutes) : 0) * 1000))}%` }}
                  />
                </div>
              </div>
              <div>
                <div className="flex justify-between text-xs text-slate-400 mb-1">
                  <span>Average Outbound Rate (CPM)</span>
                  <span className="font-mono text-amber-400">
                    ${totals.totalOutboundMinutes > 0 ? (totals.totalPlannedCost / totals.totalOutboundMinutes).toFixed(4) : "0.0000"}
                  </span>
                </div>
                <div className="w-full bg-slate-800 h-2 rounded-full overflow-hidden">
                  <div 
                    className="bg-amber-500 h-2 rounded-full" 
                    style={{ width: `${Math.min(100, Math.max(0, (totals.totalOutboundMinutes > 0 ? (totals.totalPlannedCost / totals.totalOutboundMinutes) : 0) * 1000))}%` }}
                  />
                </div>
              </div>
            </div>

            {/* Quick Summary Statement */}
            <div className="flex flex-col justify-between bg-slate-800/50 p-4 rounded-xl border border-slate-800">
              <span className="text-xs text-slate-400 uppercase tracking-widest font-mono">Deal Profitability Status</span>
              <p className="text-sm text-slate-300 mt-2">
                {totals.netMargin > 0 ? (
                  <>
                    Bilateral swap settlement parameters are currently <strong className="text-emerald-400">Receivable</strong>. Inbound revenue speed exceeds outbound cost rates, yielding a net surplus pool of <strong className="text-emerald-400">${totals.netMargin.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</strong>.
                  </>
                ) : totals.netMargin < 0 ? (
                  <>
                    Bilateral swap balance results in a <strong className="text-rose-400">Net Payable</strong> offset. Outbound cost corridors exceed inbound revenues. Try adjusting rates or committed voice volumes.
                  </>
                ) : (
                  <>
                    Inbound and Outbound logs are balanced. Add/edit telecom prefix details to recalculate swap settlement instantly.
                  </>
                )}
              </p>
            </div>

          </div>
        </section>

      </div>

      {/* ========================================================= */}
      {/* INBOUND EDIT MODAL                                        */}
      {/* ========================================================= */}
      <AnimatePresence>
        {editingInbound && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" id="modal-inbound-editor">
            {/* Backdrop */}
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setEditingInbound(null)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-xs" 
            />
            
            {/* Modal Body */}
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, y: 15 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 15 }}
              className="relative w-full max-w-lg bg-white rounded-2xl border border-slate-200 shadow-2xl overflow-hidden z-10"
            >
              <div className="p-6 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <span className="p-1.5 bg-emerald-50 text-emerald-600 rounded-lg">
                    <ArrowDownLeft size={18} />
                  </span>
                  <div>
                    <h3 className="text-lg font-bold text-slate-900">Edit Inbound Row Parameters</h3>
                    <p className="text-xs text-slate-500">Row ID: {editingInbound.id}</p>
                  </div>
                </div>
                <button
                  onClick={() => setEditingInbound(null)}
                  className="p-1.5 hover:bg-slate-200 rounded-lg text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
                >
                  <X size={16} />
                </button>
              </div>

              <form onSubmit={handleSaveInboundEdit} className="p-6 space-y-4">
                {/* Destination / Masked Name Swap */}
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-1">
                    {editingInbound.isMasked ? "Manual Masked Destination Name (Optional)" : "Destination Name / Prefix"}
                  </label>
                  {editingInbound.isMasked ? (
                    <input
                      type="text"
                      value={editingInbound.maskedName || ""}
                      onChange={(e) => setEditingInbound({ ...editingInbound, maskedName: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-2 text-sm transition-all text-slate-800 placeholder:text-slate-400 focus:outline-hidden"
                      placeholder={maskDestination(editingInbound.destination) || "Enter a custom masked alias..."}
                    />
                  ) : (
                    <input
                      type="text"
                      required
                      value={editingInbound.destination}
                      onChange={(e) => setEditingInbound({ ...editingInbound, destination: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-2 text-sm transition-all text-slate-800 placeholder:text-slate-300 focus:outline-hidden"
                      placeholder="e.g. Saudi Arabia STC Mobiles (966)"
                    />
                  )}
                </div>

                {/* Sub-group inputs */}
                <div className="grid grid-cols-2 gap-4">
                  {/* Planned Minutes */}
                  <div>
                    <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-1">
                      Planned Minutes
                    </label>
                    <input
                      type="number"
                      min="0"
                      required
                      value={editingInbound.plannedMinutes}
                      onChange={(e) => setEditingInbound({ ...editingInbound, plannedMinutes: Math.max(0, parseInt(e.target.value) || 0) })}
                      className="w-full font-mono bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-2 text-sm transition-all text-slate-800 placeholder:text-slate-300 focus:outline-hidden"
                    />
                  </div>

                  {/* RPM Rate */}
                  <div>
                    <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-1">
                      Revenue Rate (RPM)
                    </label>
                    <div className="relative">
                      <span className="absolute left-3 top-1/2 -translate-y-1/2 text-xs text-slate-400 font-mono">$</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        required
                        value={editingInbound.rpm}
                        onChange={(e) => setEditingInbound({ ...editingInbound, rpm: Math.max(0, parseFloat(e.target.value) || 0) })}
                        className="w-full font-mono bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg pl-7 pr-3 py-2 text-sm transition-all text-slate-800 placeholder:text-slate-300 focus:outline-hidden"
                      />
                    </div>
                  </div>
                </div>

                 {/* Mask privacy setting toggle */}
                <div className="flex flex-col gap-3.5 p-3.5 bg-slate-50 border border-slate-100 rounded-xl">
                  <div className="flex items-center justify-between">
                    <div>
                      <span className="block text-sm font-semibold text-slate-800">
                        Destination Obfuscation (Masking)
                      </span>
                      <span className="block text-xs text-slate-400">
                        Hides raw location names via standard character masking rules
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() => setEditingInbound({ ...editingInbound, isMasked: !editingInbound.isMasked })}
                      className={`p-2 rounded-lg border flex items-center gap-1.5 text-xs font-semibold cursor-pointer ${
                        editingInbound.isMasked 
                          ? 'border-emerald-200 bg-emerald-50 text-emerald-700' 
                          : 'border-slate-200 bg-white text-slate-500'
                      }`}
                    >
                      {editingInbound.isMasked ? (
                        <>
                          <Lock size={14} />
                          Masked
                        </>
                      ) : (
                        <>
                          <Unlock size={14} />
                          Raw Plaintext
                        </>
                      )}
                    </button>
                  </div>
                  {editingInbound.isMasked ? (
                    <div className="border-t border-slate-200/60 pt-3">
                      <label className="block text-xs font-bold text-slate-650 uppercase tracking-wider mb-1.5">
                        Destination Name / Prefix
                      </label>
                      <input
                        type="text"
                        required
                        value={editingInbound.destination}
                        onChange={(e) => setEditingInbound({ ...editingInbound, destination: e.target.value })}
                        className="w-full bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-1.5 text-xs transition-all text-slate-850 placeholder:text-slate-400 focus:outline-hidden"
                        placeholder="e.g. United Kingdom Mobiles O2 (447)"
                      />
                    </div>
                  ) : (
                    <div className="border-t border-slate-200/60 pt-3">
                      <label className="block text-xs font-bold text-slate-650 uppercase tracking-wider mb-1.5">
                        Manual Masked Destination Name (Optional)
                      </label>
                      <input
                        type="text"
                        value={editingInbound.maskedName || ""}
                        onChange={(e) => setEditingInbound({ ...editingInbound, maskedName: e.target.value })}
                        className="w-full bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-1.5 text-xs transition-all text-slate-850 placeholder:text-slate-400 focus:outline-hidden"
                        placeholder={maskDestination(editingInbound.destination) || "Enter custom masked alias..."}
                      />
                    </div>
                  )}
                </div>

                {/* Calculation math helper readout */}
                <div className="p-4 bg-emerald-50/50 rounded-xl border border-emerald-100/50 flex justify-between items-center text-sm font-mono">
                  <span className="text-slate-500 text-xs">Calculated Planned Revenue:</span>
                  <span className="text-emerald-800 font-bold text-base">
                    ${(editingInbound.plannedMinutes * editingInbound.rpm).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </span>
                </div>

                {/* Footer Controls */}
                <div className="flex gap-2 justify-end pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setEditingInbound(null)}
                    className="px-4 py-2 text-xs font-bold text-slate-600 bg-white border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-705 rounded-lg transition-colors shadow-xs flex items-center gap-1.5 cursor-pointer"
                  >
                    <Check size={14} />
                    Apply Changes
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* ========================================================= */}
      {/* OUTBOUND EDIT MODAL                                       */}
      {/* ========================================================= */}
      <AnimatePresence>
        {editingOutbound && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" id="modal-outbound-editor">
            {/* Backdrop */}
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setEditingOutbound(null)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-xs" 
            />
            
            {/* Modal Body */}
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, y: 15 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 15 }}
              className="relative w-full max-w-lg bg-white rounded-2xl border border-slate-200 shadow-2xl overflow-hidden z-10"
            >
              <div className="p-6 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <span className="p-1.5 bg-amber-50 text-amber-600 rounded-lg">
                    <ArrowUpRight size={18} />
                  </span>
                  <div>
                    <h3 className="text-lg font-bold text-slate-900">Edit Outbound Row Parameters</h3>
                    <p className="text-xs text-slate-500">Row ID: {editingOutbound.id}</p>
                  </div>
                </div>
                <button
                  onClick={() => setEditingOutbound(null)}
                  className="p-1.5 hover:bg-slate-200 rounded-lg text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
                >
                  <X size={16} />
                </button>
              </div>

              <form onSubmit={handleSaveOutboundEdit} className="p-6 space-y-4">
                {/* Destination / Masked Name Swap */}
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-1">
                    {editingOutbound.isMasked ? "Manual Masked Destination Name (Optional)" : "Destination Name / Prefix"}
                  </label>
                  {editingOutbound.isMasked ? (
                    <input
                      type="text"
                      value={editingOutbound.maskedName || ""}
                      onChange={(e) => setEditingOutbound({ ...editingOutbound, maskedName: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-2 text-sm transition-all text-slate-800 placeholder:text-slate-400 focus:outline-hidden"
                      placeholder={maskDestination(editingOutbound.destination) || "Enter a custom masked alias..."}
                    />
                  ) : (
                    <input
                      type="text"
                      required
                      value={editingOutbound.destination}
                      onChange={(e) => setEditingOutbound({ ...editingOutbound, destination: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-2 text-sm transition-all text-slate-800 placeholder:text-slate-300 focus:outline-hidden"
                      placeholder="e.g. United Kingdom Mobiles Vodafone (447)"
                    />
                  )}
                </div>

                {/* Sub-group inputs */}
                <div className="grid grid-cols-2 gap-4">
                  {/* Planned Minutes */}
                  <div>
                    <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-1">
                      Planned Minutes
                    </label>
                    <input
                      type="number"
                      min="0"
                      required
                      value={editingOutbound.plannedMinutes}
                      onChange={(e) => setEditingOutbound({ ...editingOutbound, plannedMinutes: Math.max(0, parseInt(e.target.value) || 0) })}
                      className="w-full font-mono bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-2 text-sm transition-all text-slate-800 placeholder:text-slate-300 focus:outline-hidden"
                    />
                  </div>

                  {/* CPM Cost Rate */}
                  <div>
                    <label className="block text-xs font-bold text-slate-600 uppercase tracking-widest mb-1">
                      Cost per Minute (CPM)
                    </label>
                    <div className="relative">
                      <span className="absolute left-3 top-1/2 -translate-y-1/2 text-xs text-slate-400 font-mono">$</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        required
                        value={editingOutbound.cpm}
                        onChange={(e) => setEditingOutbound({ ...editingOutbound, cpm: Math.max(0, parseFloat(e.target.value) || 0) })}
                        className="w-full font-mono bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg pl-7 pr-3 py-2 text-sm transition-all text-slate-800 placeholder:text-slate-300 focus:outline-hidden"
                      />
                    </div>
                  </div>
                </div>

                {/* Mask privacy setting toggle */}
                <div className="flex flex-col gap-3.5 p-3.5 bg-slate-50 border border-slate-100 rounded-xl">
                  <div className="flex items-center justify-between">
                    <div>
                      <span className="block text-sm font-semibold text-slate-800">
                        Destination Obfuscation (Masking)
                      </span>
                      <span className="block text-xs text-slate-400">
                        Hides raw location names via character replacement rules
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() => setEditingOutbound({ ...editingOutbound, isMasked: !editingOutbound.isMasked })}
                      className={`p-2 rounded-lg border flex items-center gap-1.5 text-xs font-semibold cursor-pointer ${
                        editingOutbound.isMasked 
                          ? 'border-amber-200 bg-amber-50 text-amber-700' 
                          : 'border-slate-200 bg-white text-slate-500'
                      }`}
                    >
                      {editingOutbound.isMasked ? (
                        <>
                          <Lock size={14} />
                          Masked
                        </>
                      ) : (
                        <>
                          <Unlock size={14} />
                          Raw Plaintext
                        </>
                      )}
                    </button>
                  </div>
                  {editingOutbound.isMasked ? (
                    <div className="border-t border-slate-200/60 pt-3">
                      <label className="block text-xs font-bold text-slate-650 uppercase tracking-wider mb-1.5">
                        Destination Name / Prefix
                      </label>
                      <input
                        type="text"
                        required
                        value={editingOutbound.destination}
                        onChange={(e) => setEditingOutbound({ ...editingOutbound, destination: e.target.value })}
                        className="w-full bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-1.5 text-xs transition-all text-slate-855 placeholder:text-slate-400 focus:outline-hidden"
                        placeholder="e.g. United Kingdom Mobiles Vodafone (447)"
                      />
                    </div>
                  ) : (
                    <div className="border-t border-slate-200/60 pt-3">
                      <label className="block text-xs font-bold text-slate-650 uppercase tracking-wider mb-1.5">
                        Manual Masked Destination Name (Optional)
                      </label>
                      <input
                        type="text"
                        value={editingOutbound.maskedName || ""}
                        onChange={(e) => setEditingOutbound({ ...editingOutbound, maskedName: e.target.value })}
                        className="w-full bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg px-3 py-1.5 text-xs transition-all text-slate-855 placeholder:text-slate-400 focus:outline-hidden"
                        placeholder={maskDestination(editingOutbound.destination) || "Enter custom masked alias..."}
                      />
                    </div>
                  )}
                </div>

                {/* Calculation cost readout */}
                <div className="p-4 bg-amber-50/50 rounded-xl border border-amber-100/50 flex justify-between items-center text-sm font-mono">
                  <span className="text-slate-500 text-xs">Calculated Planned Cost:</span>
                  <span className="text-amber-800 font-bold text-base">
                    ${(editingOutbound.plannedMinutes * editingOutbound.cpm).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </span>
                </div>

                {/* Footer Controls */}
                <div className="flex gap-2 justify-end pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setEditingOutbound(null)}
                    className="px-4 py-2 text-xs font-bold text-slate-600 bg-white border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2 text-xs font-bold text-white bg-amber-600 hover:bg-amber-705 rounded-lg transition-colors shadow-xs flex items-center gap-1.5 cursor-pointer"
                  >
                    <Check size={14} />
                    Apply Changes
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

    </div>
  );
}
