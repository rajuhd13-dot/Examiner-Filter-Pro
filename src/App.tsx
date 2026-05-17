import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { 
  Users, 
  Search, 
  RefreshCcw, 
  Printer, 
  X, 
  AlertCircle,
  FileSpreadsheet,
  ChevronLeft,
  ChevronRight,
  Loader2,
  CheckCircle2,
  ArrowRight,
  Database,
  RefreshCw,
  ShieldAlert,
  FileDown
} from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import axios from "axios";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { useAuth, AuthProvider } from "./components/auth-context";
import { FilterDropdown } from "./components/FilterDropdown";
import { cn } from "./lib/utils";

// Types
interface Subject {
  key: string;
  label: string;
}

interface FilterOptions {
  institutes: string[];
  departments: string[];
  batches: string[];
  trainings: string[];
  campuses: string[];
  tpins: string[];
  subjects: Subject[];
  rowCount: number;
}

interface SearchFilters {
  institute: string[];
  department: string[];
  batch: string[];
  trainingsSelected: string[];
  campusesSelected: string[];
  tpinsSelected: string[];
  subjectsSelected: string[];
  onlyAllowed: boolean;
  subjectLogic: "all" | "any";
  allowEnglish: number | null;
  allowOthers: number | null;
}

interface FilterResult {
  header: string[];
  rows: string[][];
  total: number;
  page: number;
  pageSize: number;
  totalPages: number;
  allow: { ENGLISH: number; OTHERS: number };
}

const DEFAULT_FILTERS: SearchFilters = {
  institute: [],
  department: [],
  batch: [],
  trainingsSelected: [],
  campusesSelected: [],
  tpinsSelected: [],
  subjectsSelected: [],
  onlyAllowed: true,
  subjectLogic: "any",
  allowEnglish: 55,
  allowOthers: 48,
};

const Dashboard: React.FC = () => {
  const { user, token, initialized, login, logout, isLoggingIn } = useAuth();
  const [options, setOptions] = useState<FilterOptions | null>(() => {
    const saved = localStorage.getItem("ex_options");
    return saved ? JSON.parse(saved) : null;
  });
  const [filters, setFilters] = useState<SearchFilters>(() => {
    try {
      const saved = localStorage.getItem("ex_filters");
      return saved ? { ...DEFAULT_FILTERS, ...JSON.parse(saved) } : DEFAULT_FILTERS;
    } catch { return DEFAULT_FILTERS; }
  });

  const [result, setResult] = useState<FilterResult | null>(() => {
    try {
      const saved = localStorage.getItem("ex_last_res");
      return saved ? JSON.parse(saved) : null;
    } catch { return null; }
  });
  const [fullResult, setFullResult] = useState<FilterResult | null>(null);
  
  const [isLoading, setIsLoading] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const fetchAbortControllerRef = useRef<AbortController | null>(null);
  
  // Persistence effect
  useEffect(() => {
    localStorage.setItem("ex_filters", JSON.stringify(filters));
  }, [filters]);
  const [errorDetails, setErrorDetails] = useState<{ message: string; advice?: string } | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(200);
  const [autoRefresh, setAutoRefresh] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState<Date>(new Date());
  const [hasSheetUpdates, setHasSheetUpdates] = useState(false);

  // Initialize data
  useEffect(() => {
    if (initialized) {
      console.log("[App] Component initialized, loading options...");
      loadOptions();
      const cleanup = startPolling();
      return cleanup;
    }
  }, [initialized, token]); 

  // Polling for sheet updates
  const startPolling = () => {
    const timer = setInterval(async () => {
      if (!token || !initialized) return;
      try {
        const res = await axios.get("/api/sync", {
          headers: token ? { Authorization: `Bearer ${token}` } : {}
        });
        if (res.data.success) {
          setLastSyncTime(new Date());
          if (options && res.data.rowCount !== options.rowCount) {
            setHasSheetUpdates(true);
            if (autoRefresh) handleSearch();
          }
        }
      } catch (err) {
        console.warn("[App] Sync/Polling error:", err);
      }
    }, 60000); // Check every 60s
    return () => clearInterval(timer);
  };

  const [isPinged, setIsPinged] = useState<boolean | null>(null);

  const checkPing = async () => {
    try {
      const res = await axios.get("/api/ping");
      const ok = !!(res.data && res.data.success);
      setIsPinged(ok);
      console.log("[App] Backend Ping:", ok ? "Online" : "Offline");
    } catch (err: any) {
      setIsPinged(false);
      console.error("[App] Ping failed:", err.message);
    }
  };

  const loadOptions = async (forceRefresh = false) => {
    if (isSyncing) return;
    setIsSyncing(true);
    setErrorDetails(null);
    console.log(`[App] Fetching initial filter options (force: ${forceRefresh})...`);
    
    // Non-blocking background checks
    checkPing();

    try {
      if (forceRefresh) {
        await axios.get("/api/clearCache", {
          headers: token ? { Authorization: `Bearer ${token}` } : {}
        });
      }

      const res = await axios.get("/api/options", {
        headers: token ? { Authorization: `Bearer ${token}` } : {},
        timeout: 150000 
      });
      
      if (res.data && res.data.success) {
        setOptions(res.data);
        localStorage.setItem("ex_options", JSON.stringify(res.data));
        setHasSheetUpdates(false);
        if (res.data.rowCount !== undefined) {
          setLastSyncTime(new Date());
        }
      } else {
        const msg = res.data?.error || "The backend returned an unsuccessful response.";
        setErrorDetails({ 
          message: msg, 
          advice: res.data?.advice 
        });
      }
    } catch (err: any) {
      const msg = err.response?.data?.error || err.message || "Failed to reach the backend server.";
      setErrorDetails({ 
        message: msg, 
        advice: err.response?.data?.advice 
      });
    } finally {
      setIsSyncing(false);
    }
  };

  const handleLookup = useCallback(async (query: string) => {
    if (!query || !token) return;
    setIsLoading(true);
    setErrorDetails(null);
    try {
      const res = await axios.get(`/api/lookup?query=${encodeURIComponent(query)}`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (res.data.success && res.data.found) {
        setResult({
          header: res.data.header,
          rows: [res.data.row],
          total: 1,
          page: 1,
          pageSize: 1,
          totalPages: 1,
          allow: { ENGLISH: 55, OTHERS: 48 }
        });
      } else {
        alert("No examiner found with that ID or Mobile number.");
      }
    } catch (err: any) {
      console.error("Lookup failed", err);
      setErrorDetails({ 
        message: "Lookup failed", 
        advice: "Check your connection and try again." 
      });
    } finally {
      setIsLoading(false);
    }
  }, [token]);

  const handleSearch = useCallback(async (newPage = 1, isSilent = false) => {
    // Validation: Require at least one specific filter or search query
    const hasSpecificFilter = 
      filters.institute.length > 0 ||
      filters.department.length > 0 ||
      filters.batch.length > 0 ||
      filters.trainingsSelected.length > 0 ||
      filters.campusesSelected.length > 0 ||
      filters.tpinsSelected.length > 0 ||
      (filters.subjectsSelected && filters.subjectsSelected.length > 0) ||
      searchQuery.trim().length > 0;

    if (!hasSpecificFilter && !isSilent) {
      if (result) {
        setResult(null); // Clear previous results if everything is deselected
        setFullResult(null);
      }
      return;
    } else if (!hasSpecificFilter && isSilent) {
      setResult(null);
      setFullResult(null);
      return;
    }

    // Remove strict token requirement as the backend proxy doesn't strictly enforce it for now 
    // and we want to avoid silent failures on page refresh if token isn't restored.
    // We still log if it's missing.
    if (!token) {
      console.warn("[App] Searching without token (might be after refresh)");
    }
    
    if (searchQuery.trim() && newPage === 1) {
      handleLookup(searchQuery.trim());
      return;
    }

    if (!isSilent) setIsLoading(true);
    setHasSheetUpdates(false);
    setErrorDetails(null);

    // Cancel pending requests to prevent 502 overloading
    if (fetchAbortControllerRef.current) {
      fetchAbortControllerRef.current.abort();
    }
    fetchAbortControllerRef.current = new AbortController();

    try {
      console.log("[App] Executing search for page:", newPage);
      const res = await axios.post("/api/filter", {
        filters,
        page: newPage,
        pageSize,
        returnAll: false
      }, {
        headers: token ? { Authorization: `Bearer ${token}` } : {},
        timeout: 120000,
        signal: fetchAbortControllerRef.current.signal
      });

      if (res.data && res.data.success) {
        setResult(res.data);
        localStorage.setItem("ex_last_res", JSON.stringify(res.data));
        setPage(newPage);
        
        // Phase 2: Background fetch all for export is removed to prevent GAS concurrency locking and 30-40s delays.
        // Full export can be done on-demand or use currently loaded results.
        setFullResult(null);

        // Scroll to results if not silent
        if (!isSilent) {
          setTimeout(() => {
            const el = document.getElementById("results-section");
            if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }, 100);
        }
      } else {
        setErrorDetails({ 
          message: res.data?.error || "Search failed", 
          advice: res.data?.advice || "Please check your connection and script deployment."
        });
      }
    } catch (err: any) {
      if (axios.isCancel(err)) {
        console.log("[App] Search cancelled due to new request.");
        return;
      }
      console.error("[App] Search error:", err);
      
      const errMsg = err.response?.data?.error || err.message || "Search failed";
      
      // Improve 502 / overload messaging so it doesn't look like a total crash
      setErrorDetails({ 
        message: errMsg.includes("502") ? "Backend Service Timeout (502)" : errMsg, 
        advice: errMsg.includes("502") || err.code === "ECONNABORTED" 
          ? "The server is overloaded from too many requests. Please wait 1-2 minutes and press 'Search' again." 
          : "The backend might be overloaded or the script is unavailable. Try again in a moment."
      });
    } finally {
      if (!fetchAbortControllerRef.current?.signal.aborted) {
        setIsLoading(false);
      }
    }
  }, [token, filters, pageSize, searchQuery, handleLookup]);

  const handleClear = () => {
    setFilters(DEFAULT_FILTERS);
    setSearchQuery("");
    setResult(null);
    setFullResult(null);
    setPage(1);
    setHasSheetUpdates(false);
    localStorage.removeItem("ex_last_res");
    localStorage.removeItem("ex_filters");
  };

  // Clear results if all filters and search query are removed
  useEffect(() => {
    const hasSpecificFilter = 
      filters.institute.length > 0 ||
      filters.department.length > 0 ||
      filters.batch.length > 0 ||
      filters.trainingsSelected.length > 0 ||
      filters.campusesSelected.length > 0 ||
      filters.tpinsSelected.length > 0 ||
      (filters.subjectsSelected && filters.subjectsSelected.length > 0);
      
    if (!hasSpecificFilter && searchQuery.trim().length === 0) {
       setResult(null);
       setFullResult(null);
       setPage(1);
    }
  }, [filters, searchQuery]);

  const handleExportExcel = () => {
    const dataToExport = fullResult || result;
    if (!dataToExport) return;

    // Transform data for Excel (similar to PDF logic if needed, but usually Excel users want raw data)
    // However, if we want to match the "Comment" column request for Excel too:
    const excelHeader = [...dataToExport.header];
    const statusIdx = excelHeader.findIndex(h => h.toLowerCase().includes("allow status"));
    if (statusIdx !== -1) excelHeader[statusIdx] = "Comment";

    const excelRows = dataToExport.rows.map(row => {
      const newRow = [...row];
      if (statusIdx !== -1) newRow[statusIdx] = ""; 
      return newRow;
    });

    const worksheet = XLSX.utils.aoa_to_sheet([excelHeader, ...excelRows]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Examiners");
    
    XLSX.writeFile(workbook, `Examiner_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleExportPDF = () => {
    const dataToExport = fullResult || result;
    if (!dataToExport) return;

    // Transform data for printing as per user request:
    // 1. Rename "Allow Status" to "Comment"
    // 2. Clear values in that column (make boxes empty)
    const printHeader = [...dataToExport.header];
    const statusIdx = printHeader.findIndex(h => h.toLowerCase().includes("allow status"));
    
    if (statusIdx !== -1) {
      printHeader[statusIdx] = "Comment";
    }

    const printRows = dataToExport.rows.map(row => {
      const newRow = [...row];
      if (statusIdx !== -1) {
        newRow[statusIdx] = ""; // Clear content
      }
      return newRow;
    });

    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const pageW = doc.internal.pageSize.getWidth();
    
    // Title
    doc.setFontSize(14);
    doc.setTextColor(26, 86, 219);
    doc.text("Examiner Filter Report", pageW / 2, 10, { align: "center" });
    
    // Summary
    doc.setFontSize(9);
    doc.setTextColor(100);
    doc.text(`Found ${dataToExport.total} records | Generated: ${new Date().toLocaleString()}`, pageW / 2, 16, { align: "center" });
    
    autoTable(doc, {
      head: [printHeader],
      body: printRows,
      startY: 22,
      styles: { fontSize: 7, cellPadding: 1, halign: 'center' },
      headStyles: { fillColor: [26, 86, 219], textColor: 255 },
      margin: { left: 5, right: 5 }
    });

    window.open(doc.output("bloburl"), "_blank");
  };

  if (!initialized) return null;

  if (!user) {
    return (
      <div className="min-h-screen bg-[#F8FAFC] flex flex-col items-center justify-center p-6">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="w-full max-w-md bg-white p-10 rounded-[32px] shadow-2xl border border-gray-100 flex flex-col items-center text-center"
        >
          <div className="w-16 h-16 bg-blue-50 rounded-2xl flex items-center justify-center mb-6">
            <Users className="w-8 h-8 text-blue-600" />
          </div>
          <h1 className="text-3xl font-extrabold text-gray-900 mb-2 tracking-tight">Examiner Filter Pro</h1>
          <p className="text-gray-500 mb-8 text-sm leading-relaxed">
            Professional filtering and data management tool for <br/>Examiner Information. Please sign in to continue.
          </p>
          
          <button
            onClick={login}
            disabled={isLoggingIn}
            className="w-full flex items-center justify-center gap-3 bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 px-6 rounded-2xl transition-all shadow-lg shadow-blue-200 active:scale-95 disabled:opacity-50"
          >
            {isLoggingIn ? (
              <Loader2 className="w-5 h-5 animate-spin" />
            ) : (
              <Users className="w-5 h-5" />
            )}
            Sign in with Google
          </button>
          
          <p className="mt-8 text-[11px] text-gray-400 font-medium uppercase tracking-[2px]">
            Powered by Google AI Studio
          </p>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#F0F4F9] text-gray-900 font-sans pb-12">
      {/* Top Navigation */}
      <header className="sticky top-0 z-[100] bg-blue-600 shadow-md">
        <div className="max-w-[1800px] mx-auto px-6 h-16 flex items-center justify-between gap-6">
          <div className="flex items-center gap-4">
            <div className="bg-white/10 p-2 rounded-xl border border-white/20">
              <FileSpreadsheet className="w-5 h-5 text-white" />
            </div>
            <h1 className="text-lg font-bold text-white tracking-tight">Examiner Filter Pro</h1>
            
            <div className="flex items-center gap-2 px-3 py-1 bg-white/10 rounded-full border border-white/20 ml-2">
              <div className={cn("w-2 h-2 rounded-full", hasSheetUpdates ? "bg-amber-400 animate-pulse" : "bg-green-400")} />
              <span className="text-[10px] font-bold text-white uppercase tracking-wider">
                {hasSheetUpdates ? "New Data Available" : "Live Sync"}
              </span>
            </div>
          </div>

          <div className="flex-1 max-w-4xl bg-white p-1.5 px-5 rounded-[20px] shadow-sm flex items-center justify-between gap-4 ml-4">
            <div className="flex items-center gap-6 shrink-0">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-blue-50 rounded-lg">
                  <Users className="w-3.5 h-3.5 text-blue-600" />
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-[9px] font-bold text-gray-400 capitalize whitespace-nowrap">Total Records</span>
                  <div className="flex items-center gap-1">
                    <span className="text-xs font-black text-gray-800 leading-none">
                      {options?.rowCount !== undefined ? options.rowCount : "-"}
                    </span>
                    <button 
                      onClick={() => loadOptions(true)}
                      className={cn("p-0.5 hover:bg-gray-100 rounded transition-colors", isSyncing && "animate-spin")}
                      title="Force Refresh Data"
                    >
                      <RefreshCw className="w-2.5 h-2.5 text-gray-300" />
                    </button>
                  </div>
                </div>
              </div>
              
              <div className="w-px h-6 bg-gray-100" />

              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-indigo-50 rounded-lg">
                  <div className={cn(
                    "w-3.5 h-3.5 rounded-full flex items-center justify-center transition-all",
                    isSyncing ? "bg-blue-500 animate-pulse" : (options ? "bg-emerald-500" : (isPinged ? "bg-amber-400" : "bg-red-500"))
                  )}>
                    <div className="w-1.5 h-1.5 bg-white rounded-full opacity-60" />
                  </div>
                </div>
                <div className="flex flex-col">
                  <span className="text-[9px] font-bold text-gray-400 capitalize whitespace-nowrap">Backend Status</span>
                  <div className="flex items-center gap-1.5">
                    <span className={cn(
                      "text-[10px] font-black leading-none uppercase tracking-wider",
                      options ? "text-emerald-600" : (isPinged ? "text-amber-600" : "text-red-600")
                    )}>
                      {isSyncing ? "Syncing..." : (options ? "Connected" : (isPinged === false ? "Offline" : "Checking..."))}
                    </span>
                    {!options && !isSyncing && (
                      <button 
                        onClick={() => loadOptions()}
                        className="p-1 hover:bg-gray-100 rounded-md transition-colors"
                        title="Retry Connection"
                      >
                        <RefreshCw className="w-2.5 h-2.5 text-gray-400" />
                      </button>
                    )}
                  </div>
                </div>
              </div>

              {result && (
                <>
                  <div className="w-px h-6 bg-gray-100" />
                  <div className="flex items-center gap-2 text-blue-600">
                    <div className="p-1.5 bg-blue-50 rounded-lg">
                      <CheckCircle2 className="w-3.5 h-3.5" />
                    </div>
                    <div className="flex flex-col">
                      <span className="text-[9px] font-bold opacity-60 capitalize whitespace-nowrap">Results</span>
                      <span className="text-xs font-black leading-none">{result.total} matching</span>
                    </div>
                  </div>
                </>
              )}
            </div>

            {/* Quick Search */}
            <div className="flex-1 flex items-center gap-2 max-w-sm px-3 py-1.5 bg-gray-50 rounded-xl border border-gray-100 focus-within:border-blue-400 focus-within:bg-white transition-all">
              <Search className="w-3.5 h-3.5 text-gray-400" />
              <input 
                type="text"
                placeholder="Quick lookup (T-PIN or Mobile)..."
                className="w-full bg-transparent border-none outline-none text-xs font-medium placeholder:text-gray-400"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                onKeyDown={async (e) => {
                  if (e.key === 'Enter') {
                    handleLookup(searchQuery);
                  }
                }}
              />
              {searchQuery ? (
                <button 
                  onClick={() => setSearchQuery("")}
                  className="p-1 hover:bg-gray-200 rounded-full transition-colors"
                >
                  <X className="w-3 h-3 text-gray-400" />
                </button>
              ) : (
                <div className="px-1.5 py-0.5 bg-gray-200/50 rounded text-[9px] font-bold text-gray-400 uppercase tracking-tighter">Enter</div>
              )}
            </div>

            <div className="flex items-center gap-2 shrink-0 justify-end">
              <div className="flex items-center gap-3 mr-1 pr-3 border-r border-gray-100">
                 <button
                    onClick={() => {
                       if (searchQuery.trim()) {
                         handleLookup(searchQuery.trim());
                       } else {
                         handleSearch(1);
                       }
                    }}
                    disabled={isLoading}
                    className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-xl font-black text-xs transition-all shadow-md shadow-blue-100 active:scale-95 disabled:opacity-50 whitespace-nowrap"
                  >
                    {isLoading ? <Loader2 className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />}
                    {searchQuery.trim() ? "Lookup" : "Search"}
                  </button>
                  <button
                    onClick={handleClear}
                    className="flex items-center gap-1.5 group text-gray-400 hover:text-gray-600 transition-colors shrink-0"
                  >
                    <X className="w-3.5 h-3.5 group-hover:rotate-90 transition-transform duration-300" />
                    <span className="text-xs font-bold">Clear</span>
                  </button>
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                  <button
                    onClick={handleExportPDF}
                    disabled={!result}
                    className="p-2 bg-purple-50 text-purple-700 hover:bg-purple-100 rounded-lg font-bold transition-all disabled:opacity-40"
                    title="Print PDF Report"
                  >
                    <Printer className="w-4 h-4" />
                  </button>
                  <button
                    onClick={handleExportExcel}
                    disabled={!result}
                    className="p-2 bg-green-50 text-green-700 hover:bg-green-100 rounded-lg font-bold transition-all disabled:opacity-40"
                    title="Export Excel (XLSX)"
                  >
                    <FileDown className="w-4 h-4" />
                  </button>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-[1800px] mx-auto p-6 flex flex-col gap-6">

        {/* Filter Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 xl:grid-cols-4 gap-x-6 gap-y-3 bg-white p-8 rounded-[32px] shadow-sm border border-gray-100">
          <FilterDropdown 
            label="Institute Name" 
            placeholder="Select Institute" 
            options={options?.institutes || []} 
            selected={filters.institute}
            onChange={(v) => setFilters(prev => ({ ...prev, institute: v }))}
          />
          <FilterDropdown 
            label="Department Name" 
            placeholder="Select Department" 
            options={options?.departments || []} 
            selected={filters.department}
            onChange={(v) => setFilters(prev => ({ ...prev, department: v }))}
          />
          <FilterDropdown 
            label="HSC Batch" 
            placeholder="Select Batch" 
            options={options?.batches || []} 
            selected={filters.batch}
            onChange={(v) => setFilters(prev => ({ ...prev, batch: v }))}
          />
          <FilterDropdown 
            label="Subject" 
            placeholder="Select Subjects" 
            options={options?.subjects?.map(s => ({ label: s.label, value: s.key })) || []} 
            selected={filters.subjectsSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, subjectsSelected: v }))}
            emptyMeansAll
          />
          <FilterDropdown 
            label="Training Report" 
            placeholder="Select Training" 
            options={options?.trainings || []} 
            selected={filters.trainingsSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, trainingsSelected: v }))}
          />
          <FilterDropdown 
            label="Physical Campus" 
            placeholder="Select Campus" 
            options={options?.campuses || []} 
            selected={filters.campusesSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, campusesSelected: v }))}
          />
          <FilterDropdown 
            label="T-PIN" 
            placeholder="Select T-PIN" 
            options={options?.tpins || []} 
            selected={filters.tpinsSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, tpinsSelected: v }))}
          />

          <div className="flex flex-col gap-3 p-4 bg-gray-50/50 rounded-[24px] border border-gray-100">
            <div className="flex flex-col gap-4">
              {/* Settings Row */}
              <div className="flex items-center gap-4">
                <label className="flex items-center gap-2 cursor-pointer group flex-1">
                  <div className={cn(
                    "w-4 h-4 rounded border flex items-center justify-center transition-all",
                    filters.onlyAllowed ? "bg-blue-600 border-blue-600 text-white shadow-sm" : "border-gray-200 bg-white group-hover:border-blue-300"
                  )}>
                    <input 
                      type="checkbox" 
                      className="hidden" 
                      checked={filters.onlyAllowed} 
                      onChange={(e) => setFilters(prev => ({ ...prev, onlyAllowed: e.target.checked }))} 
                    />
                    {filters.onlyAllowed && <CheckCircle2 className="w-3 h-3" />}
                  </div>
                  <span className="text-[11px] font-bold text-gray-600 truncate">Only ALLOWED</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer group flex-1">
                  <div className={cn(
                    "w-4 h-4 rounded border flex items-center justify-center transition-all",
                    filters.subjectLogic === "all" ? "bg-blue-600 border-blue-600 text-white shadow-sm" : "border-gray-200 bg-white group-hover:border-blue-300"
                  )}>
                    <input 
                      type="checkbox" 
                      className="hidden" 
                      checked={filters.subjectLogic === "all"} 
                      onChange={(e) => setFilters(prev => ({ ...prev, subjectLogic: e.target.checked ? "all" : "any" }))} 
                    />
                    {filters.subjectLogic === "all" && <CheckCircle2 className="w-3 h-3" />}
                  </div>
                  <span className="text-[11px] font-bold text-gray-600 truncate">Require ALL Sub</span>
                </label>
              </div>

              {/* Threshold Row */}
              <div className="flex items-center gap-3">
                <div className="flex items-center gap-2 flex-1">
                  <input 
                    type="number" 
                    value={filters.allowEnglish || ""} 
                    onChange={(e) => setFilters(prev => ({ ...prev, allowEnglish: e.target.value === "" ? null : Number(e.target.value) }))}
                    className="w-full bg-white border border-gray-100 rounded-xl px-2 py-1.5 text-center font-bold text-blue-600 text-xs focus:outline-none focus:ring-2 focus:ring-blue-50 focus:border-blue-400 transition-all shadow-sm"
                    placeholder="55"
                  />
                </div>
                <div className="flex items-center gap-2 flex-1">
                  <input 
                    type="number" 
                    value={filters.allowOthers || ""} 
                    onChange={(e) => setFilters(prev => ({ ...prev, allowOthers: e.target.value === "" ? null : Number(e.target.value) }))}
                    className="w-full bg-white border border-gray-100 rounded-xl px-2 py-1.5 text-center font-bold text-blue-600 text-xs focus:outline-none focus:ring-2 focus:ring-blue-50 focus:border-blue-400 transition-all shadow-sm"
                    placeholder="48"
                  />
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Results Section */}
        <div 
          id="results-section"
          className={cn(
            "flex flex-col gap-0 bg-white rounded-[32px] overflow-hidden shadow-sm border border-gray-100 transition-all duration-500",
            result ? "min-h-[400px]" : "min-h-[100px]"
          )}
        >
          <AnimatePresence>
            {hasSheetUpdates && !autoRefresh && (
              <motion.button
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: "auto", opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                onClick={() => handleSearch(1)}
                className="w-full bg-blue-600 text-white font-bold py-3 text-sm flex items-center justify-center gap-2 hover:bg-blue-700 transition-colors"
              >
                <RefreshCcw className="w-4 h-4 animate-spin-slow" />
                Sheet updated! Click to reload with new data
              </motion.button>
            )}
          </AnimatePresence>


          {errorDetails && (
            <div className="p-12 flex flex-col items-center justify-center text-center">
              <div className="w-16 h-16 bg-red-50 text-red-500 rounded-2xl flex items-center justify-center mb-6 shadow-sm">
                <ShieldAlert className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-black text-gray-900 mb-3 tracking-tight">Connectivity Issue</h3>
              <div className="bg-red-50/50 border border-red-100 rounded-2xl p-6 max-w-lg mb-8 text-center text-balance">
                <p className="text-red-800 text-sm font-medium leading-relaxed mb-1">{errorDetails.message}</p>
                {errorDetails.advice && (
                  <p className="text-red-600 text-xs font-bold mt-2 uppercase tracking-wide">💡 {errorDetails.advice}</p>
                )}
                
                <div className="text-left space-y-2 mt-4 pt-4 border-t border-red-100">
                  <p className="text-xs text-red-600 font-bold uppercase tracking-wider">Troubleshooting:</p>
                  <ul className="text-xs text-red-700/70 list-disc list-inside space-y-1">
                    <li>Is GAS deployed as <b>"Anyone"</b> under "Who has access"?</li>
                    <li>Did you copy the <b>NEW Exec URL</b> from the deployment screen?</li>
                    <li>Ensure <b>Execute as: Me</b> is selected in the deployment settings.</li>
                    <li>Verify the <b>GAS_DEPLOYMENT_URL</b> in AI Studio Secrets matches exactly.</li>
                  </ul>
                </div>
              </div>
              <button 
                onClick={() => { setErrorDetails(null); loadOptions(); }}
                className="flex items-center gap-2 bg-gray-900 hover:bg-black text-white px-8 py-3 rounded-2xl font-bold text-sm transition-all active:scale-95 shadow-xl shadow-gray-200"
              >
                <RefreshCw className={cn("w-4 h-4", isSyncing && "animate-spin")} />
                Retry Connection
              </button>
            </div>
          )}

          {!result && !errorDetails && !isLoading && (
            <div className="flex flex-col items-center justify-center py-10 px-4 text-center">
              {!options ? (
                <>
                  <div className="w-20 h-20 bg-amber-50 text-amber-500 rounded-3xl flex items-center justify-center mb-6 animate-pulse shadow-xl shadow-amber-100/50">
                    <Database className="w-10 h-10" />
                  </div>
                  <h3 className="text-2xl font-black text-gray-900 mb-3 tracking-tight">Database Connectivity</h3>
                  <p className="text-gray-500 text-sm max-w-sm leading-relaxed mb-8 font-medium">
                    The application hasn't connected to your Google Sheet yet. Please ensure the backend is deployed correctly.
                  </p>
                  <button 
                    onClick={() => { loadOptions(); }}
                    className="flex items-center gap-2 bg-gray-900 hover:bg-black text-white px-8 py-3 rounded-2xl font-bold text-sm transition-all active:scale-95 shadow-xl shadow-gray-200"
                  >
                    <RefreshCcw className="w-4 h-4" />
                    Connect to Sheet
                  </button>
                </>
              ) : (
                <>
                  <div className="w-16 h-16 bg-blue-50 text-blue-500 rounded-3xl flex items-center justify-center mb-4 shadow-xl shadow-blue-100/50">
                    <Search className="w-8 h-8" />
                  </div>
                  <h3 className="text-xl font-black text-gray-900 tracking-tight">Ready to Search</h3>
                </>
              )}
            </div>
          )}

          {isLoading && !result && (
            <div className="flex-1 flex flex-col items-center justify-center p-20 gap-4">
              <Loader2 className="w-12 h-12 text-blue-600 animate-spin" />
              <span className="text-sm font-bold text-gray-400 capitalize font-mono tracking-widest">Searching......</span>
            </div>
          )}

          {result && (
            <>
              <div className="px-6 py-4 bg-gray-50 border-b border-gray-100 flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="flex items-center gap-2">
                    <span className="px-3 py-1 bg-white border border-gray-200 rounded-lg text-[13px] font-bold text-gray-800">
                      {result.total} Records Found
                    </span>
                    <span className="text-xs text-gray-400 font-medium italic">
                      Showing Page {page} of {result.totalPages}
                    </span>
                  </div>
                </div>

                <div className="flex items-center gap-4">
                  <div className="flex items-center gap-2 bg-white p-1 rounded-xl border border-gray-100 shadow-sm">
                    <button
                      onClick={() => handleSearch(page - 1)}
                      disabled={page <= 1 || isLoading}
                      className="p-1.5 hover:bg-gray-100 rounded-lg transition-colors disabled:opacity-30 disabled:cursor-not-allowed"
                    >
                      <ChevronLeft className="w-5 h-5" />
                    </button>
                    <span className="px-3 text-sm font-bold text-gray-700">{page}</span>
                    <button
                      onClick={() => handleSearch(page + 1)}
                      disabled={page >= result.totalPages || isLoading}
                      className="p-1.5 hover:bg-gray-100 rounded-lg transition-colors disabled:opacity-30 disabled:cursor-not-allowed"
                    >
                      <ChevronRight className="w-5 h-5" />
                    </button>
                  </div>
                  
                  <select 
                    value={pageSize}
                    onChange={(e) => {
                      setPageSize(Number(e.target.value));
                      // Search reset happens after state updates in effect or manual trigger
                    }}
                    className="bg-white border border-gray-100 rounded-xl px-3 py-2 text-xs font-bold shadow-sm focus:outline-none"
                  >
                    {[100, 200, 300, 500].map(size => (
                      <option key={size} value={size}>{size} per page</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="flex-1 overflow-auto relative max-h-[650px] custom-scrollbar">
                <table className="w-full border-collapse text-left text-sm table-auto min-w-[1200px]">
                  <thead className="sticky top-0 z-10">
                    <tr className="bg-blue-600 text-white font-bold capitalize text-[11px]">
                      {result.header.map((h, i) => (
                        <th key={i} className="px-4 py-4 whitespace-nowrap border-r border-blue-500/30 last:border-0 text-center">
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {result.rows.map((row, i) => (
                      <tr key={i} className="hover:bg-blue-50/50 transition-colors group">
                        {row.map((cell, j) => {
                          const isStatus = result.header[j].toLowerCase().includes("allow status");
                          return (
                            <td 
                              key={j} 
                              className={cn(
                                "px-4 py-3 whitespace-nowrap text-center text-gray-600 border-r border-gray-50 last:border-0",
                                isStatus && cell === "ALLOWED" && "text-green-600 font-extrabold",
                                isStatus && cell === "NOT ALLOWED" && "text-red-500 font-extrabold"
                              )}
                            >
                              {cell}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              
              <div className="p-6 bg-gray-50/60 border-t border-gray-100 flex flex-col gap-2">
                 <p className="text-[11px] text-gray-400 font-bold capitalize text-center"> End of Report </p>
                 <div className="flex items-center justify-between mt-2">
                    <span className="text-[10px] text-gray-400">Page sizing set to {pageSize} rows</span>
                    <span className="text-[10px] text-gray-400 italic">Fast rendering mode active</span>
                 </div>
              </div>
            </>
          )}
        </div>
      </main>
      
      {/* Footer Branding */}
      <footer className="mt-4 mb-12 text-center">
        <p className="text-[11px] text-gray-400 font-bold capitalize mb-1">Examiner Information Management</p>
        <p className="text-[9px] text-gray-300">Secure Environment / Data Protected by Google Workspace Policy</p>
      </footer>
    </div>
  );
};

export default function App() {
  return (
    <AuthProvider>
      <Dashboard />
    </AuthProvider>
  );
}
