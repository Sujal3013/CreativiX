
import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import { 
  Download, FileUp, Play, CheckCircle2, AlertCircle, Loader2, 
  Image as ImageIcon, Trash2, StopCircle, FolderArchive, 
  Layers, Settings2, ShieldCheck, Terminal, XCircle, Info, 
  RefreshCw, LayoutDashboard, Database, Activity, Cpu, Moon, Sun, Type,
  Command, Sparkles, ChevronRight
} from 'lucide-react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { KeywordItem, KeywordResult, AppStatus, AdConfigContext, LogEntry } from './types';
import { generateAdImage } from './services/geminiService';

// --- Modular Components ---

const PipelineLog: React.FC<{ logs: LogEntry[] }> = ({ logs }) => {
  const logEndRef = useRef<HTMLDivElement>(null);
  useEffect(() => logEndRef.current?.scrollIntoView({ behavior: 'smooth' }), [logs]);

  return (
    <div className="flex-1 mt-8 border-t border-slate-100 dark:border-zinc-800 pt-6 overflow-hidden flex flex-col">
      <div className="flex items-center gap-2 mb-4 text-[10px] font-black text-slate-400 dark:text-zinc-500 uppercase tracking-widest">
        <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 mr-1.5" />
        <span>CreativiX Runtime Activity</span>
      </div>
      <div className="bg-slate-900 dark:bg-black rounded-2xl p-4 flex-1 overflow-y-auto text-[10px] font-mono space-y-2.5 border border-slate-800 dark:border-zinc-800 shadow-inner custom-scrollbar">
        {logs.length === 0 && <div className="text-slate-600 italic">Initializing CreativiX engine logs...</div>}
        {logs.map((log) => (
          <div key={log.id} className="flex gap-3 animate-in fade-in slide-in-from-left-2 duration-300">
            <span className="text-indigo-500/50 shrink-0">[{log.timestamp}]</span>
            <span className={`break-all ${
              log.type === 'error' ? 'text-red-400' : 
              log.type === 'warning' ? 'text-amber-400' : 
              log.type === 'success' ? 'text-emerald-400' : 'text-slate-300'
            }`}>{log.message}</span>
          </div>
        ))}
        <div ref={logEndRef} />
      </div>
    </div>
  );
};

const ConfigGrid: React.FC<{ config: AdConfigContext }> = ({ config }) => {
  const items = [
    { label: 'Persona', value: config.persona, icon: Cpu },
    { label: 'Creative Text', value: config.include_keyword_text ? 'Smart Typography Enabled' : 'Native (No Text)', icon: Type },
    { label: 'Visual Style', value: config.image_style, icon: ImageIcon },
    { label: 'Trends', value: config.visual_trends, icon: Activity },
    { label: 'Lighting', value: config.lighting_style, icon: ShieldCheck },
    { label: 'Tone', value: config.brand_tone, icon: Settings2 },
  ];

  return (
    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 mb-10">
      {items.map((item, i) => (
        <div key={i} className="bg-white/60 dark:bg-zinc-900/40 backdrop-blur-sm border border-slate-200/50 dark:border-zinc-800 p-4 rounded-2xl shadow-sm hover:shadow-md transition-all">
          <div className="flex items-center gap-3 mb-2">
            <div className="p-2 bg-indigo-50 dark:bg-indigo-900/20 text-indigo-600 dark:text-indigo-400 rounded-lg">
              <item.icon size={16} />
            </div>
            <span className="text-[10px] font-black text-slate-400 dark:text-zinc-500 uppercase tracking-widest">{item.label}</span>
          </div>
          <p className="text-xs font-bold text-slate-700 dark:text-zinc-300 line-clamp-2 leading-relaxed">{item.value}</p>
        </div>
      ))}
    </div>
  );
};

export default function App() {
  const [results, setResults] = useState<KeywordResult[]>([]);
  const [status, setStatus] = useState<AppStatus>(AppStatus.IDLE);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [adConfig, setAdConfig] = useState<AdConfigContext | null>(null);
  const [includeTextInBatch, setIncludeTextInBatch] = useState<boolean>(false);
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    return (localStorage.getItem('theme') as 'light' | 'dark') || 'light';
  });
  const isStopping = useRef(false);

  const activeConfig = useMemo(() => {
    if (!adConfig) return null;
    return { ...adConfig, include_keyword_text: includeTextInBatch };
  }, [adConfig, includeTextInBatch]);

  useEffect(() => {
    localStorage.setItem('theme', theme);
    if (theme === 'dark') {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [theme]);

  const addLog = useCallback((message: string, type: LogEntry['type'] = 'info') => {
    const newLog: LogEntry = {
      id: Math.random().toString(36).substr(2, 9),
      timestamp: new Date().toLocaleTimeString([], { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' }),
      message,
      type
    };
    setLogs(prev => [...prev.slice(-99), newLog]);
  }, []);

  const loadConfig = useCallback(async () => {
    try {
      const response = await fetch('./image_context.json');
      if (!response.ok) throw new Error('Configuration file missing');
      const config = await response.json();
      setAdConfig(config);
      setIncludeTextInBatch(config.include_keyword_text);
      addLog('Global CreativiX Context initialized.', 'success');
    } catch (err: any) {
      addLog(`Config Error: ${err.message}. Engine may be impaired.`, 'error');
    }
  }, [addLog]);

  useEffect(() => { loadConfig(); }, [loadConfig]);

  const toggleTheme = () => setTheme(prev => prev === 'light' ? 'dark' : 'light');

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    addLog(`Ingesting batch data: ${file.name}`);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data: any[] = XLSX.utils.sheet_to_json(ws);

        if (data.length === 0) throw new Error('Sheet is empty');

        const parsedKeywords: KeywordResult[] = data
          .map((row, index): KeywordResult | null => {
            const rawPhrase = row['keyword'] || row['Keyword'] || row['KEYWORD'];
            const phrase = typeof rawPhrase === 'string' ? rawPhrase.trim() : String(rawPhrase || '').trim();
            if (!phrase) return null;

            const kwId = `kw-${index}-${Date.now()}`;
            return {
              keyword: { id: kwId, phrase },
              variants: [
                { id: `${kwId}-v1`, url: '', variantType: 'Product-Focused', status: 'pending' },
                { id: `${kwId}-v2`, url: '', variantType: 'Lifestyle-Context', status: 'pending' }
              ]
            };
          })
          .filter((item): item is KeywordResult => item !== null);

        if (parsedKeywords.length === 0) throw new Error("No 'keyword' column found in sheet");

        setResults(parsedKeywords);
        setStatus(AppStatus.IDLE);
        setProgress({ current: 0, total: parsedKeywords.length });
        addLog(`Pipeline prepared with ${parsedKeywords.length} batch sequences.`, 'success');
      } catch (err: any) {
        addLog(`Parsing Failure: ${err.message}`, 'error');
      }
    };
    reader.readAsBinaryString(file);
  };

  const runPipeline = async () => {
    if (results.length === 0 || !activeConfig) return;
    
    setStatus(AppStatus.PROCESSING);
    isStopping.current = false;
    addLog(`Initiating Automated CreativiX Cycle...`, 'info');
    
    for (let i = 0; i < results.length; i++) {
      if (isStopping.current) {
        addLog('Production halt signaled.', 'warning');
        setStatus(AppStatus.IDLE);
        return;
      }
      
      const currentResult = results[i];
      if (currentResult.variants.every(v => v.status === 'completed')) {
        setProgress(prev => ({ ...prev, current: i + 1 }));
        continue;
      }

      setResults(prev => prev.map((r, idx) => 
        idx === i ? { ...r, variants: r.variants.map(v => ({ ...v, status: 'generating' })) } : r
      ));

      let quotaExceededDetected = false;
      const updatedVariants = await Promise.all(
        currentResult.variants.map(async (v) => {
          try {
            const base64 = await generateAdImage(currentResult.keyword.phrase, v.variantType, activeConfig);
            return { ...v, url: base64, status: 'completed' as const };
          } catch (error: any) {
            if (error.message === 'QUOTA_EXCEEDED') quotaExceededDetected = true;
            return { ...v, status: 'error' as const, errorMessage: error.message };
          }
        })
      );

      setResults(prev => prev.map((r, idx) => idx === i ? { ...r, variants: updatedVariants } : r));

      if (quotaExceededDetected) {
        addLog('API Capacity Limit Reached. CreativiX entering cooldown.', 'error');
        setStatus(AppStatus.QUOTA_EXCEEDED);
        return;
      }

      addLog(`Asset Cluster Generated: "${currentResult.keyword.phrase}"`, 'success');
      setProgress(prev => ({ ...prev, current: i + 1 }));
    }

    addLog('Full batch successfully synthesized.', 'success');
    setStatus(AppStatus.COMPLETED);
  };

  const downloadAllAsZip = async () => {
    addLog('Assembling CreativiX package...');
    const zip = new JSZip();
    const dateStr = new Date().toISOString().split('T')[0]; 
    const rootFolder = zip.folder(`CreativiX_Export_${dateStr}`);
    
    if (!rootFolder) return;

    let count = 0;
    results.forEach((res) => {
      res.variants.forEach((v, idx) => {
        if (v.status === 'completed' && v.url) {
          const base64Data = v.url.split(',')[1];
          const sanitizedPhrase = res.keyword.phrase.toLowerCase().replace(/[^a-z0-9]/g, '_');
          rootFolder.file(`${sanitizedPhrase}_v${idx + 1}.png`, base64Data, { base64: true });
          count++;
        }
      });
    });

    if (count === 0) {
      addLog('Zero assets ready for deployment.', 'warning');
      return;
    }

    const content = await zip.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(content);
    link.download = `CreativiX_Export_${dateStr}_${count}_assets.zip`;
    link.click();
    addLog(`Package with ${count} assets delivered.`, 'success');
  };

  const completedAssets = useMemo(() => 
    results.reduce((acc, r) => acc + r.variants.filter(v => v.status === 'completed').length, 0),
  [results]);

  const failedAssets = useMemo(() => 
    results.reduce((acc, r) => acc + r.variants.filter(v => v.status === 'error').length, 0),
  [results]);

  const totalPossibleAssets = results.length * 2;
  const hasRemainingWork = completedAssets < totalPossibleAssets;

  const visibleResults = useMemo(() => 
    results.filter(r => r.variants.some(v => v.status !== 'pending')),
  [results]);

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-zinc-950 text-slate-900 dark:text-zinc-100 overflow-hidden transition-colors duration-500">
      {/* Sidebar Dashboard */}
      <aside className="w-80 bg-white dark:bg-zinc-900 border-r border-slate-200 dark:border-zinc-800 flex flex-col p-8 z-50 transition-colors duration-500">
        <div className="flex items-center gap-3 mb-12">
          <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white shadow-xl shadow-indigo-100 dark:shadow-indigo-900/20 ring-4 ring-indigo-50 dark:ring-indigo-900/10">
            <Sparkles size={22} />
          </div>
          <span className="font-black text-xl tracking-tighter text-slate-900 dark:text-white">
            Creativi<span className="text-indigo-600 dark:text-indigo-400">X</span>
          </span>
        </div>
        
        <div className="space-y-6 flex-1 overflow-hidden flex flex-col">
          <section>
            <p className="text-[10px] font-black text-slate-400 dark:text-zinc-500 uppercase tracking-[0.2em] mb-4">Pipeline Status</p>
            <div className={`p-4 rounded-2xl border flex items-center justify-between shadow-sm transition-all ${adConfig ? 'bg-emerald-50 dark:bg-emerald-900/10 border-emerald-100 dark:border-emerald-900/20' : 'bg-red-50 dark:bg-red-900/10 border-red-100 dark:border-red-900/20 animate-pulse'}`}>
              <div className="flex items-center gap-3">
                <ShieldCheck className={adConfig ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-600 dark:text-red-400'} size={18} />
                <span className="text-xs font-bold text-slate-700 dark:text-zinc-300">Engine Sync</span>
              </div>
              {adConfig ? <CheckCircle2 className="text-emerald-600 dark:text-emerald-400" size={16} /> : <AlertCircle className="text-red-600 dark:text-red-400" size={16} />}
            </div>
          </section>

          <section>
            <p className="text-[10px] font-black text-slate-400 dark:text-zinc-500 uppercase tracking-[0.2em] mb-4">Batch Velocity</p>
            <div className="bg-slate-50 dark:bg-zinc-800/50 p-5 rounded-2xl border border-slate-100 dark:border-zinc-800 space-y-4 shadow-inner">
               <div className="flex justify-between items-end">
                  <span className="text-3xl font-black text-slate-900 dark:text-white leading-none">{Math.round((progress.current / (progress.total || 1)) * 100)}%</span>
                  <span className="text-[10px] font-black text-slate-400 dark:text-zinc-500 uppercase tracking-widest">{progress.current}/{progress.total}</span>
               </div>
               <div className="h-2 w-full bg-slate-200 dark:bg-zinc-700 rounded-full overflow-hidden">
                  <div 
                    className="h-full bg-indigo-600 dark:bg-indigo-500 transition-all duration-1000 ease-out"
                    style={{ width: `${(progress.current / (progress.total || 1)) * 100}%` }}
                  />
               </div>
            </div>
          </section>

          <PipelineLog logs={logs} />
        </div>

        <div className="mt-8">
           <button 
             onClick={loadConfig}
             className="w-full flex items-center justify-center gap-2 py-3 bg-slate-50 dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700 rounded-xl text-xs font-black text-slate-600 dark:text-zinc-400 hover:bg-white dark:hover:bg-zinc-700 hover:border-indigo-200 dark:hover:border-indigo-800 transition-all uppercase tracking-widest"
           >
             <RefreshCw size={14} className={status === AppStatus.PROCESSING ? 'animate-spin' : ''} />
             Reload Core Logic
           </button>
        </div>
      </aside>

      {/* Main Orchestrator */}
      <main className="flex-1 flex flex-col overflow-hidden relative">
        <header className="h-20 bg-white/80 dark:bg-zinc-950/80 backdrop-blur-3xl border-b border-slate-200/60 dark:border-zinc-800/60 flex items-center px-10 sticky top-0 z-40 transition-all duration-500">
           
           {/* Left: Logic Status */}
           <div className="flex-1 flex items-center gap-6">
              <div className="flex items-center gap-3">
                 <div className={`flex items-center gap-2 px-3 py-1.5 rounded-full text-[10px] font-black uppercase tracking-widest border transition-all ${
                   status === AppStatus.PROCESSING 
                   ? 'bg-indigo-50 dark:bg-indigo-900/20 border-indigo-100 dark:border-indigo-900/30 text-indigo-700 dark:text-indigo-400' 
                   : 'bg-emerald-50 dark:bg-emerald-900/20 border-emerald-100 dark:border-emerald-900/30 text-emerald-700 dark:text-emerald-400'
                 }`}>
                    <div className={`w-1.5 h-1.5 rounded-full ${status === AppStatus.PROCESSING ? 'bg-indigo-500 animate-pulse' : 'bg-emerald-500'}`} />
                    {status === AppStatus.PROCESSING ? 'Sequence Active' : 'System Ready'}
                 </div>
                 {status === AppStatus.PROCESSING && (
                   <div className="flex items-center gap-2 text-slate-400 dark:text-zinc-500 text-[10px] font-bold uppercase tracking-widest animate-in fade-in slide-in-from-left-2">
                     <ChevronRight size={12} />
                     <span>Synthesizing...</span>
                   </div>
                 )}
              </div>
           </div>
           
           {/* Center: System Control Zone */}
           <div className="flex items-center bg-slate-100/80 dark:bg-zinc-900/80 border border-slate-200/60 dark:border-zinc-800/60 rounded-2xl p-1 shadow-sm transition-all duration-300">
              <div className="flex items-center gap-4 px-4 py-1.5">
                 <div className="flex items-center gap-2.5">
                    <div className={`p-1.5 rounded-lg transition-colors ${includeTextInBatch ? 'bg-indigo-100 dark:bg-indigo-900/40 text-indigo-600 dark:text-indigo-400' : 'bg-slate-200 dark:bg-zinc-800 text-slate-400'}`}>
                      <Type size={14} />
                    </div>
                    <div className="flex flex-col">
                      <span className="text-[9px] font-black text-slate-500 dark:text-zinc-400 uppercase leading-none mb-1">Text Overlay</span>
                      <span className="text-[8px] font-bold text-slate-400 dark:text-zinc-500 uppercase leading-none">{includeTextInBatch ? 'Enabled' : 'Disabled'}</span>
                    </div>
                 </div>
                 
                 <button 
                  onClick={() => setIncludeTextInBatch(!includeTextInBatch)}
                  disabled={status === AppStatus.PROCESSING}
                  aria-label="Toggle keyword text inclusion"
                  className={`relative inline-flex h-6 w-11 shrink-0 cursor-pointer items-center rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none disabled:opacity-50 ${includeTextInBatch ? 'bg-indigo-600' : 'bg-slate-300 dark:bg-zinc-700'}`}
                >
                  <span 
                    aria-hidden="true" 
                    className={`pointer-events-none inline-block h-4 w-4 transform rounded-full bg-white shadow-lg ring-0 transition duration-200 ease-in-out ${includeTextInBatch ? 'translate-x-5' : 'translate-x-0'}`} 
                  />
                </button>
              </div>

              <div className="w-[1px] h-8 bg-slate-200/60 dark:bg-zinc-800/60 mx-1" />

              <button 
               onClick={toggleTheme}
               className="p-2.5 text-slate-400 hover:text-indigo-600 dark:hover:text-indigo-400 transition-all rounded-xl"
               title="Toggle Dark Mode"
              >
               {theme === 'light' ? <Moon size={18} /> : <Sun size={18} />}
             </button>
           </div>
           
           {/* Right: Mission Actions */}
           <div className="flex-1 flex justify-end items-center gap-4">
             {results.length > 0 && status !== AppStatus.PROCESSING && (
                <button 
                  onClick={() => { setResults([]); setStatus(AppStatus.IDLE); setProgress({ current: 0, total: 0 }); }}
                  className="flex items-center gap-2 px-4 py-2 text-slate-400 hover:text-red-500 rounded-xl transition-all text-[10px] font-black uppercase tracking-widest"
                >
                  <Trash2 size={15} />
                  Clear Pipeline
                </button>
             )}
             
             {status === AppStatus.PROCESSING ? (
                <button 
                  onClick={() => { isStopping.current = true; }}
                  className="px-8 py-2.5 bg-red-600 hover:bg-red-700 text-white font-black rounded-xl flex items-center gap-2.5 transition-all text-[10px] uppercase tracking-widest shadow-xl shadow-red-500/20"
                >
                  <StopCircle size={18} />
                  Terminate
                </button>
             ) : (
               <div className="flex gap-3">
                 {completedAssets > 0 && (
                   <button 
                     onClick={downloadAllAsZip}
                     className="px-6 py-2.5 bg-emerald-600 dark:bg-emerald-500 text-white font-black rounded-xl flex items-center gap-2.5 hover:scale-[1.02] shadow-xl transition-all text-[10px] uppercase tracking-widest"
                   >
                     <FolderArchive size={18} />
                     Export Results ({completedAssets})
                   </button>
                 )}
                 
                 {results.length > 0 && hasRemainingWork && (
                    <button 
                      onClick={runPipeline}
                      disabled={!activeConfig}
                      className={`px-8 py-2.5 bg-indigo-600 dark:bg-indigo-500 text-white font-black rounded-xl flex items-center gap-2.5 hover:scale-[1.02] shadow-2xl transition-all text-[10px] uppercase tracking-widest disabled:opacity-50`}
                    >
                      <Sparkles size={18} />
                      {completedAssets > 0 ? 'Resume Sequence' : 'Forge Ad Creatives'}
                    </button>
                 )}

                 {results.length === 0 && (
                   <div className="flex items-center gap-3 text-slate-300 dark:text-zinc-700 text-[10px] font-black uppercase tracking-[0.3em] cursor-default select-none">
                      <Command size={16} />
                      Awaiting Input
                   </div>
                 )}
               </div>
             )}
           </div>
        </header>

        <div className="flex-1 overflow-y-auto p-12 bg-pattern custom-scrollbar">
          {results.length === 0 ? (
            <div className="max-w-4xl mx-auto py-12 text-center space-y-12 animate-in fade-in duration-1000">
               <div className="space-y-4">
                  <h1 className="text-6xl font-black text-slate-900 dark:text-white tracking-tighter leading-none transition-colors">
                    Creativi<span className="text-transparent bg-clip-text bg-gradient-to-r from-indigo-600 to-violet-600 dark:from-indigo-400 dark:to-violet-400 font-black">X</span>
                  </h1>
                  <p className="text-slate-500 dark:text-zinc-400 text-xl font-medium max-w-2xl mx-auto">Automate high-fidelity creative production with a professional sequence optimized for conversion performance.</p>
               </div>

               <div className="relative group max-w-2xl mx-auto">
                 <div className="absolute -inset-4 bg-gradient-to-r from-indigo-500 to-violet-500 rounded-[3rem] opacity-20 blur-2xl group-hover:opacity-30 transition duration-1000" />
                 <div className="relative bg-white dark:bg-zinc-900 border-2 border-dashed border-slate-200 dark:border-zinc-800 rounded-[2.5rem] p-20 text-center hover:border-indigo-400 dark:hover:border-indigo-500 transition-all shadow-sm hover:shadow-2xl">
                    <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileChange} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                    <div className="w-28 h-28 bg-indigo-50 dark:bg-indigo-900/20 text-indigo-600 dark:text-indigo-400 rounded-[2.5rem] flex items-center justify-center mx-auto mb-10 group-hover:scale-110 transition-transform shadow-inner">
                      <Database size={48} />
                    </div>
                    <h3 className="text-2xl font-bold text-slate-800 dark:text-white">Load Keyword Source</h3>
                    <p className="text-slate-400 dark:text-zinc-500 mt-4 text-sm font-medium tracking-tight">Format: Excel/CSV with header 'keyword'</p>
                 </div>
               </div>

               {activeConfig && <ConfigGrid config={activeConfig} />}
            </div>
          ) : (
            <div className="max-w-[1600px] mx-auto space-y-10 animate-in fade-in slide-in-from-bottom-5 duration-700">
               <div className="bg-white dark:bg-zinc-900 border border-slate-200 dark:border-zinc-800 rounded-[2.5rem] p-10 flex flex-col md:flex-row items-center justify-between gap-10 shadow-sm relative overflow-hidden transition-colors duration-500">
                  <div className="absolute top-0 right-0 w-96 h-96 bg-indigo-500/5 rounded-full blur-3xl -mr-48 -mt-48" />
                  <div className="relative space-y-2">
                    <h2 className="text-3xl font-black text-slate-900 dark:text-white tracking-tighter">Active Batch Progress</h2>
                    <p className="text-slate-500 dark:text-zinc-400 font-medium">Forging high-performance variants for {progress.total} keyword clusters.</p>
                  </div>
                  <div className="relative flex gap-12">
                     <Stat label="Total Clusters" value={progress.total} />
                     <Stat label="Success" value={completedAssets} color="text-emerald-500 dark:text-emerald-400" />
                     <Stat label="Exceptions" value={failedAssets} color="text-red-500 dark:text-red-400" />
                  </div>
               </div>

               <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-4 gap-8 pb-40">
                  {visibleResults.map((res) => (
                    <KeywordCard key={res.keyword.id} result={res} />
                  ))}
               </div>
            </div>
          )}
        </div>

        {/* Floating Progress Overlay */}
        {status === AppStatus.PROCESSING && (
           <div className="fixed bottom-10 left-1/2 -translate-x-1/2 bg-slate-900 dark:bg-zinc-100 text-white dark:text-zinc-900 px-10 py-5 rounded-full shadow-2xl flex items-center gap-6 border border-slate-800 dark:border-zinc-300 animate-in slide-in-from-bottom-10 transition-colors duration-500 z-[60]">
              <div className="flex items-center gap-3 border-r border-slate-700 dark:border-zinc-300 pr-6">
                 <Loader2 className="animate-spin text-indigo-400 dark:text-indigo-600" size={20} />
                 <span className="text-[11px] font-black uppercase tracking-[0.2em]">Processing</span>
              </div>
              <p className="text-xs font-bold text-slate-300 dark:text-zinc-600">Keyword <span className="text-white dark:text-zinc-900">{progress.current + 1}</span> of {progress.total}</p>
              <div className="w-32 h-1.5 bg-slate-800 dark:bg-zinc-200 rounded-full overflow-hidden">
                 <div className="h-full bg-indigo-500 dark:bg-indigo-600 transition-all duration-1000" style={{ width: `${(progress.current/progress.total)*100}%` }} />
              </div>
           </div>
        )}
      </main>
    </div>
  );
}

const Stat: React.FC<{ label: string; value: number; color?: string }> = ({ label, value, color = "text-slate-900 dark:text-white" }) => (
  <div className="text-center">
    <div className={`text-5xl font-black tracking-tighter ${color}`}>{value}</div>
    <div className="text-[10px] font-black text-slate-400 dark:text-zinc-500 uppercase tracking-widest mt-2">{label}</div>
  </div>
);

// --- New Component for Meaningful Loader ---
const GeneratingOverlay: React.FC<{ variantType: string }> = ({ variantType }) => {
  const [message, setMessage] = useState("Analyzing semantic intent...");
  
  useEffect(() => {
    const messages = [
      "Analyzing semantic intent...",
      "Composing diegetic typography...",
      "Simulating natural lighting...",
      "Optimizing for high conversion...",
      "Refining documentary textures...",
      "Anchoring visual subjects...",
      "Calibrating realism engine...",
      "Integrating candid context..."
    ];
    
    // Pick initial based on variant
    if (variantType === 'Product-Focused') {
      setMessage("Magnifying product details...");
    } else {
      setMessage("Integrating candid context...");
    }

    const interval = setInterval(() => {
      setMessage(messages[Math.floor(Math.random() * messages.length)]);
    }, 2500);
    
    return () => clearInterval(interval);
  }, [variantType]);

  return (
    <div className="absolute inset-0 bg-white/80 dark:bg-zinc-900/80 backdrop-blur-md flex flex-col items-center justify-center z-10 p-4 text-center animate-in fade-in duration-300">
      <Loader2 className="animate-spin text-indigo-600 dark:text-indigo-400 mb-2" size={24} />
      <span className="text-[9px] font-black text-indigo-600 dark:text-indigo-400 uppercase tracking-tighter leading-tight animate-pulse">
        {message}
      </span>
    </div>
  );
};

const KeywordCard: React.FC<{ result: KeywordResult }> = ({ result }) => {
  const isDone = result.variants.every(v => v.status === 'completed');
  return (
    <div className={`bg-white dark:bg-zinc-900 rounded-3xl border border-slate-200 dark:border-zinc-800 shadow-sm overflow-hidden group hover:shadow-xl hover:-translate-y-1 transition-all duration-500 animate-in fade-in slide-in-from-bottom-5`}>
      <div className="px-6 py-5 border-b border-slate-100 dark:border-zinc-800 bg-slate-50/50 dark:bg-zinc-800/50 flex justify-between items-center">
        <h3 className="text-sm font-black text-slate-800 dark:text-zinc-100 truncate pr-4" title={result.keyword.phrase}>{result.keyword.phrase}</h3>
        {isDone && <CheckCircle2 size={16} className="text-emerald-500 dark:text-emerald-400 shrink-0" />}
      </div>
      <div className="p-6 grid grid-cols-2 gap-4">
        {result.variants.map((v, i) => (
          <div key={v.id} className="space-y-3">
             <div className="aspect-square bg-slate-100 dark:bg-zinc-800 rounded-2xl overflow-hidden relative border border-slate-200/60 dark:border-zinc-700 shadow-inner">
                {v.status === 'generating' && (
                  <GeneratingOverlay variantType={v.variantType} />
                )}
                {v.status === 'completed' && (
                   <img src={v.url} className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-1000" alt="Variant" />
                )}
                {v.status === 'error' && (
                  <div className="absolute inset-0 bg-red-50 dark:bg-red-900/20 flex flex-col items-center justify-center p-4 text-center">
                    <XCircle className="text-red-500 dark:text-red-400 mb-2" size={20} />
                    <p className="text-[8px] font-black text-red-700 dark:text-red-300 uppercase leading-tight line-clamp-2">{v.errorMessage}</p>
                  </div>
                )}
                {v.status === 'pending' && <div className="absolute inset-0 flex items-center justify-center opacity-20 dark:opacity-40 text-slate-400"><ImageIcon size={32} /></div>}
             </div>
             <span className="block text-[9px] font-black text-slate-400 dark:text-zinc-500 uppercase tracking-widest px-1">Variant {i+1}</span>
          </div>
        ))}
      </div>
    </div>
  );
};
