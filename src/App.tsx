/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { useLiveQuery } from 'dexie-react-hooks';
import { 
  FileText, 
  Receipt, 
  CheckCircle2, 
  FolderOpen, 
  RefreshCw, 
  Plus, 
  Search,
  LayoutDashboard,
  Settings,
  AlertCircle,
  ArrowRightLeft,
  Briefcase,
  Download,
  Trash2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { db, type Statement, type Invoice, type Project } from './lib/db';
import { parseFinancePDF, matchInvoices } from './lib/ai';
import { OutlookService } from './lib/outlook';
import { useFileSystem } from './hooks/useFileSystem';
import { useEffect } from 'react';
import JSZip from 'jszip';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export default function App() {
  const [activeTab, setActiveTab] = useState<'statements' | 'invoices' | 'match' | 'projects'>('statements');
  const [isProcessing, setIsProcessing] = useState(false);
  const [syncStatus, setSyncStatus] = useState<string>('');
  const [selectedProjectId, setSelectedProjectId] = useState<string>('');
  const { directoryHandle, requestPermission, saveFile } = useFileSystem();

  const statements = useLiveQuery(() => db.statements.toArray()) || [];
  const invoices = useLiveQuery(() => db.invoices.toArray()) || [];
  const projects = useLiveQuery(() => db.projects.toArray()) || [];
  const outlookTokens = useLiveQuery(() => db.settings.get('outlook_tokens'));

  useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'OUTLOOK_AUTH_SUCCESS') {
        db.settings.put({ key: 'outlook_tokens', value: event.data.tokens });
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const [syncStartDate, setSyncStartDate] = useState<string>('2026-01-01');

  useEffect(() => {
    const loadSyncSettings = async () => {
      const start = await db.settings.get('sync_start_date');
      if (start) setSyncStartDate(start.value);
    };
    loadSyncSettings();
  }, []);

  const updateSyncStartDate = async (date: string) => {
    setSyncStartDate(date);
    await db.settings.put({ key: 'sync_start_date', value: date });
  };

  const clearLastSync = async () => {
    if (confirm("Reset sync history? Next sync will start from your configured 'Sync Start Date'.")) {
      await db.settings.delete('last_sync_date');
      alert("Sync history cleared.");
    }
  };

  const connectOutlook = async () => {
    const res = await fetch('/api/auth/url');
    const { url } = await res.json();
    window.open(url, 'outlook_auth', 'width=600,height=600');
  };

  const resetOutlookConnection = async () => {
    if (confirm("Are you sure you want to disconnect and reset Outlook tokens?")) {
      await db.settings.delete('outlook_tokens');
      alert("Outlook connection reset successfully.");
    }
  };

  const runDailySync = async () => {
    if (!directoryHandle) return alert("Please set a storage directory first using the button in the sidebar.");
    setIsProcessing(true);
    setSyncStatus('Starting sync...');
    try {
      const lastSync = await db.settings.get('last_sync_date');
      let startBase = syncStartDate;
      if (!startBase || isNaN(new Date(startBase).getTime())) {
        startBase = '2026-01-01';
      }
      const afterDate = lastSync?.value || new Date(startBase).toISOString();
      console.log('Syncing emails received after:', afterDate);
      
      const processEmails = async (query: string) => {
        setSyncStatus(`Searching for "${query}"...`);
        const emails = await OutlookService.fetchEmails(query, afterDate);
        setSyncStatus(`Found ${emails.length} emails for "${query}"`);
        
        for (const email of emails) {
          const pdfAttachment = email.attachments?.find((a: any) => a.contentType === 'application/pdf');
          if (pdfAttachment) {
            setSyncStatus(`Processing: ${email.subject}`);
            const base64 = await OutlookService.getAttachment(email.id, pdfAttachment.id);
            const parsed = await parseFinancePDF(base64, email.subject);
            
            const blob = await (await fetch(`data:application/pdf;base64,${base64}`)).blob();
            const localPath = await saveFile(directoryHandle, blob, parsed.vendor, parsed.date, parsed.type === 'statement' ? 'Statement' : 'Invoice');

            if (parsed.type === 'statement') {
              await db.statements.add({
                vendor: parsed.vendor,
                amount: parsed.amount,
                date: parsed.date,
                issueDate: parsed.issueDate,
                statementNumber: parsed.statementNumber,
                status: 'pending',
                localFilePath: localPath,
                outlookId: email.id,
                subject: email.subject,
                projectId: selectedProjectId,
                invoiceNumbers: parsed.invoiceNumbers || [],
                items: parsed.items || []
              });
            } else {
              await db.invoices.add({
                vendor: parsed.vendor,
                amount: parsed.amount,
                date: parsed.date,
                invoiceNumber: parsed.invoiceNumber || 'N/A',
                localFilePath: localPath,
                outlookId: email.id,
                isProcessed: false,
                projectId: selectedProjectId
              });
            }
          }
        }
      };

      await processEmails('Statement');
      await processEmails('Invoice');
      await processEmails('Tax Invoice');
      await processEmails('Account Statement');
      await processEmails('Bill');
      await processEmails('Receipt');
      
      setSyncStatus('Matching documents...');
      await matchInvoices();
      await db.settings.put({ key: 'last_sync_date', value: new Date().toISOString() });
      setSyncStatus('Sync complete!');
      setTimeout(() => setSyncStatus(''), 3000);
    } catch (err) {
      console.error(err);
      setSyncStatus('Sync failed');
      alert("Sync failed: " + (err as Error).message);
    } finally {
      setIsProcessing(false);
    }
  };

  const exportProject = async (project: Project) => {
    if (!directoryHandle) return;
    setIsProcessing(true);
    try {
      const zip = new JSZip();
      const projectStatements = await db.statements.where('projectId').equals(project.id).toArray();
      const projectInvoices = await db.invoices.where('projectId').equals(project.id).toArray();

      const addFileToZip = async (path: string, prefix: string) => {
        try {
          // Navigate to file in local storage
          const parts = path.split('/');
          let currentHandle: any = directoryHandle;
          for (let i = 0; i < parts.length - 1; i++) {
            currentHandle = await currentHandle.getDirectoryHandle(parts[i]);
          }
          const fileHandle = await currentHandle.getFileHandle(parts[parts.length - 1]);
          const file = await fileHandle.getFile();
          const fileName = `${project.id}_${prefix}_${parts[parts.length - 1]}`;
          zip.file(fileName, file);
        } catch (e) {
          console.error("Failed to add file to zip", path, e);
        }
      };

      for (const s of projectStatements) {
        if (s.localFilePath) await addFileToZip(s.localFilePath, 'STMT');
      }
      for (const inv of projectInvoices) {
        if (inv.localFilePath) await addFileToZip(inv.localFilePath, 'INV');
      }

      const content = await zip.generateAsync({ type: 'blob' });
      const saveHandle = await (window as any).showSaveFilePicker({
        suggestedName: `Project_${project.id}_Export.zip`,
        types: [{ description: 'ZIP file', accept: { 'application/zip': ['.zip'] } }]
      });
      const writable = await saveHandle.createWritable();
      await writable.write(content);
      await writable.close();
    } catch (err) {
      console.error(err);
      alert("Export failed");
    } finally {
      setIsProcessing(false);
    }
  };

  const createProject = async () => {
    const id = prompt("Enter Project Number/ID:");
    const name = prompt("Enter Project Name:");
    if (id && name) {
      await db.projects.add({ id, name, createdAt: new Date().toISOString() });
      setSelectedProjectId(id);
    }
  };

  const toggleProcessed = async (id: number, current: boolean) => {
    await db.invoices.update(id, { isProcessed: !current });
  };

  const stats = useMemo(() => ({
    pending: statements.filter(s => s.status === 'pending').length,
    matched: statements.filter(s => s.status === 'matched').length,
    totalAmount: statements.reduce((acc, s) => acc + s.amount, 0)
  }), [statements]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files?.length) return;
    if (!directoryHandle) {
      alert("Please select a local storage directory first!");
      return;
    }

    setIsProcessing(true);
    try {
      for (const file of Array.from(e.target.files)) {
        const reader = new FileReader();
        const base64Promise = new Promise<string>((resolve) => {
          reader.onload = () => resolve((reader.result as string).split(',')[1]);
          reader.readAsDataURL(file);
        });

        const base64 = await base64Promise;
        const parsed = await parseFinancePDF(base64);
        
        const blob = new Blob([file], { type: 'application/pdf' });
        const localPath = await saveFile(
          directoryHandle, 
          blob, 
          parsed.vendor, 
          parsed.date, 
          parsed.type === 'statement' ? 'Statement' : 'Invoice'
        );

        if (parsed.type === 'statement') {
          await db.statements.add({
            vendor: parsed.vendor,
            amount: parsed.amount,
            date: parsed.date,
            issueDate: parsed.issueDate,
            statementNumber: parsed.statementNumber,
            status: 'pending',
            localFilePath: localPath,
            invoiceNumbers: parsed.invoiceNumbers || [],
            items: parsed.items || []
          });
        } else {
          await db.invoices.add({
            vendor: parsed.vendor,
            amount: parsed.amount,
            date: parsed.date,
            invoiceNumber: parsed.invoiceNumber || 'N/A',
            localFilePath: localPath,
            isProcessed: false
          });
        }
      }
      await matchInvoices();
    } catch (err) {
      console.error(err);
      alert("Error processing files");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-slate-900 font-sans selection:bg-indigo-100">
      {/* Sidebar */}
      <aside className="fixed left-0 top-0 h-full w-64 bg-white/80 backdrop-blur-xl border-r border-slate-200 z-50 p-6 flex flex-col">
        <div className="flex items-center gap-3 mb-10 px-2">
          <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center shadow-lg shadow-indigo-200">
            <ArrowRightLeft className="text-white w-6 h-6" />
          </div>
          <h1 className="font-bold text-lg tracking-tight">FinMatch AI</h1>
        </div>

        <nav className="space-y-1 flex-1">
          <NavItem 
            icon={<LayoutDashboard size={20} />} 
            label="Dashboard" 
            active={activeTab === 'statements'} 
            onClick={() => setActiveTab('statements')} 
          />
          <NavItem 
            icon={<Receipt size={20} />} 
            label="Invoices" 
            active={activeTab === 'invoices'} 
            onClick={() => setActiveTab('invoices')} 
          />
          <NavItem 
            icon={<CheckCircle2 size={20} />} 
            label="Match Center" 
            active={activeTab === 'match'} 
            onClick={() => setActiveTab('match')} 
          />
          <NavItem 
            icon={<Briefcase size={20} />} 
            label="Projects" 
            active={activeTab === 'projects'} 
            onClick={() => setActiveTab('projects')} 
          />
        </nav>

        <div className="mt-auto space-y-2 pt-6 border-t border-slate-100">
          {outlookTokens ? (
            <div className="space-y-2">
              <div className="flex items-center gap-2 px-4 py-2 mb-2 bg-emerald-50 text-emerald-700 rounded-lg border border-emerald-100">
                <CheckCircle2 size={16} />
                <span className="text-xs font-bold uppercase tracking-wider">Outlook Connected</span>
              </div>
              <button 
                onClick={resetOutlookConnection}
                className="w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 bg-rose-50 text-rose-700 hover:bg-rose-100"
              >
                <RefreshCw size={20} className="rotate-180" />
                <span className="text-sm font-medium">Reset Connection</span>
              </button>
              <button 
                onClick={clearLastSync}
                className="w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 bg-amber-50 text-amber-700 hover:bg-amber-100"
              >
                <RefreshCw size={20} />
                <span className="text-sm font-medium">Clear Sync History</span>
              </button>
            </div>
          ) : (
            <button 
              onClick={connectOutlook}
              className="w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 bg-indigo-50 text-indigo-700 hover:bg-indigo-100"
            >
              <RefreshCw size={20} />
              <span className="text-sm font-medium">Connect Outlook</span>
            </button>
          )}
          <button 
            onClick={async () => {
              if (!('showDirectoryPicker' in window)) {
                alert("Your browser doesn't support the File System Access API, or it's blocked in this view. Please try opening the app in a new tab using the icon at the top right.");
                return;
              }
              try {
                const handle = await requestPermission();
                if (!handle) {
                  alert("Directory selection was cancelled or failed. If you are in the preview pane, try opening the app in a new tab.");
                }
              } catch (e) {
                alert("Error accessing file system: " + (e as Error).message);
              }
            }}
            className={cn(
              "w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200",
              directoryHandle ? "bg-emerald-50 text-emerald-700" : "bg-slate-50 text-slate-600 hover:bg-slate-100"
            )}
          >
            <FolderOpen size={20} />
            <span className="text-sm font-medium">
              {directoryHandle ? 'Storage Ready' : 'Set Storage'}
            </span>
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="pl-64 min-h-screen">
        {/* Header */}
        <header className="sticky top-0 bg-white/60 backdrop-blur-md border-b border-slate-200 z-40 px-8 py-4 flex items-center justify-between">
          <div className="flex gap-8">
            <MetricCard label="Pending" value={stats.pending} color="text-amber-600" />
            <MetricCard label="Matched" value={stats.matched} color="text-emerald-600" />
            <MetricCard label="Total Volume" value={`$${stats.totalAmount.toLocaleString()}`} color="text-indigo-600" />
          </div>

          <div className="flex items-center gap-4">
            {!outlookTokens ? (
              <div className="flex items-center gap-2 text-amber-600 bg-amber-50 px-3 py-1.5 rounded-lg border border-amber-100 animate-pulse">
                <AlertCircle size={14} />
                <span className="text-xs font-medium">Outlook Disconnected</span>
              </div>
            ) : (
              <div className="flex items-center gap-2 text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-100">
                <CheckCircle2 size={14} />
                <span className="text-xs font-medium">Outlook Connected</span>
              </div>
            )}
            {outlookTokens && !directoryHandle && (
              <div className="flex items-center gap-2 text-rose-600 bg-rose-50 px-3 py-1.5 rounded-lg border border-rose-100">
                <FolderOpen size={14} />
                <span className="text-xs font-medium">Set Storage Directory</span>
              </div>
            )}
            <div className="flex flex-col items-end">
              <span className="text-[10px] uppercase font-bold text-slate-400 mb-0.5">Sync From</span>
              <input 
                type="date" 
                value={syncStartDate}
                onChange={(e) => updateSyncStartDate(e.target.value)}
                className="text-xs bg-white border border-slate-200 rounded-lg px-2 py-1.5 outline-none focus:ring-2 focus:ring-indigo-500/20"
              />
            </div>
            <select 
              value={selectedProjectId} 
              onChange={(e) => setSelectedProjectId(e.target.value)}
              className="bg-white border border-slate-200 text-sm rounded-xl px-3 py-2.5 focus:ring-2 focus:ring-indigo-500/20 outline-none"
            >
              <option value="">No Project Assigned</option>
              {projects.map(p => (
                <option key={p.id} value={p.id}>{p.id} - {p.name}</option>
              ))}
            </select>
            <button 
              onClick={runDailySync}
              disabled={isProcessing}
              className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2.5 rounded-xl font-medium text-sm flex items-center gap-2 transition-all shadow-md shadow-indigo-100 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <RefreshCw size={18} className={cn(isProcessing && "animate-spin")} />
              {isProcessing ? (syncStatus || 'Syncing...') : 'Run Daily Sync'}
            </button>
            <label className="cursor-pointer bg-white border border-slate-200 text-slate-700 px-5 py-2.5 rounded-xl font-medium text-sm flex items-center gap-2 transition-all hover:bg-slate-50">
              <Plus size={18} />
              Manual Import
              <input type="file" multiple accept=".pdf" className="hidden" onChange={handleFileUpload} disabled={isProcessing} />
            </label>
          </div>
        </header>

        <div className="p-8">
          <AnimatePresence mode="wait">
            <motion.div
              key={activeTab}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.2 }}
            >
              {activeTab === 'statements' && (
                <DataTable 
                  title="Statements" 
                  data={statements} 
                  columns={[
                    { key: 'vendor', label: 'Vendor' },
                    { key: 'statementNumber', label: 'Statement No.' },
                    { key: 'issueDate', label: 'Issue Date' },
                    { key: 'date', label: 'Period' },
                    { 
                      key: 'items', 
                      label: 'Invoices & Amounts', 
                      render: (items) => (
                        <div className="flex flex-col gap-1">
                          {items?.map((item: any, i: number) => (
                            <div key={i} className="flex justify-between gap-4 text-[10px] border-b border-slate-50 pb-0.5 last:border-0">
                              <span className="font-mono text-slate-500">{item.invoiceNumber}</span>
                              <span className="font-bold text-slate-700">${item.amount.toLocaleString()}</span>
                            </div>
                          ))}
                          {(!items || items.length === 0) && <span className="text-slate-300 italic">No breakdown</span>}
                        </div>
                      )
                    },
                    { key: 'amount', label: 'Total', format: (v) => `$${v.toLocaleString()}` },
                    { key: 'status', label: 'Status', render: (s) => <StatusBadge status={s} /> }
                  ]}
                />
              )}

              {activeTab === 'invoices' && (
                <DataTable 
                  title="Invoice Library" 
                  data={invoices} 
                  columns={[
                    { key: 'vendor', label: 'Vendor' },
                    { key: 'invoiceNumber', label: 'Inv #' },
                    { key: 'date', label: 'Date' },
                    { key: 'amount', label: 'Amount', format: (v) => `$${v.toLocaleString()}` },
                    { 
                      key: 'matchedStatementId', 
                      label: 'Statement', 
                      render: (id) => id ? (
                        <span className="text-xs font-bold text-indigo-600 bg-indigo-50 px-2 py-1 rounded-lg">
                          Linked to #{id}
                        </span>
                      ) : <span className="text-xs text-slate-400">Unlinked</span>
                    },
                    {
                      key: 'isProcessed',
                      label: 'Action',
                      render: (val, row) => (
                        <button 
                          onClick={() => toggleProcessed(row.id, val)}
                          className={cn(
                            "p-2 rounded-lg transition-all",
                            val ? "bg-emerald-100 text-emerald-700" : "bg-slate-100 text-slate-400 hover:bg-slate-200"
                          )}
                        >
                          <CheckCircle2 size={18} />
                        </button>
                      )
                    }
                  ]}
                />
              )}

              {activeTab === 'projects' && (
                <div className="space-y-6">
                  <div className="flex items-center justify-between">
                    <h2 className="text-2xl font-bold tracking-tight">Project Management</h2>
                    <button 
                      onClick={createProject}
                      className="bg-indigo-600 text-white px-4 py-2 rounded-xl text-sm font-bold flex items-center gap-2"
                    >
                      <Plus size={18} /> New Project
                    </button>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                    {projects.map(p => (
                      <div key={p.id} className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm hover:shadow-md transition-all group">
                        <div className="flex justify-between items-start mb-4">
                          <div>
                            <p className="text-[10px] font-bold text-indigo-600 uppercase tracking-widest">{p.id}</p>
                            <h3 className="text-lg font-bold text-slate-900">{p.name}</h3>
                          </div>
                          <Briefcase className="text-slate-200 group-hover:text-indigo-100 transition-colors" size={32} />
                        </div>
                        <div className="flex gap-2 mt-6">
                          <button 
                            onClick={() => exportProject(p)}
                            className="flex-1 bg-slate-50 hover:bg-indigo-50 text-slate-600 hover:text-indigo-700 py-2 rounded-xl text-xs font-bold flex items-center justify-center gap-2 transition-all"
                          >
                            <Download size={14} /> Export Project
                          </button>
                          <button 
                            onClick={() => db.projects.delete(p.id)}
                            className="p-2 text-slate-300 hover:text-rose-500 transition-colors"
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </motion.div>
          </AnimatePresence>
        </div>
      </main>

      {isProcessing && (
        <div className="fixed inset-0 bg-white/40 backdrop-blur-sm z-[100] flex items-center justify-center">
          <div className="bg-white p-8 rounded-3xl shadow-2xl border border-slate-100 flex flex-col items-center gap-4">
            <RefreshCw className="animate-spin text-indigo-600" size={40} />
            <p className="font-bold text-lg">AI is parsing your documents...</p>
            <p className="text-slate-500 text-sm">This stays 100% on your device.</p>
          </div>
        </div>
      )}
    </div>
  );
}

function NavItem({ icon, label, active, onClick }: { icon: React.ReactNode, label: string, active?: boolean, onClick: () => void }) {
  return (
    <button 
      onClick={onClick}
      className={cn(
        "w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 group",
        active ? "bg-indigo-50 text-indigo-700" : "text-slate-500 hover:bg-slate-50 hover:text-slate-900"
      )}
    >
      <span className={cn("transition-transform duration-200", active && "scale-110")}>{icon}</span>
      <span className="text-sm font-semibold">{label}</span>
      {active && <motion.div layoutId="activeNav" className="ml-auto w-1.5 h-1.5 rounded-full bg-indigo-600" />}
    </button>
  );
}

function MetricCard({ label, value, color }: { label: string, value: string | number, color: string }) {
  return (
    <div>
      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-0.5">{label}</p>
      <p className={cn("text-xl font-black tracking-tight", color)}>{value}</p>
    </div>
  );
}

function StatusBadge({ status }: { status: string }) {
  const styles = {
    pending: "bg-amber-100 text-amber-700 border-amber-200",
    matched: "bg-emerald-100 text-emerald-700 border-emerald-200",
    error: "bg-rose-100 text-rose-700 border-rose-200"
  };
  return (
    <span className={cn("px-2.5 py-1 rounded-full text-[10px] font-bold uppercase border", styles[status as keyof typeof styles])}>
      {status}
    </span>
  );
}

function DataTable({ title, data, columns }: { title: string, data: any[], columns: any[] }) {
  return (
    <div className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden">
      <div className="px-8 py-6 border-b border-slate-100 flex items-center justify-between">
        <h2 className="text-xl font-bold tracking-tight">{title}</h2>
        <div className="relative">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
          <input 
            type="text" 
            placeholder="Search..." 
            className="pl-10 pr-4 py-2 bg-slate-50 border-none rounded-xl text-sm focus:ring-2 focus:ring-indigo-500/20 transition-all w-64"
          />
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-slate-50/50">
              {columns.map(col => (
                <th key={col.key} className="px-8 py-4 text-[11px] font-bold text-slate-400 uppercase tracking-widest">
                  {col.label}
                </th>
              ))}
              <th className="px-8 py-4"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {data.map((row, i) => (
              <tr key={i} className="hover:bg-slate-50/50 transition-colors group">
                {columns.map(col => (
                  <td key={col.key} className="px-8 py-4 text-sm font-medium text-slate-600">
                    {col.render ? col.render(row[col.key], row) : (col.format ? col.format(row[col.key]) : row[col.key])}
                  </td>
                ))}
                <td className="px-8 py-4 text-right">
                  <button className="p-2 text-slate-400 hover:text-indigo-600 opacity-0 group-hover:opacity-100 transition-all">
                    <FileText size={18} />
                  </button>
                </td>
              </tr>
            ))}
            {data.length === 0 && (
              <tr>
                <td colSpan={columns.length + 1} className="px-8 py-20 text-center">
                  <div className="flex flex-col items-center gap-3 text-slate-400">
                    <AlertCircle size={40} strokeWidth={1.5} />
                    <p className="text-sm font-medium">No data found. Import some PDFs to get started.</p>
                  </div>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
