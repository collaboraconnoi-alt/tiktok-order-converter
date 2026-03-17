import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, Copy, Check, Trash2, AlertCircle } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

/**
 * Utility for tailwind class merging
 */
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface TikTokRow {
  'Buyer Username'?: string;
  'Recipient'?: string;
  'City'?: string;
  'Postcode'?: string;
  'Street Name'?: string;
  'Email'?: string;
  'Phone #'?: string;
  'House Name or Number'?: string;
  'Order ID'?: string;
  [key: string]: any;
}

interface ConvertedRow {
  USER: string;
  TRACKING: string;
  RECIPIENT: string;
  LOCATION: string;
  POSTCODE: string;
  ADDRESS: string;
  EMAIL: string;
  PHONE: string;
  NOTES: string;
  PRODUCT_TYPE: string;
  TIKTOK_ORDER: string;
}

const COLUMN_HEADERS = [
  'USER', 'TRACKING', 'RECIPIENT', 'LOCATION', 'POSTCODE', 'ADDRESS', 'EMAIL', 'PHONE', 'NOTES', 'PRODUCT TYPE', 'TIKTOK ORDER'
];

export default function App() {
  const [data, setData] = useState<ConvertedRow[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [copied, setCopied] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const cleanPhone = (phone: string | number | undefined): string => {
    if (!phone) return '';
    let str = String(phone).trim();
    
    // Remove common Italian prefix variations at the start
    // This handles: +39, (+39), ( +39 ), 0039, 39
    // We use a more robust regex that accounts for potential spaces and parentheses
    str = str.replace(/^(\+39|\(\+39\)|0039|39|\(\s*\+39\s*\))\s*/, '');
    
    // If it still starts with +39 or (39) after some weird formatting
    str = str.replace(/^\+39/, '');
    str = str.replace(/^\(39\)/, '');
    
    return str.trim();
  };

  const processFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const bstr = e.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to JSON using headers
        const rawData = XLSX.utils.sheet_to_json(worksheet) as TikTokRow[];
        
        // Also get raw rows to handle "Column AU" if headers fail
        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

        if (rawData.length === 0) {
          setError("The file appears to be empty.");
          return;
        }

        const converted: ConvertedRow[] = rawData.map((row, index) => {
          // 1. Try to find by known keys (case-insensitive)
          const keys = Object.keys(row);
          const zipKey = keys.find(k => 
            k.toLowerCase().includes('zip') || 
            k.toLowerCase().includes('postcode') || 
            k.toLowerCase().includes('post code') ||
            k.toLowerCase().includes('cap') // Italian for ZIP
          );
          
          let zip = zipKey ? row[zipKey] : '';

          // 2. Fallback: If not found, look at raw rows between City and Street Name
          if (!zip && rawRows[0]) {
            const headerRow = rawRows[0] as string[];
            const cityIdx = headerRow.findIndex(h => String(h).toLowerCase().includes('city'));
            const streetIdx = headerRow.findIndex(h => String(h).toLowerCase().includes('street name'));
            
            if (cityIdx !== -1 && streetIdx !== -1 && Math.abs(cityIdx - streetIdx) === 2) {
              // Usually Zip is right between them
              const midIdx = (cityIdx + streetIdx) / 2;
              zip = rawRows[index + 1] ? rawRows[index + 1][midIdx] : '';
            } else if (cityIdx !== -1 && rawRows[index + 1]) {
              // Often it's just the next column
              zip = rawRows[index + 1][cityIdx + 1];
            }
          }

          // 3. Last resort: Column AU (index 46)
          if (!zip) {
            zip = rawRows[index + 1] ? rawRows[index + 1][46] : '';
          }

          return {
            USER: String(row['Buyer Username'] || '').trim(),
            TRACKING: '',
            RECIPIENT: String(row['Recipient'] || '').trim(),
            LOCATION: String(row['City'] || '').trim(),
            POSTCODE: String(zip || '').trim(),
            ADDRESS: String(row['Street Name'] || '').trim(),
            EMAIL: String(row['Email'] || '').trim(),
            PHONE: cleanPhone(row['Phone #']),
            NOTES: String(row['House Name or Number'] || '').trim(),
            PRODUCT_TYPE: '',
            TIKTOK_ORDER: String(row['Order ID'] || '').trim(),
          };
        });

        setData(converted);
        setError(null);
      } catch (err) {
        console.error(err);
        setError("Failed to parse the XLSX file. Please ensure it's a valid TikTok export.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.name.endsWith('.xlsx'))) {
      processFile(file);
    } else {
      setError("Please upload a valid .xlsx file.");
    }
  }, []);

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const copyToClipboard = () => {
    if (data.length === 0) return;

    // Create TSV format for Google Sheets - ONLY rows, no header
    const rows = data.map(row => [
      row.USER, row.TRACKING, row.RECIPIENT, row.LOCATION, row.POSTCODE, 
      row.ADDRESS, row.EMAIL, row.PHONE, row.NOTES, row.PRODUCT_TYPE, row.TIKTOK_ORDER
    ].join('\t'));

    const fullText = rows.join('\n');

    navigator.clipboard.writeText(fullText).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

  const clearData = () => {
    setData([]);
    setError(null);
  };

  return (
    <div className="min-h-screen bg-[#F9FAFB] text-[#111827] font-sans p-4 md:p-8">
      <div className="max-w-7xl mx-auto space-y-8">
        {/* Header */}
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h1 className="text-3xl font-bold tracking-tight text-[#111827]">TikTok Shipping Converter</h1>
            <p className="text-[#6B7280] mt-1">Transform your TikTok exports for Google Sheets instantly.</p>
          </div>
          {data.length > 0 && (
            <div className="flex items-center gap-3">
              <button
                onClick={clearData}
                className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-[#EF4444] hover:bg-[#FEF2F2] rounded-lg transition-colors"
              >
                <Trash2 size={18} />
                Clear
              </button>
              <button
                onClick={copyToClipboard}
                className={cn(
                  "flex items-center gap-2 px-6 py-2.5 rounded-lg font-semibold text-sm transition-all shadow-sm",
                  copied 
                    ? "bg-[#10B981] text-white" 
                    : "bg-[#111827] text-white hover:bg-[#1F2937] active:scale-95"
                )}
              >
                {copied ? <Check size={18} /> : <Copy size={18} />}
                {copied ? 'Copied!' : 'Copy for Sheets'}
              </button>
            </div>
          )}
        </header>

        {/* Error Message */}
        <AnimatePresence>
          {error && (
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="bg-[#FEF2F2] border border-[#FEE2E2] text-[#B91C1C] px-4 py-3 rounded-lg flex items-center gap-3"
            >
              <AlertCircle size={20} />
              <span className="text-sm font-medium">{error}</span>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Upload Zone */}
        {data.length === 0 ? (
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className={cn(
              "relative group border-2 border-dashed rounded-2xl p-12 transition-all duration-300 flex flex-col items-center justify-center text-center space-y-4",
              isDragging 
                ? "border-[#3B82F6] bg-[#EFF6FF]" 
                : "border-[#E5E7EB] bg-white hover:border-[#D1D5DB]"
            )}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={onDrop}
          >
            <div className="w-16 h-16 bg-[#F3F4F6] rounded-full flex items-center justify-center text-[#9CA3AF] group-hover:text-[#3B82F6] group-hover:bg-[#DBEAFE] transition-colors">
              <Upload size={32} />
            </div>
            <div className="space-y-1">
              <p className="text-lg font-semibold text-[#111827]">
                Click to upload or drag and drop
              </p>
              <p className="text-sm text-[#6B7280]">
                TikTok Shipping Export (.xlsx)
              </p>
            </div>
            <input
              type="file"
              accept=".xlsx"
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              onChange={onFileChange}
            />
          </motion.div>
        ) : (
          /* Data Table */
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            className="bg-white border border-[#E5E7EB] rounded-xl shadow-sm overflow-hidden"
          >
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[1200px]">
                <thead>
                  <tr className="bg-[#F9FAFB] border-bottom border-[#E5E7EB]">
                    {COLUMN_HEADERS.map((header) => (
                      <th 
                        key={header} 
                        className="px-4 py-3 text-[11px] font-bold text-[#6B7280] uppercase tracking-wider border-r border-[#E5E7EB] last:border-r-0"
                      >
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB]">
                  {data.map((row, idx) => (
                    <tr key={idx} className="hover:bg-[#F9FAFB] transition-colors group">
                      <td className="px-4 py-3 text-sm font-medium text-[#111827] border-r border-[#E5E7EB]">{row.USER}</td>
                      <td className="px-4 py-3 text-sm text-[#9CA3AF] italic border-r border-[#E5E7EB] bg-[#FDFDFD]">Empty</td>
                      <td className="px-4 py-3 text-sm text-[#374151] border-r border-[#E5E7EB]">{row.RECIPIENT}</td>
                      <td className="px-4 py-3 text-sm text-[#374151] border-r border-[#E5E7EB]">{row.LOCATION}</td>
                      <td className="px-4 py-3 text-sm font-mono text-[#374151] border-r border-[#E5E7EB]">{row.POSTCODE}</td>
                      <td className="px-4 py-3 text-sm text-[#374151] border-r border-[#E5E7EB]">{row.ADDRESS}</td>
                      <td className="px-4 py-3 text-sm text-[#374151] border-r border-[#E5E7EB]">{row.EMAIL}</td>
                      <td className="px-4 py-3 text-sm font-mono text-[#374151] border-r border-[#E5E7EB]">{row.PHONE}</td>
                      <td className="px-4 py-3 text-sm text-[#374151] border-r border-[#E5E7EB]">{row.NOTES}</td>
                      <td className="px-4 py-3 text-sm text-[#9CA3AF] italic border-r border-[#E5E7EB] bg-[#FDFDFD]">Empty</td>
                      <td className="px-4 py-3 text-sm font-mono text-[#6B7280]">{row.TIKTOK_ORDER}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="px-4 py-3 bg-[#F9FAFB] border-t border-[#E5E7EB] flex items-center justify-between">
              <span className="text-xs font-medium text-[#6B7280]">
                Showing {data.length} {data.length === 1 ? 'row' : 'rows'}
              </span>
              <div className="flex items-center gap-2 text-[#9CA3AF]">
                <FileSpreadsheet size={14} />
                <span className="text-[10px] uppercase tracking-widest font-bold">XLSX Processed</span>
              </div>
            </div>
          </motion.div>
        )}

        {/* Instructions */}
        <footer className="pt-8 border-t border-[#E5E7EB]">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
            <div className="space-y-2">
              <h3 className="text-sm font-bold text-[#374151] uppercase tracking-wider">1. Upload</h3>
              <p className="text-sm text-[#6B7280]">Export your shipping data from TikTok Shop as an XLSX file and drop it here.</p>
            </div>
            <div className="space-y-2">
              <h3 className="text-sm font-bold text-[#374151] uppercase tracking-wider">2. Review</h3>
              <p className="text-sm text-[#6B7280]">Columns are automatically mapped and phone numbers cleaned (removing +39 prefix).</p>
            </div>
            <div className="space-y-2">
              <h3 className="text-sm font-bold text-[#374151] uppercase tracking-wider">3. Paste</h3>
              <p className="text-sm text-[#6B7280]">Click "Copy for Sheets" and paste directly into your Google Sheets file.</p>
            </div>
          </div>
        </footer>
      </div>
    </div>
  );
}
