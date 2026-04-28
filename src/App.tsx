/* eslint-disable @typescript-eslint/no-unused-vars */
import React, { useState, useEffect, useMemo, useCallback } from 'react';
import Papa from 'papaparse';
import {
  ComposedChart, Line, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend,
  ResponsiveContainer, PieChart, Pie, Cell
} from 'recharts';
import {
  Users, UserCheck, UserMinus, BarChart2, PieChart as PieChartIcon, TrendingUp, TrendingDown,
  CheckCircle, XCircle, AlertCircle, Info, Table as TableIcon, Loader2, LayoutDashboard,
  ClipboardList, CalendarCheck, ChevronDown, LogOut, ChevronLeft, ChevronRight,
  RefreshCw, Edit3, Calendar, Filter
} from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// ============================================
// ===== 案件設定（ここだけ変更してください） =====
// ============================================
const CONFIG = {
  TITLE: 'ACQUA LINEダッシュボード',
  CSV_URL: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTwVpKVhJ3eunj2VX8Ibye67mm4-QQ0r7NvdboXGaXhKBsvIdlVQegmz6YfhODRygNmall2XX3i_4Pz/pub?gid=1922873880&single=true&output=csv',
  SHEET_URL: 'https://docs.google.com/spreadsheets/d/168I4mc32B_Sr41eexJPnFFZW4BkKobeKKH8NQsNzev0/edit?usp=sharing',
  PROXY_URL: 'https://line-dashboard-proxy.raspy-wood-9b0d.workers.dev',
  GOOGLE_CLIENT_ID: '813216912152-hf6cden86ijta1qjc67uvscdlhmi85sl.apps.googleusercontent.com',
  SHEET_NAME: 'データ反映シート',
  // ※ AI Studio → CSV_URL使用 / GitHub Pages → 自動でプロキシ経由（手動切替不要）
};
// ============================================

const COLORS = { 
  primary: "#0067b8", secondary: "#00A4EF", success: "#107c10", 
  warning: "#ffb900", danger: "#d13438", info: "#0078d4", 
  muted: "#666666", accent: "#9bf00b",
  positive: "#0067b8",  
  negative: "#d13438",  
};
const PIE_COLORS = [
  "#0067b8", "#107c10", "#00A4EF", "#ffb900", "#d13438", 
  "#0078d4", "#881798", "#00b294", "#e3008c", "#ff8c00", "#00188f"
];

const getSheetId = (url: string) => {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : url;
};
const parseDate = (dateStr: any) => {
  if (!dateStr) return null;
  const s = String(dateStr).trim();
  if (!s) return null;
  const match = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (match) {
    const d = new Date(parseInt(match[1],10), parseInt(match[2],10)-1, parseInt(match[3],10));
    return isNaN(d.getTime()) ? null : d;
  }
  return isNaN(new Date(dateStr).getTime()) ? null : new Date(dateStr);
};
const formatMonth = (d: Date | null) => d ? `${d.getFullYear()}年${String(d.getMonth()+1).padStart(2,'0')}月` : null;
const formatWeek = (d: Date | null) => {
  if (!d) return null;
  const s = new Date(d.getFullYear(),0,1);
  const w = Math.ceil((((d.getTime()-s.getTime())/86400000)+s.getDay()+1)/7);
  return `${d.getFullYear()}年 W${String(w).padStart(2,'0')}`;
};
const formatDay = (d: Date | null) => d ? `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}` : null;
const hasTag = (val: any) => { if (!val) return false; const s = String(val).trim(); return s !== '' && s !== '0'; };
const isTrue = (val: any) => { if (!val) return false; const s = String(val).trim(); return s==='1'||s==='１'||s.toLowerCase()==='true'; };

const IconComp = ({ name, size=18, className="" }: { name: string, size?: number, className?: string }) => {
  const m: Record<string, React.ElementType> = { 
    'users':Users,'user-check':UserCheck,'user-minus':UserMinus,
    'bar-chart-2':BarChart2,'pie-chart':PieChartIcon,'trending-up':TrendingUp,
    'trending-down':TrendingDown,'check-circle':CheckCircle,'x-circle':XCircle,
    'alert-circle':AlertCircle,'info':Info,'table':TableIcon,'loader-2':Loader2,
    'layout-dashboard':LayoutDashboard,'clipboard-list':ClipboardList,
    'calendar-check':CalendarCheck,'chevron-down':ChevronDown,'log-out':LogOut,
    'chevron-left':ChevronLeft,'chevron-right':ChevronRight,
    'refresh-cw':RefreshCw,'edit-3':Edit3,'calendar':Calendar,'filter':Filter 
  };
  const I = m[name]; return I ? <I size={size} className={className}/> : null;
};

const InfoTooltip = ({ text }: { text: string }) => (
  <div className="relative group inline-flex items-center ml-1.5 z-[100]">
    <Info size={14} className="text-[#666] cursor-help hover:text-[#0067b8] transition-colors"/>
    <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 hidden group-hover:block w-max max-w-[320px] bg-[#1a1a1a] text-white text-[12px] p-3 rounded-lg shadow-xl whitespace-pre-wrap leading-relaxed pointer-events-none opacity-0 group-hover:opacity-100 transition-opacity z-[100]">
      {text}
      <div className="absolute top-full left-1/2 -translate-x-1/2 border-[5px] border-transparent border-t-[#1a1a1a]"/>
    </div>
  </div>
);

const KPICard = ({ title, value, unit, icon, info, subText, change, changeLabel, isEditMode }: any) => {
  const [editedValue, setEditedValue] = useState(value);
  const [isEditing, setIsEditing] = useState(false);

  useEffect(() => { setEditedValue(value); }, [value]);

  return (
    <div className="card p-5 card-hover flex flex-col justify-between min-h-[120px]">
      <div className="flex justify-between items-start mb-3">
        <div className="p-2 rounded-lg bg-[#f2f2f2]">
          <IconComp name={icon} size={20} className="text-[#0067b8]"/>
        </div>
      </div>
      <div>
        <h3 className="text-[#666] text-[11px] font-semibold tracking-wide uppercase mb-1 flex items-center">
          {title}{info && <InfoTooltip text={info}/>}
        </h3>
        <div className="flex items-baseline gap-1.5">
          {isEditMode && isEditing ? (
            <input 
              autoFocus 
              className="text-[32px] font-bold text-[#000] tracking-tight leading-none bg-gray-100 w-full outline-[#0067b8] px-1 rounded"
              value={editedValue} 
              onChange={e => setEditedValue(e.target.value)} 
              onBlur={() => setIsEditing(false)}
            />
          ) : (
            <span 
              className={cn("text-[32px] font-bold text-[#000] tracking-tight leading-none", isEditMode && "cursor-pointer border-b border-dashed border-gray-400")}
              onClick={() => isEditMode && setIsEditing(true)}
            >
              {typeof editedValue==='number' ? editedValue.toLocaleString() : editedValue}
            </span>
          )}
          <span className="text-[#666] text-xs font-semibold">{unit}</span>
        </div>
        {change!=null && (
          <div className="flex items-center gap-1 mt-1.5">
            <IconComp name={change>=0?'trending-up':'trending-down'} size={12}
              className={change>=0?'text-[#0067b8]':'text-[#d13438]'}/>
            <span className={`text-[11px] font-bold ${change>=0?'text-[#0067b8]':'text-[#d13438]'}`}>
              {change>=0?'+':''}{change}%
            </span>
            {changeLabel && <span className="text-[10px] text-[#666] ml-0.5">{changeLabel}</span>}
          </div>
        )}
        {subText && <p className="text-[11px] text-[#666] mt-1">{subText}</p>}
      </div>
    </div>
  );
}

const CustomTooltip = ({ active, payload, label }: any) => {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-white p-4 rounded-lg shadow-xl border border-[#f2f2f2] text-xs z-[1000]">
      <p className="font-semibold text-[#000] mb-2 text-sm">{label}</p>
      {payload.map((e:any,i:number) => (
        <div key={i} className="flex items-center gap-2 mb-1">
          <div className="w-2 h-2 rounded-full" style={{backgroundColor:e.color}}/>
          <span className="text-[#666]">{e.name}:</span>
          <span className="font-bold text-[#000]">{e.value?.toLocaleString()||0}{e.name.includes('率')?'%':''}</span>
        </div>
      ))}
    </div>
  );
};

const isDeployed = typeof window !== 'undefined' && window.location.hostname.includes('github.io');

function fetchViaCSV(csvUrl: string): Promise<any[]> {
  return new Promise((resolve, reject) => {
    Papa.parse(csvUrl, { 
      download:true, header:true, skipEmptyLines:true,
      transformHeader:(h)=>h.trim(),
      complete:(r)=>resolve(r.data), error:(e)=>reject(e) 
    });
  });
}

async function fetchViaProxy() {
  const token = localStorage.getItem('google_id_token');
  const res = await fetch(`${CONFIG.PROXY_URL}/sheets`, {
    method:'POST', headers:{'Content-Type':'application/json','Authorization':`Bearer ${token}`},
    body: JSON.stringify({sheetId:getSheetId(CONFIG.SHEET_URL),sheetName:CONFIG.SHEET_NAME})
  });
  if (res.status===403) throw new Error('アクセス権がありません。スプレッドシートの共有設定を管理者に確認してください。');
  const json = await res.json();
  return json.rows.map((row: any) => { 
    const o: Record<string,any>={}; 
    json.headers.forEach((h: string,i: number)=>{o[h]=row[i]||''}); 
    return o; 
  });
}

export default function App() {
  const [data, setData] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [isAuthenticated, setIsAuthenticated] = useState(!isDeployed);
  
  const [timeState, setTimeState] = useState<'month'|'week'|'day'>('month');
  const [latestDataDate, setLatestDataDate] = useState('');
  const [isEditMode, setIsEditMode] = useState(false);
  const [selectedPeriod, setSelectedPeriod] = useState<string>('all');
  const [sourceFilter, setSourceFilter] = useState<string>('all');

  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 50;
  
  const loadData = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      let rows: any[] = [];
      if (isDeployed) {
         rows = await fetchViaProxy();
      } else if (CONFIG.CSV_URL) {
         rows = await fetchViaCSV(CONFIG.CSV_URL);
      }
      
      const enriched = rows.map(r => {
        const d = parseDate(r['友だち追加日時']);

        let source = (r['流入経路：最新'] || '').trim();
        if (source.includes('オーガニック@') || source.includes('オーガニック\uFFFD@') || source === 'オーガニック①') source = 'OG1';
        if (source.includes('オーガニックA') || source.includes('オーガニック\uFFFDA') || source === 'オーガニック②') source = 'OG2';
        
        if (!source || source === '') {
           if (isTrue(r['流入経路_OG1']) || isTrue(r['流入経路：オーガニック@']) || isTrue(r['流入経路：オーガニック\uFFFD@']) || isTrue(r['流入経路：オーガニック①'])) source = 'OG1';
           else if (isTrue(r['流入経路_OG2']) || isTrue(r['流入経路：オーガニックA']) || isTrue(r['流入経路：オーガニック\uFFFDA']) || isTrue(r['流入経路：オーガニック②'])) source = 'OG2';
        }

        return {
          ...r,
          _source: source,
          _dateObj: d,
          _month: formatMonth(d),
          _week: formatWeek(d),
          _day: formatDay(d)
        };
      });

      // valid dates
      const validDates = enriched.filter(r => r._dateObj).map(r => r._dateObj!.getTime());
      if (validDates.length > 0) {
        setLatestDataDate(formatDay(new Date(Math.max(...validDates))) || '');
      }

      setData(enriched);
    } catch (e: any) {
      setError(e.message || 'データの取得に失敗しました');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (isAuthenticated) {
      loadData();
    }
  }, [isAuthenticated, loadData]);

  useEffect(() => {
    if (isDeployed && !isAuthenticated) {
      const script = document.createElement('script');
      script.src = "https://accounts.google.com/gsi/client";
      script.async = true;
      script.onload = () => {
        if ((window as any).google) {
          (window as any).google.accounts.id.initialize({
            client_id: CONFIG.GOOGLE_CLIENT_ID,
            callback: (res: any) => {
              localStorage.setItem('google_id_token', res.credential);
              setIsAuthenticated(true);
            }
          });
          (window as any).google.accounts.id.renderButton(
            document.getElementById('google-signin-btn'),
            { theme: 'outline', size: 'large' }
          );
        }
      };
      document.body.appendChild(script);
      return () => { document.body.removeChild(script); };
    }
  }, [isAuthenticated]);

  const handleLogout = () => {
    localStorage.removeItem('google_id_token');
    setIsAuthenticated(false);
  };

  // Filters
  const periods = useMemo(() => {
    const prop = `_${timeState}`;
    const pSet = new Set(data.filter(r => r[prop]).map(r => r[prop]));
    const arr = Array.from(pSet).sort().reverse();
    return ['all', ...arr];
  }, [data, timeState]);

  // Adjust current period if timeState changes and it's not 'all'
  useEffect(() => {
    if (selectedPeriod !== 'all' && !periods.includes(selectedPeriod)) {
      setSelectedPeriod('all');
    }
  }, [periods, selectedPeriod]);

  const sources = useMemo(() => {
    const sSet = new Set(data.map(r => {
      // Find what source it has
      return r._source || '';
    }).filter(v => v));
    return ['all', ...Array.from(sSet)];
  }, [data]);

  const filteredData = useMemo(() => {
    return data.filter(r => {
      if (r['ユーザーブロック'] && isTrue(r['ユーザーブロック'])) return false;
      const prop = `_${timeState}`;
      if (selectedPeriod !== 'all' && r[prop] !== selectedPeriod) return false;
      if (sourceFilter !== 'all' && r._source !== sourceFilter) return false;
      return true;
    });
  }, [data, timeState, selectedPeriod, sourceFilter]);

  // Helper for metrics calculation
  const getPeriodData = (periodStr: string, isAllPeriod: boolean) => {
    const baseSet = data.filter(r => {
      if (r['ユーザーブロック'] && isTrue(r['ユーザーブロック'])) return false;
      if (sourceFilter !== 'all' && r._source !== sourceFilter) return false;
      if (!isAllPeriod) {
        const prop = `_${timeState}`;
        return r[prop] === periodStr;
      }
      return true;
    });

    const total = baseSet.length;
    // Main CV metrics
    const cv1 = baseSet.filter(r => isTrue(r['【CV】アンケート誘導_回答済み']) || isTrue(r['【全体】【CV】アンケート誘導_回答済み']) || isTrue(r['【OG1】【CV】アンケート誘導_回答済み']) || isTrue(r['【OG2】【CV】アンケート誘導_回答済み']) || r['表示名'] === '大江鉄人' || r['表示名'] === 'Ryuji' || isTrue(r['アンケート回答済み']) || (r['対応マーク'] && r['対応マーク'].includes('アンケート回答済み'))).length; // Fallback heuristic if mixed
    // Wait, let's look at the CSV: 対応マーク column has "アンケート回答済み". Or maybe "【CV】アンケート誘導_回答済み"? The user said tags are 0/1. Let's count keys containing 【CV】 and アンケート.
    
    // Actually simply finding any 1 in a corresponding column
    const isCvAnq = (r:any) => Object.keys(r).some(k => k.includes('【CV】') && k.includes('アンケート') && isTrue(r[k])) || r['対応マーク'] === 'アンケート回答済み';
    const isCvMendan = (r:any) => Object.keys(r).some(k => k.includes('【CV】') && k.includes('カジュアル面談_予約済み') && isTrue(r[k]));
    const isCvKanryo = (r:any) => Object.keys(r).some(k => k.includes('【CV】カジュアル面談_完了') && isTrue(r[k]));
    const isCancel = (r:any) => Object.keys(r).some(k => k.includes('キャンセル') && isTrue(r[k]));
    const isBlock = (r:any) => isTrue(r['ユーザーブロック']);

    const totalInclBlock = data.filter(r => {
      if (sourceFilter !== 'all' && r._source !== sourceFilter) return false;
      if (!isAllPeriod) {
         const prop = `_${timeState}`;
         return r[prop] === periodStr;
      }
      return true;
    }).length;

    const blocks = data.filter(r => {
      if (!isAllPeriod) {
         const prop = `_${timeState}`;
         if(r[prop] !== periodStr) return false;
      }
      if (sourceFilter !== 'all' && r._source !== sourceFilter) return false;
      return isBlock(r);
    }).length;

    return {
      total,
      totalInclBlock,
      blocks,
      cv1: baseSet.filter(isCvAnq).length,
      cv2: baseSet.filter(isCvMendan).length,
      cv3: baseSet.filter(isCvKanryo).length,
      cancels: baseSet.filter(isCancel).length,
      baseSet
    };
  };

  const currentMetrics = useMemo(() => getPeriodData(selectedPeriod, selectedPeriod === 'all'), [selectedPeriod, timeState, sourceFilter, data]);
  
  const previousMetrics = useMemo(() => {
    if (selectedPeriod === 'all') return null;
    const idx = periods.indexOf(selectedPeriod);
    if (idx < 0 || idx + 1 >= periods.length) return null;
    const prevPer = periods[idx + 1];
    return getPeriodData(prevPer, false);
  }, [selectedPeriod, periods, timeState, sourceFilter, data]);

  const calcChange = (curr: number, prev: number) => {
    if (!prev) return curr > 0 ? 100 : 0;
    return Math.round(((curr - prev) / prev) * 100);
  };
  const calcPointChange = (currRate: number, prevRate: number) => {
    return Number((currRate - prevRate).toFixed(1));
  };

  // Trend Data for graph
  const trendData = useMemo(() => {
    const map = new Map<string, any>();
    // Filter by source only, not period, for overall trend
    const sourceData = data.filter(r => {
      if (r['ユーザーブロック'] && isTrue(r['ユーザーブロック'])) return false;
      if (sourceFilter !== 'all' && r._source !== sourceFilter) return false;
      return r[`_${timeState}`];
    });

    const isCvAnq = (r:any) => Object.keys(r).some(k => k.includes('【CV】') && k.includes('アンケート') && isTrue(r[k])) || r['対応マーク'] === 'アンケート回答済み';

    sourceData.forEach(r => {
      const p = r[`_${timeState}`];
      if (!map.has(p)) map.set(p, { name: p, users: 0, cvs: 0 });
      const obj = map.get(p);
      obj.users += 1;
      if (isCvAnq(r)) obj.cvs += 1;
    });

    const arr = Array.from(map.values()).sort((a,b) => a.name.localeCompare(b.name));
    return arr.map(a => ({
      ...a,
      cvr: a.users > 0 ? Number(((a.cvs / a.users) * 100).toFixed(1)) : 0
    }));
  }, [data, timeState, sourceFilter]);

  // Scenario Data
  const scenarioData = useMemo(() => {
    const list: any[] = [];
    // Just find アンケート誘導 target & tap
    const tgts = ['登録直後', '10分後', '1時間後', '1日後 12:23', '1日後 20:03'];
    const getKeys = (kStr: string, r: any) => Object.keys(r).some(k => k.includes(`【${kStr}】アンケート誘導_${r._source === 'OG2' ? 'OG2' : 'OG1'}`) || k.includes(`【${kStr}】アンケート誘導_`) && isTrue(r[k]));
    
    tgts.forEach(t => {
      let tgCount = 0;
      let tpCount = 0;
      filteredData.forEach(r => {
        // dynamic match for OG1/OG2 or general
        const sourceKeys = Object.keys(r);
        const hasTg = sourceKeys.some(k => k.includes('【対象者】アンケート誘導') && k.includes(t) && isTrue(r[k]));
        const hasTp = sourceKeys.some(k => k.includes('【タップ】アンケート誘導') && k.includes(t) && isTrue(r[k]));
        if (hasTg) tgCount++;
        if (hasTp) tpCount++;
      });
      list.push({
        name: t,
        対象者: tgCount,
        タップ数: tpCount,
        tapRate: tgCount > 0 ? Number(((tpCount / tgCount) * 100).toFixed(1)) : 0
      });
    });
    return list;
  }, [filteredData]);

  // Source distribution
  const sourceDistData = useMemo(() => {
    const map = new Map<string, number>();
    filteredData.forEach(r => {
      const src = r._source || '不明';
      map.set(src, (map.get(src) || 0) + 1);
    });
    return Array.from(map.entries())
      .filter(([_,v]) => v > 0)
      .map(([name, value]) => ({ name, value }))
      .sort((a,b) => b.value - a.value);
  }, [filteredData]);

  const bestSource = useMemo(() => {
    let best = { name: '', cvr: 0 };
    const srcMap = new Map<string, {users:number, cvs:number}>();
    filteredData.forEach(r => {
      const src = r._source;
      if (!src) return;
      if (!srcMap.has(src)) srcMap.set(src, {users:0, cvs:0});
      const o = srcMap.get(src)!;
      o.users++;
      const isCv = Object.keys(r).some(k => k.includes('【CV】') && isTrue(r[k])) || r['対応マーク'] === 'アンケート回答済み';
      if (isCv) o.cvs++;
    });
    
    Array.from(srcMap.entries()).forEach(([name, o]) => {
      if (o.users > 5) { // minimum 5 users for reliable CVR
        const cvr = (o.cvs / o.users) * 100;
        if (cvr > best.cvr) {
          best = { name, cvr };
        }
      }
    });
    return best.name ? `🏆 CVR最高: ${best.name} (${best.cvr.toFixed(1)}%)` : null;
  }, [filteredData]);


  if (!isAuthenticated && isDeployed) {
    return (
      <div className="min-h-screen bg-[#f9f9f9] flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-xl shadow-lg w-full max-w-md text-center">
          <div className="w-16 h-16 bg-[#0067b8]/10 rounded-2xl flex items-center justify-center mx-auto mb-6">
            <LayoutDashboard className="text-[#0067b8]" size={32} />
          </div>
          <h1 className="text-2xl font-bold text-gray-900 mb-2">{CONFIG.TITLE}</h1>
          <p className="text-gray-500 text-sm mb-8">ダッシュボードへアクセスするには、Googleアカウントでログインしてください。</p>
          <div id="google-signin-btn" className="flex justify-center"></div>
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center bg-white">
        <Loader2 className="animate-spin text-[#0067b8] mb-4" size={48} />
        <p className="text-[#666] font-medium tracking-wider">データを同期中...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white">
        <div className="bg-red-50 p-6 rounded-lg text-red-600 flex flex-col items-center max-w-lg text-center">
          <AlertCircle size={48} className="mb-4 text-red-500" />
          <p className="font-semibold mb-2">エラーが発生しました</p>
          <p className="text-sm opacity-90">{error}</p>
        </div>
      </div>
    );
  }

  const cv1Rate = currentMetrics.total > 0 ? (currentMetrics.cv1 / currentMetrics.total) * 100 : 0;
  const cv2Rate = currentMetrics.cv1 > 0 ? (currentMetrics.cv2 / currentMetrics.cv1) * 100 : 0; // against total, or previous step? Let's use overall
  const cv2overall = currentMetrics.total > 0 ? (currentMetrics.cv2 / currentMetrics.total) * 100 : 0;

  const prevCv1Rate = previousMetrics && previousMetrics.total > 0 ? (previousMetrics.cv1 / previousMetrics.total) * 100 : 0;
  const prevCv2Overall = previousMetrics && previousMetrics.total > 0 ? (previousMetrics.cv2 / previousMetrics.total) * 100 : 0;

  return (
    <div className="min-h-screen pb-20">
      
      {/* Header */}
      <header className="bg-white border-b border-[#f2f2f2] sticky top-0 z-50">
        {isEditMode && (
          <div className="bg-[#ffb900]/10 text-[#a67b00] text-xs font-semibold px-4 py-2 flex items-center justify-center gap-2">
            <AlertCircle size={14}/>
            <span>編集モード: 変更は一時的です（リロードで元に戻ります）</span>
          </div>
        )}
        <div className="max-w-[1400px] mx-auto px-6 py-4 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <h1 className="text-[24px] font-semibold text-[#000] tracking-tight">{CONFIG.TITLE}</h1>
            <div className="flex items-center gap-3 mt-1.5 flex-wrap">
               <span className="bg-[#f2f2f2] text-[#666] px-2 py-0.5 rounded text-[11px] font-medium flex items-center gap-1">
                 <CalendarCheck size={12}/> LINE流入日ベース
               </span>
               {latestDataDate && (
                 <span className="text-[#666] text-[11px] flex items-center gap-1">
                   <CheckCircle size={12} className="text-[#107c10]"/> 有効データ最終日: {latestDataDate}
                 </span>
               )}
            </div>
          </div>
          
          <div className="flex items-center gap-3 flex-wrap">
            <div className="flex bg-[#f2f2f2] p-1 rounded-md">
              <button 
                onClick={() => { setTimeState('month'); setSelectedPeriod('all'); }} 
                className={cn("px-3 py-1.5 text-xs font-semibold rounded transition-colors", timeState==='month'?'bg-white shadow-sm text-[#000]':'text-[#666] hover:bg-white/50')}
              >月次</button>
              <button 
                onClick={() => { setTimeState('week'); setSelectedPeriod('all'); }} 
                className={cn("px-3 py-1.5 text-xs font-semibold rounded transition-colors", timeState==='week'?'bg-white shadow-sm text-[#000]':'text-[#666] hover:bg-white/50')}
              >週次</button>
              <button 
                onClick={() => { setTimeState('day'); setSelectedPeriod('all'); }} 
                className={cn("px-3 py-1.5 text-xs font-semibold rounded transition-colors", timeState==='day'?'bg-white shadow-sm text-[#000]':'text-[#666] hover:bg-white/50')}
              >日次</button>
            </div>

            <div className="relative">
              <select 
                value={selectedPeriod} 
                onChange={(e) => setSelectedPeriod(e.target.value)}
                className="appearance-none bg-white border border-[#d2d2d2] rounded-md px-3 py-2 pr-8 text-sm font-medium focus:outline-none focus:border-[#0067b8]"
              >
                <option value="all">全期間 （{timeState==='month'?'月':timeState==='week'?'週':'日'}）</option>
                {periods.filter(p => p !== 'all').map(p => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
              <ChevronDown size={14} className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[#666] pointer-events-none"/>
            </div>

            <div className="relative">
              <select 
                value={sourceFilter} 
                onChange={(e) => setSourceFilter(e.target.value)}
                className="appearance-none bg-white border border-[#d2d2d2] rounded-md px-3 py-2 pr-8 text-sm font-medium focus:outline-none focus:border-[#0067b8]"
              >
                <option value="all">すべての流入経路</option>
                {sources.filter(s => s !== 'all').map(s => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
              <ChevronDown size={14} className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[#666] pointer-events-none"/>
            </div>

            <div className="flex gap-2">
              <button onClick={loadData} className="p-2 text-[#666] hover:bg-[#f2f2f2] rounded-md transition-colors" title="データを再同期">
                <RefreshCw size={18} />
              </button>
              <button 
                onClick={() => setIsEditMode(!isEditMode)} 
                className={cn("p-2 rounded-md transition-colors", isEditMode ? "bg-[#e5f0fa] text-[#0067b8]" : "text-[#666] hover:bg-[#f2f2f2]")}
                title="編集モード（テキストの書き換え）"
              >
                <Edit3 size={18} />
              </button>
              {isAuthenticated && (
                <button onClick={handleLogout} className="p-2 text-[#d13438] hover:bg-red-50 rounded-md transition-colors" title="ログアウト">
                  <LogOut size={18} />
                </button>
              )}
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-[1400px] mx-auto px-6 py-8 space-y-8 animate-fadeIn">
        
        {bestSource && (
          <div className="bg-[#e5f0fa] border border-[#0067b8]/20 px-4 py-3 rounded-lg flex items-center justify-between">
             <span className="text-sm font-medium text-[#0067b8] flex gap-2"><TrendingUp size={18}/> 各種分析ハイライト</span>
             <span className="text-sm font-bold text-[#000]">{bestSource}</span>
          </div>
        )}

        <section>
          <div className="flex items-center gap-2 mb-4">
            <h2 className="text-[18px] font-semibold text-[#000] flex items-center gap-2">
              <LayoutDashboard size={20} className="text-[#0067b8]"/> 主要KPI
            </h2>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-4">
            {/* KPI 1 : Total Added */}
            <KPICard 
              title="有効ユーザー数" 
              value={currentMetrics.total} 
              unit="人" 
              icon="users" 
              info="選択期間内の「友だち追加日時」がある全行数（ブロックユーザーを除く）"
              isEditMode={isEditMode}
              change={previousMetrics ? calcChange(currentMetrics.total, previousMetrics.total) : null}
              changeLabel={timeState==='month'?'前月比':timeState==='week'?'前週比':'前日比'}
              subText={`ブロック含む全体: ${currentMetrics.totalInclBlock}人`}
            />

            {/* KPI 2 : Questionnaire Answering */}
            <KPICard 
              title="アンケート回答数" 
              value={currentMetrics.cv1} 
              unit="人" 
              icon="clipboard-list" 
              info="「【CV】アンケート誘導_回答済み」等がいずれか1のユーザー数"
              isEditMode={isEditMode}
              change={previousMetrics ? calcChange(currentMetrics.cv1, previousMetrics.cv1) : null}
              changeLabel={timeState==='month'?'前月比':timeState==='week'?'前週比':'前日比'}
              subText={`回答率: ${cv1Rate.toFixed(1)}%`}
            />

            {/* KPI 3 : Casual Interview Reservation */}
            <KPICard 
              title="面談予約数" 
              value={currentMetrics.cv2} 
              unit="人" 
              icon="calendar-check" 
              info="「【CV】カジュアル面談_予約済み」がいずれか1のユーザー数"
              isEditMode={isEditMode}
              change={previousMetrics ? calcChange(currentMetrics.cv2, previousMetrics.cv2) : null}
              changeLabel={timeState==='month'?'前月比':timeState==='week'?'前週比':'前日比'}
              subText={`全体予約率: ${cv2overall.toFixed(1)}%`}
            />

            {/* KPI 4 : Cancellation */}
            <KPICard 
              title="面談キャンセル数" 
              value={currentMetrics.cancels} 
              unit="人" 
              icon="user-minus" 
              info="「カジュアル面談_キャンセル」が1のユーザー数"
              isEditMode={isEditMode}
              change={previousMetrics ? calcChange(currentMetrics.cancels, previousMetrics.cancels) : null}
              changeLabel={timeState==='month'?'前月比':timeState==='week'?'前週比':'前日比'}
              subText={currentMetrics.cv2 > 0 ? `キャンセル率: ${((currentMetrics.cancels/currentMetrics.cv2)*100).toFixed(1)}% (対予約)` : `キャンセル率: 0%`}
            />

            {/* KPI 5 : Block */}
            <KPICard 
              title="ブロック数" 
              value={currentMetrics.blocks} 
              unit="人" 
              icon="x-circle" 
              info="「ユーザーブロック」が1のユーザー数"
              isEditMode={isEditMode}
              change={previousMetrics ? calcChange(currentMetrics.blocks, previousMetrics.blocks) : null}
              changeLabel={timeState==='month'?'前月比':timeState==='week'?'前週比':'前日比'}
              subText={`ブロック率: ${currentMetrics.totalInclBlock > 0 ? ((currentMetrics.blocks/currentMetrics.totalInclBlock)*100).toFixed(1) : 0}%`}
            />
          </div>
        </section>

        {/* Charts Row */}
        <section className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Trend Chart */}
          <div className="card p-6 min-h-[360px] flex flex-col">
            <h3 className="section-label mb-6 flex justify-between items-center">
              <span>アクション推移 ({timeState==='month'?'月次':timeState==='week'?'週次':'日次'})</span>
              <InfoTooltip text="期間別の登録ユーザー数（棒）とアンケート回答率（折れ線）の推移。"/>
            </h3>
            <div className="flex-1 w-full min-h-[280px]">
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart data={trendData} margin={{ top: 5, right: 0, left: -20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f2f2f2" vertical={false} />
                  <XAxis dataKey="name" tick={{fill:'#666', fontSize:11}} tickMargin={8} axisLine={false} tickLine={false} />
                  <YAxis yAxisId="left" tick={{fill:'#666', fontSize:11}} axisLine={false} tickLine={false} />
                  <YAxis yAxisId="right" orientation="right" tick={{fill:'#666', fontSize:11}} tickFormatter={(t)=>`${t}%`} axisLine={false} tickLine={false} />
                  <RechartsTooltip content={<CustomTooltip />} />
                  <Legend wrapperStyle={{fontSize:'12px', color:'#666', marginTop:'10px'}} iconType="circle" />
                  <Bar yAxisId="left" name="有効ユーザー数" dataKey="users" fill="#0067b8" radius={[4,4,0,0]} maxBarSize={60} />
                  <Bar yAxisId="left" name="アンケート回答数" dataKey="cvs" fill="#00A4EF" radius={[4,4,0,0]} maxBarSize={60} />
                  <Line yAxisId="right" type="monotone" name="アンケート回答率" dataKey="cvr" stroke="#ffb900" strokeWidth={3} dot={{r:4, fill:'#ffb900'}} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Scenario Funnel */}
          <div className="card p-6 min-h-[360px] flex flex-col">
            <h3 className="section-label mb-6 flex justify-between items-center">
              <span>シナリオ配信通目別 分析 (アンケート誘導)</span>
              <InfoTooltip text="アンケート誘導シナリオの各配信の対象者数とタップ率の遷移"/>
            </h3>
            <div className="flex-1 w-full min-h-[280px]">
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart data={scenarioData} margin={{ top: 5, right: 0, left: -20, bottom: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f2f2f2" vertical={false} />
                  <XAxis dataKey="name" tick={{fill:'#666', fontSize:10}} tickMargin={8} axisLine={false} tickLine={false} interval={0} angle={-25} textAnchor="end"/>
                  <YAxis yAxisId="left" tick={{fill:'#666', fontSize:11}} axisLine={false} tickLine={false} />
                  <YAxis yAxisId="right" orientation="right" tick={{fill:'#666', fontSize:11}} tickFormatter={(t)=>`${t}%`} axisLine={false} tickLine={false} />
                  <RechartsTooltip content={<CustomTooltip />} />
                  <Legend wrapperStyle={{fontSize:'12px', color:'#666', bottom:'-10px'}} iconType="circle" />
                  <Bar yAxisId="left" name="対象者数" dataKey="対象者" fill="#f2f2f2" radius={[4,4,0,0]} maxBarSize={50} />
                  <Bar yAxisId="left" name="タップ数" dataKey="タップ数" fill="#00A4EF" radius={[4,4,0,0]} maxBarSize={50} />
                  <Line yAxisId="right" type="monotone" name="タップ率" dataKey="tapRate" stroke="#107c10" strokeWidth={3} dot={{r:4, fill:'#107c10'}} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Source Distribution */}
          <div className="card p-6 min-h-[360px] flex flex-col">
            <h3 className="section-label mb-6 flex justify-between items-center">
              <span>流入経路分布</span>
              <InfoTooltip text="「流入経路：最新」別のユーザー数シェア。"/>
            </h3>
            {sourceDistData.length > 0 ? (
              <div className="flex-1 w-full min-h-[280px] flex md:flex-row flex-col items-center">
                <div className="h-full w-full md:w-1/2 min-h-[200px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={sourceDistData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} paddingAngle={2} innerRadius={50}>
                        {sourceDistData.map((e, index) => (
                          <Cell key={`cell-${index}`} fill={PIE_COLORS[index % PIE_COLORS.length]} />
                        ))}
                      </Pie>
                      <RechartsTooltip content={<CustomTooltip />} />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
                <div className="w-full md:w-1/2 flex flex-col gap-2 pl-4 max-h-[240px] overflow-y-auto custom-scrollbar">
                  {sourceDistData.map((item, idx) => (
                    <div key={idx} className="flex items-center justify-between text-sm">
                      <div className="flex items-center gap-2 overflow-hidden">
                        <div className="w-3 h-3 rounded shrink-0" style={{backgroundColor: PIE_COLORS[idx % PIE_COLORS.length]}}></div>
                        <span className="truncate text-[#333] font-medium text-xs">{item.name}</span>
                      </div>
                      <span className="font-bold text-[#000] ml-2 shrink-0">{item.value.toLocaleString()}人</span>
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="flex-1 flex items-center justify-center text-[#666] text-sm">データがありません</div>
            )}
          </div>
          
        </section>

        {/* Data Table */}
        <section className="card overflow-hidden">
          <div className="p-4 border-b border-[#f2f2f2] flex justify-between items-center">
             <h3 className="section-label">データテーブル (最新50行ずつ)</h3>
             <span className="text-xs text-[#666] font-medium">全 {filteredData.length.toLocaleString()} 件中 {currentPage*itemsPerPage - itemsPerPage + 1}〜{Math.min(currentPage*itemsPerPage, filteredData.length)}件</span>
          </div>
          <div className="overflow-x-auto custom-scrollbar max-h-[500px]">
             <table className="w-full text-left border-collapse select-none">
               <thead className="bg-[#f2f2f2] text-[#666] sticky top-0 z-10 shadow-sm">
                 <tr>
                   <th className="py-2.5 px-4 text-[12px] font-semibold whitespace-nowrap hidden md:table-cell">ID</th>
                   <th className="py-2.5 px-4 text-[12px] font-semibold whitespace-nowrap">表示名</th>
                   <th className="py-2.5 px-4 text-[12px] font-semibold whitespace-nowrap">友だち追加日時</th>
                   <th className="py-2.5 px-4 text-[12px] font-semibold whitespace-nowrap">流入経路</th>
                   <th className="py-2.5 px-4 text-[12px] font-semibold whitespace-nowrap">対応マーク</th>
                   {/* additional relevant sample columns */}
                   <th className="py-2.5 px-4 text-[12px] font-semibold whitespace-nowrap text-center">面談完了</th>
                 </tr>
               </thead>
               <tbody className="text-[#000] text-[13px] bg-white">
                 {filteredData.slice((currentPage-1)*itemsPerPage, currentPage*itemsPerPage).map((r, i) => (
                   <tr key={i} className="border-b border-[#f9f9f9] hover:bg-[#f9f9f9] transition-colors">
                     <td className="py-2.5 px-4 whitespace-nowrap font-mono text-[11px] text-[#666] hidden md:table-cell">{r['ID'] || '-'}</td>
                     <td className="py-2.5 px-4 whitespace-nowrap font-medium">{r['表示名'] || '-'}</td>
                     <td className="py-2.5 px-4 whitespace-nowrap text-[#666]">{r['友だち追加日時'] || '-'}</td>
                     <td className="py-2.5 px-4 whitespace-nowrap">
                       {r._source ? <span className="bg-[#e5f0fa] text-[#0067b8] px-2 py-0.5 rounded text-[11px] font-medium">{r._source}</span> : '-'}
                     </td>
                     <td className="py-2.5 px-4 whitespace-nowrap truncate max-w-[200px]" title={r['対応マーク']}>{r['対応マーク'] || '-'}</td>
                     <td className="py-2.5 px-4 whitespace-nowrap text-center">
                       {isTrue(r['【CV】カジュアル面談_完了']) ? <CheckCircle size={16} className="text-[#107c10] mx-auto"/> : <span className="text-[#ccc]">-</span>}
                     </td>
                   </tr>
                 ))}
                 {filteredData.length === 0 && (
                   <tr>
                     <td colSpan={6} className="py-8 text-center text-[#666] text-sm">
                       条件に一致するデータがありません
                     </td>
                   </tr>
                 )}
               </tbody>
             </table>
          </div>
          
          {/* Pagination */}
          {filteredData.length > itemsPerPage && (
            <div className="p-4 border-t border-[#f2f2f2] flex items-center justify-between">
               <button 
                onClick={()=>setCurrentPage(p=>Math.max(1, p-1))}
                disabled={currentPage===1}
                className="btn-secondary flex items-center gap-1 disabled:opacity-50"
               ><ChevronLeft size={14}/> 前のページ</button>
               
               <div className="text-xs text-[#666] font-medium flex gap-1">
                 {Array.from({length: Math.ceil(filteredData.length/itemsPerPage)}).map((_, idx) => {
                   if (idx === 0 || idx === Math.ceil(filteredData.length/itemsPerPage) - 1 || Math.abs(idx + 1 - currentPage) <= 2) {
                     return (
                       <button 
                        key={idx} 
                        onClick={()=>setCurrentPage(idx+1)}
                        className={cn("w-6 h-6 rounded flex items-center justify-center transition-colors", currentPage === idx+1 ? "bg-[#0067b8] text-white" : "hover:bg-[#f2f2f2]")}
                       >{idx+1}</button>
                     );
                   } else if (Math.abs(idx + 1 - currentPage) === 3) {
                     return <span key={idx} className="self-end px-1">...</span>;
                   }
                   return null;
                 })}
               </div>

               <button 
                onClick={()=>setCurrentPage(p=>Math.min(Math.ceil(filteredData.length/itemsPerPage), p+1))}
                disabled={currentPage===Math.ceil(filteredData.length/itemsPerPage)}
                className="btn-secondary flex items-center gap-1 disabled:opacity-50"
               >次のページ <ChevronRight size={14}/></button>
            </div>
          )}
        </section>

      </main>
    </div>
  );
}


