import React, { useState, useEffect, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  Plus, Trash2, Upload, Download, Mail, Check, ChevronUp, ChevronDown,
  FileSpreadsheet, Building2, Package, Send, RefreshCw, Pencil, X,
  Search, Stamp, ArrowRight, AlertCircle, CheckCircle2
} from 'lucide-react';

/* ------------------------------------------------------------------ */
/* Utilities                                                           */
/* ------------------------------------------------------------------ */

const uid = () => Math.random().toString(36).slice(2, 10);

const today = () => {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`;
};

const colLetter = (idx) => {
  let s = '';
  let i = idx + 1;
  while (i > 0) {
    const m = (i - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
};

const normalize = (s) => String(s || '').replace(/\s+/g, '').toLowerCase();

const SYNONYMS = {
  '품목코드': ['코드', '품번', 'sku', 'itemcode', '상품코드', '품목코드'],
  '품목명': ['품명', '제품명', '상품명', '품목', 'item', 'itemname', '제품', '상품'],
  '규격': ['사양', '규격', '스펙', 'spec', 'size', '옵션', '모델명'],
  '수량': ['수량', '발주수량', 'qty', 'ea', '개수', '주문수량'],
  '단가': ['단가', '구매단가', '공급단가', 'unitprice', '매입단가'],
  '공급가액': ['금액', '공급가액', '합계', '총액', 'amount', 'total', '공급가'],
  '비고': ['비고', '메모', 'remark', 'note', '특이사항'],
};

function autoMatch(masterFields, supplierCols) {
  const mapping = {};
  masterFields.forEach((f) => {
    const syns = (SYNONYMS[f.name] || [f.name]).map(normalize);
    const fname = normalize(f.name);
    let best = null;
    supplierCols.forEach((c) => {
      const h = normalize(c.header);
      if (!h) return;
      let score = 0;
      if (h === fname) score = 100;
      else if (syns.includes(h)) score = 90;
      else if (h.includes(fname) || fname.includes(h)) score = 70;
      else if (syns.some((s) => h.includes(s) || s.includes(h))) score = 60;
      if (score > 0 && (!best || score > best.score)) best = { colIndex: c.colIndex, score };
    });
    mapping[f.id] = best ? best.colIndex : null;
  });
  return mapping;
}

function arrayBufferToBase64(buf) {
  let binary = '';
  const bytes = new Uint8Array(buf);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}
function base64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function detectHeaderRow(aoa) {
  let bestIdx = 0, bestScore = -1;
  const limit = Math.min(aoa.length, 15);
  for (let i = 0; i < limit; i++) {
    const row = aoa[i] || [];
    const nonEmptyStrings = row.filter((c) => typeof c === 'string' && c.trim() !== '').length;
    if (nonEmptyStrings > bestScore) { bestScore = nonEmptyStrings; bestIdx = i; }
  }
  return bestIdx;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// rows: [{ values: { [masterFieldId]: rawValue } }]  →  매입처 원본 양식에 값을 채운 파일 생성
async function buildSupplierFile(supplier, rows, masterFields, loadSupplierTemplate) {
  const base64 = await loadSupplierTemplate(supplier.id);
  if (!base64) throw new Error('NO_TEMPLATE');

  const wb = XLSX.read(base64ToArrayBuffer(base64), { type: 'array', cellStyles: true });
  const ws = wb.Sheets[supplier.sheetName];
  let maxR = 0, maxC = 0;
  const existingRange = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
  if (existingRange) { maxR = existingRange.e.r; maxC = existingRange.e.c; }

  rows.forEach((row, i) => {
    const r = supplier.dataStartRowIdx + i;
    masterFields.forEach((f) => {
      const c = supplier.mapping?.[f.id];
      if (c === null || c === undefined) return;
      const raw = row.values[f.id];
      if (raw === undefined || String(raw).trim() === '') return;
      const num = Number(raw);
      const isNum = raw !== '' && !isNaN(num) && String(raw).trim() !== '' && /^-?\d+(\.\d+)?$/.test(String(raw).trim());
      const addr = XLSX.utils.encode_cell({ r, c });
      ws[addr] = isNum ? { t: 'n', v: num } : { t: 's', v: String(raw) };
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    });
  });
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  const fileName = `${today()}_${supplier.name}_발주서.xlsx`;
  return { fileName, base64: arrayBufferToBase64(out), itemCount: rows.length };
}

function findSupplierByName(suppliers, rawName) {
  const n = normalize(rawName);
  if (!n) return null;
  return (
    suppliers.find((s) => normalize(s.name) === n) ||
    suppliers.find((s) => normalize(s.name).includes(n) || n.includes(normalize(s.name))) ||
    null
  );
}

// Vercel 배포판에서는 window.storage 대신 이 브라우저의 localStorage에 저장합니다.
// (매입처/양식 정보는 이 브라우저에만 저장되며 서버로 전송되지 않습니다)
function nsKey(key, shared) {
  return `po-system:${shared ? 'shared' : 'local'}:${key}`;
}
async function storageGet(key, fallback, shared = false) {
  try {
    const raw = localStorage.getItem(nsKey(key, shared));
    if (raw === null) return fallback;
    return JSON.parse(raw);
  } catch (e) { return fallback; }
}
async function storageSet(key, value, shared = false) {
  try { localStorage.setItem(nsKey(key, shared), JSON.stringify(value)); }
  catch (e) { console.error('storage set failed', key, e); }
}

/* ------------------------------------------------------------------ */
/* Defaults                                                             */
/* ------------------------------------------------------------------ */

const DEFAULT_MASTER_FIELDS = [
  { id: uid(), name: '품목코드' },
  { id: uid(), name: '품목명' },
  { id: uid(), name: '규격' },
  { id: uid(), name: '수량' },
  { id: uid(), name: '단가' },
  { id: uid(), name: '공급가액' },
  { id: uid(), name: '비고' },
];

/* ------------------------------------------------------------------ */
/* Main App                                                            */
/* ------------------------------------------------------------------ */

export default function App() {
  const [ready, setReady] = useState(false);
  const [tab, setTab] = useState('master');
  const [masterFields, setMasterFields] = useState(DEFAULT_MASTER_FIELDS);
  const [suppliers, setSuppliers] = useState([]);
  const [sendLog, setSendLog] = useState([]);
  const [genList, setGenList] = useState([]); // {id, supplierId, supplierName, email, fileName, base64, createdAt, itemCount}
  const [toast, setToast] = useState(null);

  useEffect(() => {
    (async () => {
      const mf = await storageGet('master-fields', DEFAULT_MASTER_FIELDS);
      const sup = await storageGet('suppliers', []);
      const log = await storageGet('send-log', []);
      setMasterFields(mf); setSuppliers(sup); setSendLog(log);
      setReady(true);
    })();
  }, []);

  useEffect(() => { if (ready) storageSet('master-fields', masterFields); }, [masterFields, ready]);
  useEffect(() => { if (ready) storageSet('suppliers', suppliers.map(stripBinary)); }, [suppliers, ready]);
  useEffect(() => { if (ready) storageSet('send-log', sendLog); }, [sendLog, ready]);

  function stripBinary(s) {
    // template binary is stored under its own key, keep metadata light here
    const { templateBase64, ...rest } = s;
    return rest;
  }

  const showToast = useCallback((msg, kind = 'ok') => {
    setToast({ msg, kind, id: uid() });
    setTimeout(() => setToast((t) => (t && t.msg === msg ? null : t)), 2800);
  }, []);

  async function saveSupplierTemplate(id, base64) {
    await storageSet(`supplier-template:${id}`, base64);
  }
  async function loadSupplierTemplate(id) {
    return await storageGet(`supplier-template:${id}`, null);
  }

  if (!ready) {
    return (
      <div className="w-full h-full min-h-[500px] flex items-center justify-center bg-stone-50 text-slate-500 text-sm">
        불러오는 중…
      </div>
    );
  }

  return (
    <div className="w-full min-h-[640px] bg-stone-50 text-slate-800 flex flex-col font-sans">
      <Header tab={tab} setTab={setTab} supplierCount={suppliers.length} mailCount={genList.length} />

      <div className="flex-1 max-w-6xl w-full mx-auto px-4 md:px-6 py-6">
        {tab === 'master' && (
          <MasterFieldsTab masterFields={masterFields} setMasterFields={setMasterFields} showToast={showToast} />
        )}
        {tab === 'suppliers' && (
          <SuppliersTab
            masterFields={masterFields}
            suppliers={suppliers}
            setSuppliers={setSuppliers}
            saveSupplierTemplate={saveSupplierTemplate}
            loadSupplierTemplate={loadSupplierTemplate}
            showToast={showToast}
          />
        )}
        {tab === 'generate' && (
          <GenerateTab
            masterFields={masterFields}
            suppliers={suppliers}
            loadSupplierTemplate={loadSupplierTemplate}
            genList={genList}
            setGenList={setGenList}
            showToast={showToast}
          />
        )}
        {tab === 'mail' && (
          <MailTab
            genList={genList}
            sendLog={sendLog}
            setSendLog={setSendLog}
            showToast={showToast}
          />
        )}
      </div>

      {toast && (
        <div className={`fixed bottom-5 left-1/2 -translate-x-1/2 px-4 py-2.5 rounded shadow-lg text-sm font-medium flex items-center gap-2 ${toast.kind === 'ok' ? 'bg-slate-800 text-white' : 'bg-red-700 text-white'}`}>
          {toast.kind === 'ok' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
          {toast.msg}
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Header                                                               */
/* ------------------------------------------------------------------ */

function Header({ tab, setTab, supplierCount, mailCount }) {
  const tabs = [
    { id: 'master', label: '통합 양식', icon: Package },
    { id: 'suppliers', label: '매입처 관리', icon: Building2, badge: supplierCount },
    { id: 'generate', label: '발주서 생성', icon: FileSpreadsheet },
    { id: 'mail', label: '메일 발송', icon: Mail, badge: mailCount },
  ];
  return (
    <div className="border-b border-slate-300 bg-white">
      <div className="max-w-6xl mx-auto px-4 md:px-6 pt-5 pb-0">
        <div className="flex items-center gap-2.5 mb-4">
          <div className="w-9 h-9 rounded-sm bg-slate-800 flex items-center justify-center shrink-0">
            <Stamp size={18} className="text-red-500" />
          </div>
          <div>
            <h1 className="text-lg font-semibold tracking-tight leading-none">발주서 자동 생성</h1>
            <p className="text-xs text-slate-500 mt-1">매입처별 양식 매칭 · 발주서 생성 · 메일 발송</p>
          </div>
        </div>
        <div className="flex gap-1">
          {tabs.map((t) => {
            const Icon = t.icon;
            const active = tab === t.id;
            return (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                className={`flex items-center gap-1.5 px-3.5 py-2.5 text-sm font-medium border-b-2 transition-colors ${
                  active ? 'border-red-700 text-slate-900' : 'border-transparent text-slate-500 hover:text-slate-800'
                }`}
              >
                <Icon size={15} />
                {t.label}
                {typeof t.badge === 'number' && t.badge > 0 && (
                  <span className={`text-[11px] px-1.5 py-0.5 rounded-full ${active ? 'bg-slate-800 text-white' : 'bg-slate-200 text-slate-600'}`}>
                    {t.badge}
                  </span>
                )}
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Tab 1: 통합 양식 (Master Fields)                                     */
/* ------------------------------------------------------------------ */

function MasterFieldsTab({ masterFields, setMasterFields, showToast }) {
  const [newName, setNewName] = useState('');
  const fileRef = useRef(null);
  const [busy, setBusy] = useState(false);

  const addField = () => {
    const name = newName.trim();
    if (!name) return;
    setMasterFields([...masterFields, { id: uid(), name }]);
    setNewName('');
  };
  const removeField = (id) => setMasterFields(masterFields.filter((f) => f.id !== id));
  const renameField = (id, name) => setMasterFields(masterFields.map((f) => (f.id === id ? { ...f, name } : f)));
  const move = (idx, dir) => {
    const next = [...masterFields];
    const j = idx + dir;
    if (j < 0 || j >= next.length) return;
    [next[idx], next[j]] = [next[j], next[idx]];
    setMasterFields(next);
  };

  const handleUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const headerRowIdx = detectHeaderRow(aoa);
      const headerRow = aoa[headerRowIdx] || [];
      const seen = new Set();
      const detected = [];
      headerRow.forEach((h) => {
        const name = String(h || '').trim();
        if (!name) return;
        const n = normalize(name);
        if (seen.has(n)) return;
        seen.add(n);
        detected.push({ id: uid(), name });
      });
      if (detected.length === 0) {
        showToast && showToast('열 제목을 찾지 못했습니다. 다른 파일을 시도해주세요.', 'err');
      } else {
        setMasterFields(detected);
        showToast && showToast(`${detected.length}개 항목을 순서대로 불러왔습니다.`);
      }
    } catch (err) {
      console.error(err);
      showToast && showToast('엑셀 파일을 읽는 중 오류가 발생했습니다.', 'err');
    } finally {
      setBusy(false);
      if (fileRef.current) fileRef.current.value = '';
    }
  };

  return (
    <div className="max-w-2xl">
      <div className="flex items-start justify-between gap-4 mb-4">
        <SectionHeading
          title="통합 양식 항목"
          desc="모든 매입처 발주서 생성의 기준이 되는 공통 항목입니다. 매입처를 등록할 때 이 항목을 기준으로 자동 매칭됩니다."
        />
        <button
          onClick={() => fileRef.current?.click()}
          disabled={busy}
          className="shrink-0 flex items-center gap-1.5 text-xs px-2.5 py-1.5 border border-slate-300 rounded-md text-slate-600 hover:bg-slate-50 disabled:opacity-50"
        >
          <Upload size={13} /> {busy ? '분석 중…' : '기존 발주서로 항목 불러오기'}
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUpload} />
      </div>

      <div className="bg-white border border-slate-200 rounded-md divide-y divide-slate-100 mb-4">
        {masterFields.length === 0 && (
          <div className="px-4 py-8 text-center text-sm text-slate-400">등록된 항목이 없습니다.</div>
        )}
        {masterFields.map((f, idx) => (
          <div key={f.id} className="flex items-center gap-2 px-3 py-2 group">
            <span className="text-xs font-mono text-slate-400 w-6 shrink-0">{idx + 1}</span>
            <input
              value={f.name}
              onChange={(e) => renameField(f.id, e.target.value)}
              className="flex-1 text-sm px-2 py-1.5 border border-transparent hover:border-slate-200 focus:border-slate-400 rounded outline-none bg-transparent"
            />
            <div className="flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
              <IconBtn onClick={() => move(idx, -1)} title="위로"><ChevronUp size={14} /></IconBtn>
              <IconBtn onClick={() => move(idx, 1)} title="아래로"><ChevronDown size={14} /></IconBtn>
              <IconBtn onClick={() => removeField(f.id)} title="삭제" danger><Trash2 size={14} /></IconBtn>
            </div>
          </div>
        ))}
      </div>

      <div className="flex gap-2">
        <input
          value={newName}
          onChange={(e) => setNewName(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && addField()}
          placeholder="새 항목 이름 (예: 납기일)"
          className="flex-1 text-sm px-3 py-2 border border-slate-300 rounded-md outline-none focus:border-slate-500"
        />
        <button onClick={addField} className="flex items-center gap-1.5 px-3.5 py-2 bg-slate-800 text-white text-sm rounded-md hover:bg-slate-700">
          <Plus size={15} /> 추가
        </button>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Tab 2: 매입처 관리                                                   */
/* ------------------------------------------------------------------ */

function SuppliersTab({ masterFields, suppliers, setSuppliers, saveSupplierTemplate, loadSupplierTemplate, showToast }) {
  const [query, setQuery] = useState('');
  const [selectedId, setSelectedId] = useState(suppliers[0]?.id || null);
  const [showAddModal, setShowAddModal] = useState(false);
  const [newName, setNewName] = useState('');
  const [newEmail, setNewEmail] = useState('');

  const filtered = suppliers.filter((s) => normalize(s.name).includes(normalize(query)) || normalize(s.email).includes(normalize(query)));
  const selected = suppliers.find((s) => s.id === selectedId) || null;

  const addSupplier = () => {
    if (!newName.trim()) return showToast('매입처명을 입력해주세요.', 'err');
    const s = { id: uid(), name: newName.trim(), email: newEmail.trim(), mapping: {}, headerRowIdx: null, dataStartRowIdx: null, sheetName: null, cols: [], fileName: null };
    setSuppliers([s, ...suppliers]);
    setSelectedId(s.id);
    setNewName(''); setNewEmail(''); setShowAddModal(false);
  };
  const removeSupplier = (id) => {
    setSuppliers(suppliers.filter((s) => s.id !== id));
    if (selectedId === id) setSelectedId(null);
  };
  const updateSupplier = (id, patch) => setSuppliers(suppliers.map((s) => (s.id === id ? { ...s, ...patch } : s)));

  return (
    <div className="grid grid-cols-1 md:grid-cols-[280px_1fr] gap-5">
      {/* Left: supplier list */}
      <div>
        <SectionHeading title="매입처 목록" desc={`총 ${suppliers.length}곳`} compact />
        <div className="relative mb-2">
          <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            value={query} onChange={(e) => setQuery(e.target.value)}
            placeholder="매입처 검색"
            className="w-full text-sm pl-8 pr-2.5 py-2 border border-slate-300 rounded-md outline-none focus:border-slate-500"
          />
        </div>
        <button
          onClick={() => setShowAddModal(true)}
          className="w-full flex items-center justify-center gap-1.5 px-3 py-2 mb-2 border border-dashed border-slate-300 text-slate-600 text-sm rounded-md hover:border-slate-500 hover:text-slate-900"
        >
          <Plus size={14} /> 매입처 등록
        </button>
        <div className="border border-slate-200 rounded-md divide-y divide-slate-100 max-h-[520px] overflow-y-auto bg-white">
          {filtered.length === 0 && <div className="px-3 py-6 text-center text-xs text-slate-400">매입처가 없습니다.</div>}
          {filtered.map((s) => (
            <button
              key={s.id}
              onClick={() => setSelectedId(s.id)}
              className={`w-full text-left px-3 py-2.5 hover:bg-slate-50 flex items-center justify-between gap-2 ${selectedId === s.id ? 'bg-slate-50' : ''}`}
            >
              <div className="min-w-0">
                <div className="text-sm font-medium truncate flex items-center gap-1.5">
                  {s.name}
                  {selectedId === s.id && <ArrowRight size={12} className="text-red-700 shrink-0" />}
                </div>
                <div className="text-xs text-slate-400 truncate">{s.email || '이메일 미등록'}</div>
              </div>
              <span className={`shrink-0 text-[10px] px-1.5 py-0.5 rounded-full ${s.sheetName ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'}`}>
                {s.sheetName ? '양식 등록됨' : '양식 필요'}
              </span>
            </button>
          ))}
        </div>
      </div>

      {/* Right: detail */}
      <div>
        {selected && (
          <SupplierDetail
            key={selected.id}
            supplier={selected}
            masterFields={masterFields}
            updateSupplier={updateSupplier}
            removeSupplier={removeSupplier}
            saveSupplierTemplate={saveSupplierTemplate}
            loadSupplierTemplate={loadSupplierTemplate}
            showToast={showToast}
          />
        )}

        {!selected && (
          <div className="h-64 flex items-center justify-center text-sm text-slate-400 border border-dashed border-slate-200 rounded-md">
            왼쪽에서 매입처를 선택하거나 새로 등록해주세요.
          </div>
        )}
      </div>

      {showAddModal && (
        <div
          className="fixed inset-0 bg-slate-900/40 flex items-center justify-center z-50 px-4"
          onClick={() => setShowAddModal(false)}
        >
          <div
            className="bg-white rounded-md shadow-xl w-full max-w-sm p-5"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-4">
              <div className="text-sm font-semibold">매입처 등록</div>
              <IconBtn onClick={() => setShowAddModal(false)} title="닫기"><X size={15} /></IconBtn>
            </div>
            <div className="space-y-3 mb-4">
              <div>
                <label className="text-xs text-slate-500 mb-1 block">매입처명 *</label>
                <input
                  autoFocus
                  value={newName} onChange={(e) => setNewName(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && addSupplier()}
                  className="w-full text-sm px-3 py-2 border border-slate-300 rounded-md outline-none focus:border-slate-500" placeholder="예: 대한상사"
                />
              </div>
              <div>
                <label className="text-xs text-slate-500 mb-1 block">발주 접수 이메일</label>
                <input
                  value={newEmail} onChange={(e) => setNewEmail(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && addSupplier()}
                  className="w-full text-sm px-3 py-2 border border-slate-300 rounded-md outline-none focus:border-slate-500" placeholder="order@supplier.com"
                />
              </div>
            </div>
            <div className="flex justify-end gap-2">
              <button onClick={() => setShowAddModal(false)} className="px-3.5 py-2 text-slate-500 text-sm rounded-md hover:bg-slate-100">취소</button>
              <button onClick={addSupplier} className="px-3.5 py-2 bg-slate-800 text-white text-sm rounded-md hover:bg-slate-700">등록</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function SupplierDetail({ supplier, masterFields, updateSupplier, removeSupplier, saveSupplierTemplate, loadSupplierTemplate, showToast }) {
  const fileRef = useRef(null);
  const [editingInfo, setEditingInfo] = useState(false);
  const [name, setName] = useState(supplier.name);
  const [email, setEmail] = useState(supplier.email);
  const [busy, setBusy] = useState(false);

  useEffect(() => { setName(supplier.name); setEmail(supplier.email); }, [supplier.id]);

  const saveInfo = () => { updateSupplier(supplier.id, { name: name.trim() || supplier.name, email: email.trim() }); setEditingInfo(false); };

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellStyles: true });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const headerRowIdx = detectHeaderRow(aoa);
      const headerRow = aoa[headerRowIdx] || [];
      const cols = headerRow.map((h, i) => ({ colIndex: i, header: String(h || '').trim() })).filter((c, i) => c.header !== '' || i < headerRow.length);
      const mapping = autoMatch(masterFields, cols);
      const base64 = arrayBufferToBase64(buf);
      await saveSupplierTemplate(supplier.id, base64);
      updateSupplier(supplier.id, {
        sheetName, cols, mapping,
        headerRowIdx, dataStartRowIdx: headerRowIdx + 1,
        fileName: file.name,
      });
      showToast(`${file.name} 업로드 및 자동 매칭 완료`);
    } catch (err) {
      console.error(err);
      showToast('엑셀 파일을 읽는 중 오류가 발생했습니다.', 'err');
    } finally {
      setBusy(false);
      if (fileRef.current) fileRef.current.value = '';
    }
  };

  const rematch = () => {
    if (!supplier.cols?.length) return;
    updateSupplier(supplier.id, { mapping: autoMatch(masterFields, supplier.cols) });
    showToast('자동 매칭을 다시 실행했습니다.');
  };

  const setColumnMapping = (colIndex, fieldId) => {
    const next = { ...supplier.mapping };
    // 이 열에 이미 연결되어 있던 항목은 해제
    Object.keys(next).forEach((fid) => {
      if (next[fid] === colIndex) next[fid] = null;
    });
    if (fieldId) next[fieldId] = colIndex;
    updateSupplier(supplier.id, { mapping: next });
  };

  const setRow = (key, val) => {
    const n = Math.max(1, Number(val) || 1);
    updateSupplier(supplier.id, { [key]: n - 1 });
  };

  return (
    <div className="space-y-4">
      {/* Basic info card */}
      <div className="bg-white border border-slate-200 rounded-md p-4">
        {!editingInfo ? (
          <div className="flex items-start justify-between">
            <div>
              <div className="text-base font-semibold">{supplier.name}</div>
              <div className="text-sm text-slate-500 mt-0.5">{supplier.email || '이메일 미등록'}</div>
            </div>
            <div className="flex gap-1">
              <IconBtn onClick={() => setEditingInfo(true)} title="정보 수정"><Pencil size={14} /></IconBtn>
              <IconBtn onClick={() => removeSupplier(supplier.id)} title="매입처 삭제" danger><Trash2 size={14} /></IconBtn>
            </div>
          </div>
        ) : (
          <div className="flex flex-wrap items-end gap-2">
            <div>
              <label className="text-xs text-slate-500 mb-1 block">매입처명</label>
              <input value={name} onChange={(e) => setName(e.target.value)} className="text-sm px-3 py-2 border border-slate-300 rounded-md outline-none focus:border-slate-500" />
            </div>
            <div>
              <label className="text-xs text-slate-500 mb-1 block">발주 접수 이메일</label>
              <input value={email} onChange={(e) => setEmail(e.target.value)} className="text-sm px-3 py-2 border border-slate-300 rounded-md outline-none focus:border-slate-500" />
            </div>
            <button onClick={saveInfo} className="px-3 py-2 bg-slate-800 text-white text-sm rounded-md hover:bg-slate-700">저장</button>
            <button onClick={() => setEditingInfo(false)} className="px-3 py-2 text-slate-500 text-sm rounded-md hover:bg-slate-100">취소</button>
          </div>
        )}
      </div>

      {/* Template upload */}
      <div className="bg-white border border-slate-200 rounded-md p-4">
        <div className="flex items-center justify-between mb-3">
          <div className="text-sm font-semibold flex items-center gap-1.5"><FileSpreadsheet size={15} /> 매입처 엑셀 양식</div>
          <div className="flex gap-2">
            {supplier.sheetName && (
              <button onClick={rematch} className="flex items-center gap-1 text-xs px-2.5 py-1.5 border border-slate-300 rounded-md text-slate-600 hover:bg-slate-50">
                <RefreshCw size={12} /> 자동 매칭 다시
              </button>
            )}
            <button onClick={() => fileRef.current?.click()} disabled={busy} className="flex items-center gap-1 text-xs px-2.5 py-1.5 bg-slate-800 text-white rounded-md hover:bg-slate-700 disabled:opacity-50">
              <Upload size={12} /> {busy ? '업로드 중…' : supplier.sheetName ? '다시 업로드' : '엑셀 업로드'}
            </button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile} />
          </div>
        </div>

        {!supplier.sheetName ? (
          <div className="text-xs text-slate-400 border border-dashed border-slate-200 rounded-md py-6 text-center">
            이 매입처가 사용하는 발주서 양식(.xlsx)을 업로드하면, 통합 양식 항목이 자동으로 매칭됩니다.
          </div>
        ) : (
          <>
            <div className="text-xs text-slate-500 mb-3 flex flex-wrap gap-x-4 gap-y-1">
              <span>파일: <span className="text-slate-700 font-medium">{supplier.fileName}</span></span>
              <span>시트: <span className="text-slate-700 font-medium">{supplier.sheetName}</span></span>
              <span className="flex items-center gap-1">
                헤더 행:
                <input type="number" min={1} value={supplier.headerRowIdx + 1}
                  onChange={(e) => setRow('headerRowIdx', e.target.value)}
                  className="w-14 px-1.5 py-0.5 border border-slate-300 rounded text-slate-700" />
              </span>
              <span className="flex items-center gap-1">
                데이터 시작 행:
                <input type="number" min={1} value={supplier.dataStartRowIdx + 1}
                  onChange={(e) => setRow('dataStartRowIdx', e.target.value)}
                  className="w-14 px-1.5 py-0.5 border border-slate-300 rounded text-slate-700" />
              </span>
            </div>

            <div className="border border-slate-200 rounded-md overflow-hidden">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs">
                    <th className="text-left font-medium px-3 py-2 w-48">매입처 양식의 열</th>
                    <th className="text-left font-medium px-3 py-2">통합 항목</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {supplier.cols.map((c) => {
                    const matchedField = masterFields.find((f) => supplier.mapping?.[f.id] === c.colIndex);
                    const matched = !!matchedField;
                    return (
                      <tr key={c.colIndex}>
                        <td className="px-3 py-2 font-medium text-slate-700">
                          {colLetter(c.colIndex)}열{c.header ? ` — ${c.header}` : ' (제목 없음)'}
                        </td>
                        <td className="px-3 py-2">
                          <select
                            value={matched ? matchedField.id : ''}
                            onChange={(e) => setColumnMapping(c.colIndex, e.target.value)}
                            className={`text-sm px-2 py-1.5 border rounded-md outline-none w-full max-w-xs ${matched ? 'border-slate-300' : 'border-amber-300 bg-amber-50'}`}
                          >
                            <option value="">매칭 안 함</option>
                            {masterFields.map((f) => (
                              <option key={f.id} value={f.id}>{f.name}</option>
                            ))}
                          </select>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Tab 3: 발주서 생성                                                   */
/* ------------------------------------------------------------------ */

function GenerateTab({ masterFields, suppliers, loadSupplierTemplate, genList, setGenList, showToast }) {
  const ready = suppliers.filter((s) => s.sheetName);
  const [supplierId, setSupplierId] = useState(ready[0]?.id || '');
  const supplier = suppliers.find((s) => s.id === supplierId);

  const mappedFields = supplier ? masterFields.filter((f) => supplier.mapping?.[f.id] !== null && supplier.mapping?.[f.id] !== undefined) : [];

  const blankRow = () => {
    const v = {};
    mappedFields.forEach((f) => (v[f.id] = ''));
    return { id: uid(), values: v };
  };

  const [rows, setRows] = useState([]);
  useEffect(() => { setRows(mappedFields.length ? [blankRow()] : []); }, [supplierId]);

  const addRow = () => setRows([...rows, blankRow()]);
  const removeRow = (id) => setRows(rows.filter((r) => r.id !== id));
  const setCell = (rowId, fieldId, val) => setRows(rows.map((r) => (r.id === rowId ? { ...r, values: { ...r.values, [fieldId]: val } } : r)));

  const generate = async () => {
    if (!supplier) return;
    const validRows = rows.filter((r) => Object.values(r.values).some((v) => String(v).trim() !== ''));
    if (validRows.length === 0) return showToast('발주 품목을 1개 이상 입력해주세요.', 'err');

    try {
      const { fileName, base64, itemCount } = await buildSupplierFile(supplier, validRows, masterFields, loadSupplierTemplate);
      const buf = base64ToArrayBuffer(base64);
      downloadBlob(new Blob([buf], { type: 'application/octet-stream' }), fileName);

      const item = {
        id: uid(), supplierId: supplier.id, supplierName: supplier.name, email: supplier.email,
        fileName, base64, createdAt: Date.now(), itemCount,
      };
      setGenList([item, ...genList]);
      setRows([blankRow()]);
      showToast(`${fileName} 생성 완료`);
    } catch (err) {
      console.error(err);
      if (err.message === 'NO_TEMPLATE') showToast('저장된 양식을 찾을 수 없습니다. 다시 업로드해주세요.', 'err');
      else showToast('발주서 생성 중 오류가 발생했습니다.', 'err');
    }
  };

  /* ---------------- 통합 양식 업로드로 일괄 생성 ---------------- */
  const masterFileRef = useRef(null);
  const [batchBusy, setBatchBusy] = useState(false);

  const downloadMasterTemplate = () => {
    const header = ['매입처명', ...masterFields.map((f) => f.name)];
    const ws = XLSX.utils.aoa_to_sheet([header]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '발주입력');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    downloadBlob(new Blob([out], { type: 'application/octet-stream' }), '통합_발주양식.xlsx');
  };

  const handleMasterUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBatchBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const headerRowIdx = detectHeaderRow(aoa);
      const dataRows = aoa.slice(headerRowIdx + 1);

      const groups = {}; // supplierId -> rows[]
      const unmatched = new Set();
      const noTemplate = new Set();

      dataRows.forEach((row) => {
        const supplierNameRaw = String(row[0] ?? '').trim();
        if (!supplierNameRaw) return;
        const found = findSupplierByName(suppliers, supplierNameRaw);
        if (!found) { unmatched.add(supplierNameRaw); return; }
        if (!found.sheetName) { noTemplate.add(found.name); return; }
        const values = {};
        masterFields.forEach((f, idx) => { values[f.id] = row[idx + 1] !== undefined ? row[idx + 1] : ''; });
        if (!Object.values(values).some((v) => String(v).trim() !== '')) return;
        if (!groups[found.id]) groups[found.id] = [];
        groups[found.id].push({ values });
      });

      const newItems = [];
      for (const sId of Object.keys(groups)) {
        const s = suppliers.find((x) => x.id === sId);
        try {
          const { fileName, base64, itemCount } = await buildSupplierFile(s, groups[sId], masterFields, loadSupplierTemplate);
          newItems.push({ id: uid(), supplierId: s.id, supplierName: s.name, email: s.email, fileName, base64, createdAt: Date.now(), itemCount });
        } catch (err) {
          console.error(err);
          noTemplate.add(s.name);
        }
      }

      if (newItems.length) setGenList((prev) => [...newItems, ...prev]);

      if (newItems.length && unmatched.size === 0 && noTemplate.size === 0) {
        showToast(`${newItems.length}개 매입처 발주서를 자동 생성했습니다.`);
      } else {
        const parts = [];
        if (newItems.length) parts.push(`${newItems.length}곳 생성 완료`);
        if (unmatched.size) parts.push(`매칭 안 됨: ${[...unmatched].join(', ')}`);
        if (noTemplate.size) parts.push(`양식 없음: ${[...noTemplate].join(', ')}`);
        showToast(parts.join(' / '), unmatched.size || noTemplate.size ? 'err' : 'ok');
      }
    } catch (err) {
      console.error(err);
      showToast('통합 발주 파일을 읽는 중 오류가 발생했습니다.', 'err');
    } finally {
      setBatchBusy(false);
      if (masterFileRef.current) masterFileRef.current.value = '';
    }
  };

  const redownload = (item) => {
    const buf = base64ToArrayBuffer(item.base64);
    downloadBlob(new Blob([buf], { type: 'application/octet-stream' }), item.fileName);
  };

  const downloadAllCombined = () => {
    if (genList.length === 0) return;
    const wb = XLSX.utils.book_new();
    const usedNames = new Set();
    genList.forEach((item) => {
      const src = XLSX.read(base64ToArrayBuffer(item.base64), { type: 'array', cellStyles: true });
      const sheetName0 = src.SheetNames[0];
      let name = item.supplierName.slice(0, 28).replace(/[\\/?*[\]:]/g, '');
      let final = name || 'Sheet';
      let n = 1;
      while (usedNames.has(final)) { final = `${name}(${++n})`; }
      usedNames.add(final);
      XLSX.utils.book_append_sheet(wb, src.Sheets[sheetName0], final);
    });
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
    downloadBlob(new Blob([out], { type: 'application/octet-stream' }), `${today()}_발주서_전체.xlsx`);
  };

  return (
    <div className="space-y-6">
      <div>
        <SectionHeading title="통합 양식으로 일괄 생성" desc="A열에 매입처명, 그 다음 열부터 통합 항목 순서로 채운 발주 파일을 올리면, 매입처별로 자동 매칭되어 여러 발주서가 한 번에 생성됩니다." />
        <div className="bg-white border border-slate-200 rounded-md p-4 flex flex-wrap items-center gap-2">
          <button onClick={downloadMasterTemplate} className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 border border-slate-300 rounded-md text-slate-600 hover:bg-slate-50">
            <Download size={13} /> 통합 발주 양식 다운로드
          </button>
          <button onClick={() => masterFileRef.current?.click()} disabled={batchBusy} className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 bg-red-700 text-white rounded-md hover:bg-red-800 disabled:opacity-50">
            <Upload size={13} /> {batchBusy ? '생성 중…' : '통합 발주 파일 업로드 → 자동 생성'}
          </button>
          <input ref={masterFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleMasterUpload} />
          <span className="text-xs text-slate-400">매입처명은 등록된 매입처명과 일치(또는 일부 포함)해야 자동 매칭됩니다.</span>
        </div>
      </div>

      <div>
        <SectionHeading title="매입처 1곳씩 직접 입력" desc="매입처를 선택하고 발주 품목을 입력하면, 등록된 양식에 맞춰 셀 위치대로 채워진 발주서를 생성합니다." />

        {ready.length === 0 ? (
          <div className="text-sm text-slate-400 border border-dashed border-slate-200 rounded-md py-8 text-center">
            먼저 '매입처 관리'에서 매입처를 등록하고 엑셀 양식을 업로드해주세요.
          </div>
        ) : (
          <div className="bg-white border border-slate-200 rounded-md p-4">
            <div className="flex items-center gap-3 mb-4">
              <label className="text-xs text-slate-500">매입처</label>
              <select value={supplierId} onChange={(e) => setSupplierId(e.target.value)} className="text-sm px-3 py-2 border border-slate-300 rounded-md outline-none focus:border-slate-500">
                {ready.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              {supplier && !supplier.email && (
                <span className="text-xs text-amber-600 flex items-center gap-1"><AlertCircle size={12} /> 이메일 미등록 — 메일 발송 불가</span>
              )}
            </div>

            {mappedFields.length === 0 ? (
              <div className="text-sm text-slate-400 py-6 text-center">이 매입처는 매칭된 항목이 없습니다. 매입처 관리에서 매칭을 확인해주세요.</div>
            ) : (
              <>
                <div className="border border-slate-200 rounded-md overflow-x-auto mb-3">
                  <table className="w-full text-sm min-w-[600px]">
                    <thead>
                      <tr className="bg-slate-50 text-slate-500 text-xs">
                        {mappedFields.map((f) => <th key={f.id} className="text-left font-medium px-3 py-2 whitespace-nowrap">{f.name}</th>)}
                        <th className="w-10"></th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {rows.map((row) => (
                        <tr key={row.id}>
                          {mappedFields.map((f) => (
                            <td key={f.id} className="px-2 py-1.5">
                              <input
                                value={row.values[f.id]}
                                onChange={(e) => setCell(row.id, f.id, e.target.value)}
                                className="w-full text-sm px-2 py-1.5 border border-slate-200 rounded-md outline-none focus:border-slate-500"
                              />
                            </td>
                          ))}
                          <td className="px-2">
                            <IconBtn onClick={() => removeRow(row.id)} title="행 삭제" danger><X size={13} /></IconBtn>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="flex justify-between">
                  <button onClick={addRow} className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 border border-slate-300 rounded-md text-slate-600 hover:bg-slate-50">
                    <Plus size={13} /> 품목 추가
                  </button>
                  <button onClick={generate} className="flex items-center gap-1.5 px-4 py-2 bg-red-700 text-white text-sm rounded-md hover:bg-red-800">
                    <Stamp size={14} /> 발주서 생성 · 다운로드
                  </button>
                </div>
              </>
            )}
          </div>
        )}
      </div>

      {genList.length > 0 && (
        <div>
          <SectionHeading title="생성된 발주서" desc="개별로 받거나, 매입처별 시트가 담긴 파일 하나로 한번에 받을 수 있습니다." compact />
          <div className="flex justify-end mb-2">
            <button onClick={downloadAllCombined} className="flex items-center gap-1.5 text-xs px-3 py-1.5 border border-slate-300 rounded-md text-slate-700 hover:bg-slate-50">
              <Download size={13} /> 전체 한번에 받기 (매입처별 시트 1개 파일)
            </button>
          </div>
          <div className="border border-slate-200 rounded-md overflow-hidden bg-white">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-slate-500 text-xs">
                  <th className="text-left font-medium px-3 py-2">매입처</th>
                  <th className="text-left font-medium px-3 py-2">파일명</th>
                  <th className="text-left font-medium px-3 py-2">품목수</th>
                  <th className="text-right font-medium px-3 py-2">작업</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {genList.map((item) => (
                  <tr key={item.id}>
                    <td className="px-3 py-2.5 font-medium">{item.supplierName}</td>
                    <td className="px-3 py-2.5 text-slate-500 font-mono text-xs">{item.fileName}</td>
                    <td className="px-3 py-2.5 text-slate-500">{item.itemCount}</td>
                    <td className="px-3 py-2.5">
                      <div className="flex justify-end gap-1.5">
                        <button onClick={() => redownload(item)} className="flex items-center gap-1 text-xs px-2 py-1 border border-slate-300 rounded text-slate-600 hover:bg-slate-50">
                          <Download size={12} /> 다운로드
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="text-xs text-slate-400 mt-2">
            ※ 생성된 발주서를 매입처에 메일로 보내려면 '메일 발송' 탭을 이용해주세요.
          </p>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Tab 4: 메일 발송                                                     */
/* ------------------------------------------------------------------ */

function MailTab({ genList, sendLog, setSendLog, showToast }) {
  const mailDraft = (item) => {
    if (!item.email) { showToast('이 매입처는 등록된 이메일이 없습니다.', 'err'); return null; }
    const subject = `${today()} ${item.supplierName} 발주서`;
    const body = `안녕하세요, ${item.supplierName} 담당자님.\n\n발주서를 보내드립니다. 확인 부탁드립니다.\n첨부파일: ${item.fileName}\n\n감사합니다.`;
    return { subject, body, to: item.email };
  };

  const openMail = (item) => {
    const d = mailDraft(item);
    if (!d) return;
    window.open(`mailto:${encodeURIComponent(d.to)}?subject=${encodeURIComponent(d.subject)}&body=${encodeURIComponent(d.body)}`, '_blank');
  };

  const markSent = (item) => {
    const d = mailDraft(item);
    setSendLog([{ id: uid(), supplierName: item.supplierName, email: item.email, fileName: item.fileName, subject: d ? d.subject : '', sentAt: Date.now() }, ...sendLog]);
    showToast(`${item.supplierName} 발송 이력에 기록했습니다.`);
  };

  const [sendingId, setSendingId] = useState(null);
  const sendReal = async (item) => {
    const d = mailDraft(item);
    if (!d) return;
    setSendingId(item.id);
    try {
      const res = await fetch('/api/send-mail', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ to: d.to, subject: d.subject, body: d.body, fileName: item.fileName, base64: item.base64 }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || '메일 발송에 실패했습니다.');
      setSendLog([{ id: uid(), supplierName: item.supplierName, email: item.email, fileName: item.fileName, subject: d.subject, sentAt: Date.now() }, ...sendLog]);
      showToast(`${item.supplierName}에게 메일을 발송했습니다.`);
    } catch (err) {
      console.error(err);
      showToast(err.message || '메일 발송 중 오류가 발생했습니다.', 'err');
    } finally {
      setSendingId(null);
    }
  };

  return (
    <div className="space-y-6">
      <div>
        <SectionHeading
          title="메일 발송"
          desc="'발주서 생성' 탭에서 만든 발주서를 매입처 담당자에게 보낼 메일 초안을 열고, 발송 이력을 남길 수 있습니다."
        />

        {genList.length === 0 ? (
          <div className="text-sm text-slate-400 border border-dashed border-slate-200 rounded-md py-8 text-center">
            아직 생성된 발주서가 없습니다. 먼저 '발주서 생성' 탭에서 발주서를 만들어주세요.
          </div>
        ) : (
          <>
            <div className="border border-slate-200 rounded-md overflow-hidden bg-white">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs">
                    <th className="text-left font-medium px-3 py-2">매입처</th>
                    <th className="text-left font-medium px-3 py-2">받는 이메일</th>
                    <th className="text-left font-medium px-3 py-2">파일명</th>
                    <th className="text-right font-medium px-3 py-2">작업</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {genList.map((item) => (
                    <tr key={item.id}>
                      <td className="px-3 py-2.5 font-medium">{item.supplierName}</td>
                      <td className="px-3 py-2.5 text-slate-500">
                        {item.email || <span className="text-amber-600">미등록</span>}
                      </td>
                      <td className="px-3 py-2.5 text-slate-500 font-mono text-xs">{item.fileName}</td>
                      <td className="px-3 py-2.5">
                        <div className="flex justify-end gap-1.5">
                          <button onClick={() => openMail(item)} className="flex items-center gap-1 text-xs px-2 py-1 border border-slate-300 rounded text-slate-600 hover:bg-slate-50">
                            <Mail size={12} /> 메일 초안
                          </button>
                          <button onClick={() => markSent(item)} className="flex items-center gap-1 text-xs px-2 py-1 border border-slate-300 rounded text-slate-600 hover:bg-slate-50">
                            <Check size={12} /> 발송 완료 처리
                          </button>
                          <button
                            onClick={() => sendReal(item)}
                            disabled={sendingId === item.id}
                            className="flex items-center gap-1 text-xs px-2 py-1 bg-red-700 text-white rounded hover:bg-red-800 disabled:opacity-50"
                          >
                            <Send size={12} /> {sendingId === item.id ? '발송 중…' : '네이버로 실제 발송'}
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="text-xs text-slate-400 mt-2">
              ※ '메일 초안'은 기본 메일 앱에 제목·본문만 채워 열어주며 첨부파일은 직접 붙여야 합니다. '네이버로 실제 발송'은 서버(Vercel)에 설정된 네이버 계정으로 발주서 파일을 첨부해 바로 메일을 보냅니다 — 배포 시 환경변수(NAVER_EMAIL, NAVER_APP_PASSWORD) 설정이 필요합니다.
            </p>
          </>
        )}
      </div>

      {sendLog.length > 0 && (
        <div>
          <SectionHeading title="발송 이력" compact />
          <div className="border border-slate-200 rounded-md overflow-hidden bg-white">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-slate-500 text-xs">
                  <th className="text-left font-medium px-3 py-2">일시</th>
                  <th className="text-left font-medium px-3 py-2">매입처</th>
                  <th className="text-left font-medium px-3 py-2">제목</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {sendLog.slice(0, 30).map((l) => (
                  <tr key={l.id}>
                    <td className="px-3 py-2 text-slate-500 text-xs">{new Date(l.sentAt).toLocaleString('ko-KR')}</td>
                    <td className="px-3 py-2 font-medium">{l.supplierName}</td>
                    <td className="px-3 py-2 text-slate-500">{l.subject}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Small shared components                                             */
/* ------------------------------------------------------------------ */

function SectionHeading({ title, desc, compact }) {
  return (
    <div className={compact ? 'mb-2' : 'mb-4'}>
      <h2 className="text-sm font-semibold text-slate-800">{title}</h2>
      {desc && <p className="text-xs text-slate-500 mt-1">{desc}</p>}
    </div>
  );
}

function IconBtn({ children, onClick, title, danger }) {
  return (
    <button
      onClick={onClick}
      title={title}
      className={`p-1.5 rounded hover:bg-slate-100 ${danger ? 'text-red-600 hover:bg-red-50' : 'text-slate-500'}`}
    >
      {children}
    </button>
  );
}
