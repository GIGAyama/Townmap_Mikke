// ==========================================================================
// app.jsx — 画面の原本。ここを直す。
//
// 以前は App.html の <script type="text/babel"> に直接書いてあり、
// ブラウザが開くたびに @babel/standalone（約3MB）で翻訳し直していた。
// いまは npm run build がビルド時に 1 回だけ翻訳して app.html を作る。
//
// ⚠️ 生成物（app.html）を直さないこと。次のビルドで消える。
// ==========================================================================
const { useState, useEffect, useCallback, useMemo, useRef, createContext, useContext } = React;

// ==========================================
// System & Utilities
// ==========================================
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

const runGas = (funcName, ...args) => new Promise((resolve, reject) => {
  if (typeof google === 'undefined' || !google.script) { reject(new Error('サーバーに接続できません')); return; }
  google.script.run
    .withSuccessHandler(res => { try { resolve(JSON.parse(res)); } catch (e) { resolve(res); } })
    .withFailureHandler(reject)[funcName](...args);
});

// ── API ラッパー ──
// 束ねられたスプレッドシートが 1 つの学級そのものなので、クラスコードもトークンも
// 渡さない。本人確認はサーバー側の Session.getActiveUser()（Bound.gs）が行う。
// LOCK_BUSY はサーバーの混雑。指数バックオフで 3 回まで自動リトライする。
const callApi = async (name, ...args) => {
  for (let attempt = 0; ; attempt++) {
    const res = await runGas('bd' + name, ...args);
    if (res && res.success === false) {
      if (res.code === 'LOCK_BUSY' && attempt < 3) {
        await sleep(600 * Math.pow(2, attempt) + Math.random() * 400);
        continue;
      }
      const err = new Error(res.error || res.message || 'エラーが発生しました');
      err.code = res.code;
      throw err;
    }
    return res;
  }
};

// 記録系の操作。入口は 1 つで、誰が何をしてよいかはサーバー（Bound.gs）が決める。
const doActionApi = async (action, data) =>
  callApi('ExecuteAction', JSON.stringify({ action, ...data }));

// Enter キー送信の共通ハンドラ。日本語 IME の変換確定 Enter では発火させない
const onEnterKey = (fn) => (e) => {
  if (e.key === 'Enter' && !e.nativeEvent.isComposing) fn();
};

const PIN_COLORS = ['#f43f5e', '#0ea5e9', '#eab308', '#22c55e', '#a855f7'];
const REACTION_EMOJIS = ['👍'];

const isPointInPolygon = (point, vs) => {
  let x = point.x, y = point.y, inside = false;
  for (let i = 0, j = vs.length - 1; i < vs.length; j = i++) {
    let xi = vs[i].x, yi = vs[i].y, xj = vs[j].x, yj = vs[j].y;
    let intersect = ((yi > y) !== (yj > y)) && (x < (xj - xi) * (y - yi) / (yj - yi) + xi);
    if (intersect) inside = !inside;
  }
  return inside;
};

const compressImage = async (file, maxDim = 600, quality = 0.7) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let w = img.width, h = img.height;
        if (w > h) { if (w > maxDim) { h *= maxDim / w; w = maxDim; } } else { if (h > maxDim) { w *= maxDim / h; h = maxDim; } }
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = reject; img.src = e.target.result;
    };
    reader.onerror = reject; reader.readAsDataURL(file);
  });
};

const AppContext = createContext();

// ==========================================
// 画像(Images_画像シート)の遅延読込
//   'imgref:xxx' → api('GetImage') で Data URL に解決（メモリキャッシュ付き）
// ==========================================
const imageCache = new Map();
const resolveImage = (api, src) => {
  if (!src || typeof src !== 'string' || !src.startsWith('imgref:')) return Promise.resolve(src);
  if (imageCache.has(src)) return Promise.resolve(imageCache.get(src));
  const p = api('GetImage', src).then(r => {
    const d = (r && r.dataUrl) || '';
    imageCache.set(src, d);
    return d;
  }).catch(() => { imageCache.delete(src); return ''; });
  imageCache.set(src, p);
  return p;
};

const useResolvedImage = (src) => {
  const { api } = useContext(AppContext);
  const isRef = src && typeof src === 'string' && src.startsWith('imgref:');
  const [url, setUrl] = useState(isRef ? null : src);
  useEffect(() => {
    let alive = true;
    if (src && typeof src === 'string' && src.startsWith('imgref:')) {
      setUrl(null);
      Promise.resolve(resolveImage(api, src)).then(u => { if (alive) setUrl(u); });
    } else {
      setUrl(src);
    }
    return () => { alive = false; };
  }, [src]);
  return url;
};

const SmartImg = ({ src, className, ...props }) => {
  const url = useResolvedImage(src);
  if (!url) return <div className={(className || '') + ' bg-slate-100 animate-pulse'} />;
  return <img src={url} className={className} {...props} />;
};

// 写真アップロードボタン（Drive 不使用。圧縮して Images シートへ保存）
const ImageUploadButton = ({ onDone, maxDim = 600, quality = 0.7, className, children }) => {
  const { api, showToast } = useContext(AppContext);
  const inputRef = useRef(null);
  const [busy, setBusy] = useState(false);
  const handleFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    setBusy(true);
    try {
      let dataUrl = await compressImage(file, maxDim, quality);
      if (dataUrl.length > 380000) dataUrl = await compressImage(file, Math.round(maxDim * 0.7), 0.55);
      if (dataUrl.length > 380000) throw new Error('画像が大きすぎます');
      const res = await api('UploadImage', dataUrl);
      imageCache.set(res.imageRef, dataUrl);
      onDone(res.imageRef);
    } catch (err) {
      showToast(err.message || '画像の保存に失敗しました', 'error');
    }
    setBusy(false);
  };
  return (
    <React.Fragment>
      <button onClick={() => inputRef.current && inputRef.current.click()} disabled={busy} className={className}>
        {busy ? <span className="animate-pulse">保存中...</span> : children}
      </button>
      <input ref={inputRef} type="file" accept="image/*" className="hidden" onChange={handleFile} />
    </React.Fragment>
  );
};

// ==========================================
// UI Components
// ==========================================
const RubyText = ({ text, kana }) => (
  <ruby style={{ rubyPosition: 'over' }}>
    {text}<rt className="text-[0.65em] text-slate-500 opacity-80 font-normal select-none tracking-normal">{kana}</rt>
  </ruby>
);

const LoadingOverlay = ({ label }) => (
  <div className="fixed inset-0 bg-white/90 z-toast flex items-center justify-center backdrop-blur-sm no-print">
    <div className="text-center flex flex-col items-center">
      <div className="relative mb-6 w-16 h-16">
        <div className="absolute bottom-2 left-1/2 -translate-x-1/2 z-10">
          <div className="text-6xl text-brand-500 animate-bounce-pin drop-shadow-md">📍</div>
        </div>
        <div className="absolute -bottom-1 left-1/2 -translate-x-1/2">
          <div className="w-10 h-3 bg-slate-300 rounded-[50%] animate-shadow-pulse"></div>
        </div>
      </div>
      <p className="text-xl font-extrabold text-brand-600 animate-pulse tracking-wider">
        {label || <span><RubyText text="準備" kana="じゅんび" />しています...</span>}
      </p>
    </div>
  </div>
);

const Toast = ({ msg, type = 'success', onHide }) => {
  useEffect(() => { const t = setTimeout(onHide, 3000); return () => clearTimeout(t); }, [msg]);
  if(!msg) return null;
  // センタリング（-translate-x-1/2）と pop-in アニメーションの transform が
  // 打ち消し合わないよう、要素を分ける
  return (
    <div className="fixed bottom-24 sm:bottom-10 left-1/2 -translate-x-1/2 z-toast no-print max-w-[90vw]">
      <div className="animate-pop-in bg-slate-800 text-white px-6 py-3.5 rounded-full shadow-float font-bold flex items-center gap-3">
        <span className="text-xl">{type === 'error' ? '⚠️' : '✨'}</span>
        <span className="text-sm">{msg}</span>
      </div>
    </div>
  );
};

const SvgIcon = {
  Close: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" className="w-5 h-5 stroke-current"><path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" /></svg>,
  Image: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="w-6 h-6"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>,
  Trash: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="w-4 h-4"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>,
  Send: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="w-5 h-5"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
};

// ==========================================
// Feature Components（アプリ本体）
// ==========================================
const PinListModal = ({ pins, onClose, onSelectPin }) => {
  const { state } = useContext(AppContext);
  return (
    <div className="fixed inset-0 bg-slate-900/60 z-[9500] flex items-center justify-center p-4 backdrop-blur-sm animate-pop-in no-print" onClick={onClose} onPointerDown={e=>e.stopPropagation()} onPointerUp={e=>e.stopPropagation()}>
      <div className="bg-white w-full max-w-4xl rounded-[24px] shadow-float overflow-hidden flex flex-col relative max-h-[85vh]" onClick={e => e.stopPropagation()}>
        <div className="bg-brand-500 px-6 py-4 text-white font-bold flex justify-between items-center shrink-0">
          <span className="flex items-center gap-3 text-lg">
            <span className="bg-white/20 p-2 rounded-xl">🔍</span>
            <span>かこんだ<RubyText text="範囲" kana="はんい" />のピン<RubyText text="一覧" kana="いちらん" /> <span className="opacity-80 text-sm ml-1">({pins.length}件)</span></span>
          </span>
          <button onClick={onClose} className="text-white/70 hover:text-white hover:bg-white/20 p-2 rounded-full transition"><SvgIcon.Close /></button>
        </div>
        <div className="flex-1 overflow-y-auto p-4 sm:p-6 bg-surface custom-scrollbar">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
            {pins.map(pin => {
              const authorName = state.users.find(u => u.email === pin.email)?.name || '不明';
              return (
                <div key={pin.pin_id} onClick={() => onSelectPin(pin)} className="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 cursor-pointer hover:shadow-md hover:border-brand-300 hover:-translate-y-1 transition-all flex flex-col gap-3 group">
                  <div className="flex gap-3 items-start">
                    {pin.image_url ? (
                      <SmartImg src={pin.image_url} className="w-20 h-20 rounded-xl object-cover border border-slate-100 shrink-0" />
                    ) : (
                      <div className="w-20 h-20 bg-slate-50 rounded-xl flex items-center justify-center text-3xl shrink-0 border border-slate-100 shadow-inner-soft">
                        {pin.color.startsWith('#') ? '📍' : pin.color}
                      </div>
                    )}
                    <div className="flex-1 min-w-0">
                      <h4 className="font-bold text-slate-800 text-sm truncate mb-1 group-hover:text-brand-600 transition-colors">{pin.title}</h4>
                      <p className="text-xs text-slate-500 line-clamp-3 leading-relaxed">{pin.memo || 'メモはありません'}</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-2 mt-auto pt-3 border-t border-slate-50">
                    {pin.color.startsWith('#') ? <span className="w-3.5 h-3.5 rounded-full shadow-inner" style={{backgroundColor: pin.color}}></span> : <span className="text-sm leading-none">{pin.color}</span>}
                    <span className="text-[11px] font-bold text-slate-600 bg-slate-100 px-2.5 py-1 rounded-full truncate">
                       {authorName}
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
};

// ==========================================
// 学級の管理（コンテナバインド版のみ）
//
// 参加の承認・名簿から外す・受付の開閉・シートの点検。
// どの操作もサーバー側（Bound.gs）で先生かどうかを確かめてから実行される。
// ==========================================
const BoundClassPanel = () => {
  const { showToast, doAction } = useContext(AppContext);
  const [info, setInfo] = useState(null);
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const [selected, setSelected] = useState({});
  const [schema, setSchema] = useState(null);
  const [classNameInput, setClassNameInput] = useState('');

  const load = async () => {
    setError(null);
    try {
      const res = await runGas('bdListMembers');
      if (res && res.success === false) throw new Error(res.error || res.message);
      setInfo(res);
      setClassNameInput(res.className || '');
    } catch (e) { setError(e.message || 'エラーが発生しました'); }
  };
  useEffect(() => { load(); }, []);

  const run = async (action, data, okMsg) => {
    setBusy(true);
    try { await doAction(action, data || {}); if (okMsg) showToast(okMsg); await load(); }
    catch (e) { showToast(e.message || 'エラーが発生しました', 'error'); }
    setBusy(false);
  };

  const checkSchema = async () => {
    setBusy(true);
    try {
      const res = await runGas('bdCheckSchema');
      if (res && res.success === false) throw new Error(res.error || res.message);
      setSchema(res);
    } catch (e) { showToast(e.message || 'エラーが発生しました', 'error'); }
    setBusy(false);
  };

  const repairSchema = async () => {
    setBusy(true);
    try {
      const res = await doAction('repair_schema', {});
      const lines = [].concat(res.done || [], res.skipped || []);
      showToast(lines.length ? '直しました（' + lines.length + ' 件）' : '直すところはありませんでした');
      await checkSchema();
    } catch (e) { showToast(e.message || 'エラーが発生しました', 'error'); }
    setBusy(false);
  };

  if (error) return (
    <div className="max-w-3xl mx-auto w-full p-6 sm:p-10">
      <p className="text-rose-500 font-bold mb-4">{error}</p>
      <button onClick={load} className="bg-slate-800 text-white px-6 py-3 rounded-xl font-bold">もう一度ためす</button>
    </div>
  );
  if (!info) return <div className="p-20 flex justify-center"><div className="animate-spin rounded-full h-12 w-12 border-t-4 border-b-4 border-brand-500"></div></div>;

  const pending = info.members.filter(m => m.status === 'pending');
  const active = info.members.filter(m => m.status === 'active' && m.role !== 'teacher');
  const chosen = Object.keys(selected).filter(k => selected[k]);

  return (
    <div className="max-w-3xl mx-auto w-full p-6 sm:p-10 space-y-6 animate-pop-in">

      {/* 参加の承認 */}
      <section className="bg-white p-8 rounded-3xl shadow-soft border border-slate-100">
        <h3 className="font-extrabold text-slate-800 text-xl mb-2 flex items-center gap-2"><span>🙋</span> 参加の承認（{pending.length} 人）</h3>
        <p className="text-sm text-slate-500 mb-6">承認するまで、その子は地図を見ることも書きこむこともできません。</p>
        {pending.length === 0 ? (
          <p className="text-sm text-slate-400 bg-slate-50 rounded-xl p-4 border border-slate-200">いま待っている人はいません。</p>
        ) : (
          <div className="space-y-2 mb-4">
            {pending.map(m => (
              <label key={m.email} className="flex items-center gap-3 p-3 bg-slate-50 rounded-xl border border-slate-200 cursor-pointer">
                <input type="checkbox" checked={!!selected[m.email]} onChange={e=>setSelected(p=>({...p, [m.email]: e.target.checked}))} className="w-5 h-5 accent-brand-500" />
                <span className="font-bold text-slate-700">{m.displayName || '(名前なし)'}</span>
                {m.number && <span className="text-xs text-slate-400 font-bold">{m.number}番</span>}
                <span className="text-xs text-slate-400 ml-auto break-all">{m.email}</span>
              </label>
            ))}
          </div>
        )}
        {pending.length > 0 && (
          <div className="flex gap-2 flex-wrap">
            <button disabled={busy || !chosen.length} onClick={()=>{ run('approve_members', { emails: chosen }, '承認しました'); setSelected({}); }} className="bg-brand-500 text-white px-6 py-3 rounded-xl font-bold disabled:opacity-50">選んだ人を承認する</button>
            <button disabled={busy} onClick={()=>{ run('approve_members', { emails: pending.map(m=>m.email) }, '全員を承認しました'); setSelected({}); }} className="bg-slate-800 text-white px-6 py-3 rounded-xl font-bold disabled:opacity-50">全員まとめて承認</button>
          </div>
        )}
      </section>

      {/* 参加している人 */}
      <section className="bg-white p-8 rounded-3xl shadow-soft border border-slate-100">
        <h3 className="font-extrabold text-slate-800 text-xl mb-2 flex items-center gap-2"><span>👥</span> 参加している人（{active.length} 人）</h3>
        <p className="text-sm text-slate-500 mb-6">名簿から外しても、その子が書いた記録は消えません（もう一度承認すれば戻れます）。</p>
        <div className="space-y-2">
          {active.map(m => (
            <div key={m.email} className="flex items-center gap-3 p-3 bg-slate-50 rounded-xl border border-slate-200">
              <span className="font-bold text-slate-700">{m.displayName || '(名前なし)'}</span>
              {m.number && <span className="text-xs text-slate-400 font-bold">{m.number}番</span>}
              <span className="text-xs text-slate-400 ml-auto break-all">{m.email}</span>
              <button disabled={busy} onClick={()=>run('remove_member', { email: m.email }, '名簿から外しました')} className="text-xs font-bold text-rose-600 bg-rose-50 border border-rose-200 px-3 py-1.5 rounded-full shrink-0">外す</button>
            </div>
          ))}
          {active.length === 0 && <p className="text-sm text-slate-400 bg-slate-50 rounded-xl p-4 border border-slate-200">まだ誰も参加していません。</p>}
        </div>
      </section>

      {/* 学級の設定 */}
      <section className="bg-white p-8 rounded-3xl shadow-soft border border-slate-100">
        <h3 className="font-extrabold text-slate-800 text-xl mb-6 flex items-center gap-2"><span>⚙️</span> 学級の設定</h3>
        <label className="block text-[11px] font-bold text-slate-500 mb-2 uppercase tracking-wider">学級の名前（児童の画面に出ます）</label>
        <div className="flex gap-2 mb-6">
          <input type="text" value={classNameInput} onChange={e=>setClassNameInput(e.target.value)} placeholder="3年2組" className="flex-1 px-5 py-3 border border-slate-200 rounded-xl bg-slate-50 font-bold focus:ring-2 focus:ring-brand-500 focus:bg-white outline-none transition" />
          <button disabled={busy} onClick={()=>run('set_class_name', { name: classNameInput.trim() }, '学級の名前を保存しました')} className="bg-slate-800 text-white px-6 rounded-xl font-bold disabled:opacity-50">保存</button>
        </div>
        <div className="space-y-3">
          <label className="flex items-center gap-3 p-4 bg-slate-50 rounded-xl border border-slate-200 cursor-pointer">
            <input type="checkbox" checked={info.joinOpen} disabled={busy} onChange={e=>run('set_join_open', { value: e.target.checked })} className="w-5 h-5 accent-brand-500" />
            <span className="font-bold text-slate-700 text-sm">新しい参加を受けつける</span>
          </label>
          <label className="flex items-center gap-3 p-4 bg-slate-50 rounded-xl border border-slate-200 cursor-pointer">
            <input type="checkbox" checked={info.requireApproval} disabled={busy} onChange={e=>run('set_require_approval', { value: e.target.checked })} className="w-5 h-5 accent-brand-500" />
            <span className="font-bold text-slate-700 text-sm">参加には先生の承認を必要にする</span>
          </label>
        </div>
        <p className="text-xs text-slate-400 mt-4 leading-relaxed">先生として登録されているアカウント: <span className="font-mono break-all">{info.teacherEmail || '(未登録)'}</span></p>
      </section>

      {/* シートの点検 */}
      <section className="bg-white p-8 rounded-3xl shadow-soft border border-slate-100">
        <h3 className="font-extrabold text-slate-800 text-xl mb-2 flex items-center gap-2"><span>🩺</span> シートの点検</h3>
        <p className="text-sm text-slate-500 mb-6 leading-relaxed">
          このアプリは、スプレッドシートの列を<strong>見出しの名前</strong>で読み書きします。列を足したり並べ替えたりしても動きますが、
          見出しごと消えると読めなくなります。ここで確かめられます。
          <br />「足りないところを直す」は<strong>足すことしかしません</strong>（列を消したり動かしたりはしません）。
        </p>
        <div className="flex gap-2 flex-wrap mb-4">
          <button disabled={busy} onClick={checkSchema} className="bg-brand-500 text-white px-6 py-3 rounded-xl font-bold disabled:opacity-50">シートを点検する</button>
          <button disabled={busy || !schema || !schema.fixable} onClick={repairSchema} className="bg-slate-800 text-white px-6 py-3 rounded-xl font-bold disabled:opacity-50">足りないところを直す</button>
        </div>
        {schema && (
          <pre className={`text-xs leading-relaxed whitespace-pre-wrap p-4 rounded-xl border ${schema.ok ? 'bg-emerald-50 border-emerald-200 text-emerald-900' : 'bg-amber-50 border-amber-200 text-amber-900'}`}>{schema.report}</pre>
        )}
      </section>
    </div>
  );
};

const TeacherConsole = ({ onClose }) => {
  const { state, dispatch, showToast, ctx, api, doAction } = useContext(AppContext);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [bulkRoster, setBulkRoster] = useState('');
  const [newUnit, setNewUnit] = useState({ name: '', map_name: '基本マップ', url: '' });
  const [newMap, setNewMap] = useState({ name: '', url: '', keepPins: true });
  const [apiKeyInput, setApiKeyInput] = useState('');
  const [customStampsInput, setCustomStampsInput] = useState('');
  const [processing, setProcessing] = useState(false);

  const [selectedStudent, setSelectedStudent] = useState(null);
  const [aiAnalysis, setAiAnalysis] = useState('');
  const [isAnalyzing, setIsAnalyzing] = useState(false);

  const students = state.users.filter(u => u.role !== 'teacher');
  const handlePrint = () => window.print();

  useEffect(() => {
    if (state.activeUnit?.custom_stamps) {
      setCustomStampsInput(state.activeUnit.custom_stamps.join(''));
    } else {
      setCustomStampsInput('📍🐛🌸🚗⚠️🏠❓💡');
    }
  }, [state.activeUnit?.custom_stamps]);

  const executeAction = async (action, data, successMsg, resetCb) => {
    setProcessing(true);
    try {
      await doAction(action, data);
      showToast(successMsg);
      if (resetCb) resetCb();
    } catch(e) { showToast(e.message || 'エラーが発生しました', 'error'); }
    setProcessing(false);
  };

  const handleSaveRoster = () => executeAction('save_users', { users: bulkRoster.split('\n').map(line => { const [email, name, group] = line.split(','); return email && name && group ? { email: email.trim(), name: name.trim(), group_id: group.trim() } : null; }).filter(Boolean) }, '名簿を登録しました', () => setBulkRoster(''));
  const handleSaveUnit = () => executeAction('save_unit', { unit_id: `u_${Date.now()}`, name: newUnit.name, map_name: newUnit.map_name, map_url: newUnit.url }, '新しい単元を開始しました', () => setNewUnit({name:'', map_name: '基本マップ', url:''}));
  const handleAddMap = () => executeAction('add_map', { unit_id: state.activeUnit.unit_id, map_id: `m_${Date.now()}`, name: newMap.name, map_url: newMap.url, copy_from_map_id: newMap.keepPins ? state.activeMapId : null }, '地図を追加しました', () => setNewMap({name:'', url:'', keepPins: true}));

  const toggleSetting = async (field) => {
    if(!state.activeUnit) return;
    const prevVal = state.activeUnit[field] !== false;
    const newVal = !prevVal;
    dispatch(p => ({...p, activeUnit: {...p.activeUnit, [field]: newVal}}));
    try {
      await doAction(field === 'chat_enabled' ? 'toggle_chat' : 'toggle_stamp', { unit_id: state.activeUnit.unit_id, [field]: newVal });
      showToast(`設定を${newVal ? 'ON' : 'OFF'}にしました`);
    } catch(e) {
      dispatch(p => ({...p, activeUnit: {...p.activeUnit, [field]: prevVal}}));
      showToast('通信エラー', 'error');
    }
  };

  const handleSaveStamps = () => {
    const stampsArray = Array.from(customStampsInput.replace(/\s+/g, ''));
    if (stampsArray.length === 0) return showToast('スタンプを入力してください', 'error');
    executeAction('update_custom_stamps', { unit_id: state.activeUnit.unit_id, custom_stamps: stampsArray }, 'スタンプを更新しました', () => dispatch(p => ({...p, activeUnit: {...p.activeUnit, custom_stamps: stampsArray}})));
  };

  const handleSaveApiKey = () => executeAction('save_api_key', { api_key: apiKeyInput.trim() }, 'APIキーを保存しました', () => { dispatch(p => ({...p, hasApiKey: true})); setApiKeyInput(''); });

  const handleAnalyzeStudent = async () => {
    if (!selectedStudent) return;
    setIsAnalyzing(true); setAiAnalysis('');
    try {
      // クラスモードでは email を露出させないため uid（匿名ID）で指定し、サーバーが解決する
      // email は児童の画面に出さないため、uid（匿名ID）で指定してサーバーが解決する
      const payload = { unit_id: state.activeUnit.unit_id, uid: selectedStudent.email };
      const res = await api('GenerateAIPortfolio', JSON.stringify(payload));
      setAiAnalysis(res.portfolio);
    } catch(e) { showToast(e.message || '分析エラー', 'error'); }
    setIsAnalyzing(false);
  };

  return (
    <div className="fixed inset-0 bg-slate-900/60 z-modal flex justify-end print-full" onPointerDown={e=>e.stopPropagation()} onPointerUp={e=>e.stopPropagation()}>
      <div className="w-full max-w-4xl bg-surface h-full shadow-float animate-slide-in-right flex flex-col relative print-full">

        <div className="bg-slate-800 text-white px-6 py-4 shrink-0 flex justify-between items-center no-print">
          <h2 className="text-xl font-bold flex items-center gap-2"><span className="bg-white/10 p-1.5 rounded-lg">🏫</span> 教師用コンソール</h2>
          <button onClick={onClose} className="text-white/70 hover:text-white hover:bg-white/20 p-2 rounded-full transition"><SvgIcon.Close /></button>
        </div>

        <div className="flex px-2 bg-white border-b border-slate-200 overflow-x-auto shrink-0 custom-scrollbar no-print">
          {[
            { id: 'dashboard', icon: '📊', label: '学習分析' },
            { id: 'unit', icon: '🗺️', label: '単元・設定' },
            { id: 'roster', icon: '👥', label: '名簿登録' },
            { id: 'class', icon: '🏫', label: '学級の管理' },
            { id: 'settings', icon: '⚙️', label: 'AI設定' }
          ].map(tab => (
            <button key={tab.id} onClick={()=>setActiveTab(tab.id)} className={`px-5 py-4 font-bold text-sm whitespace-nowrap border-b-2 transition-colors flex items-center gap-2 ${activeTab===tab.id ? 'border-brand-500 text-brand-600' : 'border-transparent text-slate-400 hover:text-slate-600 hover:bg-slate-50'}`}>
              <span>{tab.icon}</span>{tab.label}
            </button>
          ))}
        </div>

        <div className="flex-1 overflow-y-auto relative flex flex-col print-full">

          {/* 📊 ダッシュボード */}
          {activeTab === 'dashboard' && (
            <div className="flex flex-1 overflow-hidden print-full">
              <div className="w-1/3 border-r border-slate-200 bg-white flex flex-col no-print">
                <div className="p-4 bg-slate-50 font-bold text-slate-600 text-xs border-b border-slate-200 shrink-0 uppercase tracking-wider">児童リスト ({students.length})</div>
                <div className="flex-1 overflow-y-auto p-2 space-y-1 custom-scrollbar">
                  {students.map(s => (
                    <button key={s.email} onClick={() => {setSelectedStudent(s); setAiAnalysis('');}}
                            className={`w-full text-left p-3 rounded-xl transition-all flex items-center gap-3 ${selectedStudent?.email === s.email ? 'bg-brand-50 border border-brand-200 text-brand-700 shadow-sm' : 'hover:bg-slate-50 text-slate-600 border border-transparent'}`}>
                      <span className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold text-white shrink-0 ${selectedStudent?.email === s.email ? 'bg-brand-500' : 'bg-slate-300'}`}>
                        {(s.group_id || '?').charAt(0)}
                      </span>
                      <div className="flex-1 truncate font-bold text-sm">{s.name}</div>
                    </button>
                  ))}
                  {students.length === 0 && <p className="text-slate-400 text-sm p-4 text-center mt-10">まだ児童が参加していません</p>}
                </div>
              </div>

              <div className="flex-1 bg-surface flex flex-col overflow-hidden relative print-full">
                {selectedStudent ? (() => {
                  const studentPins = state.pins.filter(p => p.email === selectedStudent.email);
                  const studentChats = state.chats.filter(c => c.email === selectedStudent.email);

                  return (
                    <div className="flex-1 overflow-y-auto p-8 space-y-8 custom-scrollbar print-full">
                      <div className="flex justify-between items-end">
                        <div>
                          <div className="text-sm font-bold text-brand-600 mb-1">{selectedStudent.group_id}</div>
                          <h2 className="text-3xl font-extrabold text-slate-800">{selectedStudent.name} <span className="text-lg font-normal text-slate-500 ml-1">の活動記録</span></h2>
                        </div>
                        <div className="flex gap-3 no-print">
                          <button onClick={handlePrint} className="bg-white border border-slate-200 text-slate-600 font-bold px-4 py-2 rounded-xl shadow-sm hover:bg-slate-50 transition flex items-center gap-2 text-sm">🖨️ 印刷</button>
                          <div className="bg-white px-5 py-2 rounded-xl shadow-sm border border-slate-200 text-center"><div className="text-[10px] text-slate-400 font-bold uppercase tracking-wide">ピン</div><div className="text-xl font-extrabold text-accent-500">{studentPins.length}</div></div>
                          <div className="bg-white px-5 py-2 rounded-xl shadow-sm border border-slate-200 text-center"><div className="text-[10px] text-slate-400 font-bold uppercase tracking-wide">発言</div><div className="text-xl font-extrabold text-brand-500">{studentChats.length}</div></div>
                        </div>
                      </div>

                      <div className="bg-gradient-to-br from-indigo-500 via-purple-500 to-fuchsia-500 rounded-[24px] p-[2px] shadow-float print-full">
                        <div className="bg-white rounded-[22px] p-6 h-full">
                          <div className="flex justify-between items-center mb-4 no-print">
                            <h3 className="font-bold text-indigo-900 text-lg flex items-center gap-2"><span>✨</span> AI ポートフォリオ生成</h3>
                            {state.hasApiKey ? (
                              <button onClick={handleAnalyzeStudent} disabled={isAnalyzing} className="bg-indigo-50 text-indigo-600 border border-indigo-100 hover:bg-indigo-100 font-bold px-4 py-2 rounded-full text-sm transition shadow-sm">
                                {isAnalyzing ? '分析中...' : (aiAnalysis ? '再分析する' : '分析を実行する')}
                              </button>
                            ) : (
                              <button onClick={() => setActiveTab('settings')} className="bg-accent-50 text-accent-600 font-bold px-4 py-2 rounded-full text-sm animate-pulse hover:bg-accent-100 transition border border-accent-100">
                                APIキーを設定してください
                              </button>
                            )}
                          </div>
                          {isAnalyzing && (
                            <div className="py-12 flex flex-col items-center justify-center text-indigo-400 animate-pulse no-print">
                              <span className="text-5xl mb-4 drop-shadow-sm">🧠</span>
                              <p className="font-bold tracking-wide">活動記録をAIが読み込んでいます...</p>
                            </div>
                          )}
                          {aiAnalysis && !isAnalyzing && (
                            <div className="bg-indigo-50/50 p-5 rounded-xl text-sm font-medium text-slate-700 whitespace-pre-wrap leading-relaxed border border-indigo-100/50">
                              {aiAnalysis}
                            </div>
                          )}
                          {!aiAnalysis && !isAnalyzing && (
                            <div className="py-10 text-center no-print">
                              <p className="text-slate-400 text-sm">{state.hasApiKey ? 'ボタンを押すと、ピンの内容や発言から、この児童の「興味関心」や「良いところ」を自動でまとめます。' : 'AI分析を利用するには、上のボタンからGemini APIキーを設定してください。'}</p>
                            </div>
                          )}
                        </div>
                      </div>

                      <div>
                        <h4 className="font-bold text-slate-700 mb-4 flex items-center gap-2"><span>📍</span> 刺したピン一覧</h4>
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          {studentPins.map(p => (
                            <div key={p.pin_id} className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex items-start gap-4">
                              {p.image_url ? <SmartImg src={p.image_url} className="w-16 h-16 rounded-xl object-cover shrink-0 border border-slate-100" /> : <div className="w-16 h-16 bg-slate-50 rounded-xl flex items-center justify-center text-2xl shrink-0 border border-slate-100 shadow-inner-soft">{p.color.startsWith('#')?'📍':p.color}</div>}
                              <div className="flex-1 min-w-0">
                                <h5 className="font-bold text-slate-800 text-sm truncate mb-1">{p.title}</h5>
                                <p className="text-xs text-slate-500 line-clamp-2 leading-relaxed">{p.memo || 'メモなし'}</p>
                              </div>
                            </div>
                          ))}
                          {studentPins.length === 0 && <p className="text-sm text-slate-400 col-span-2 bg-slate-50 p-4 rounded-xl text-center border border-slate-100">まだピンを刺していません</p>}
                        </div>
                      </div>
                    </div>
                  );
                })() : (
                  <div className="flex-1 flex flex-col items-center justify-center text-slate-400 font-bold no-print">
                    <span className="text-4xl mb-4 opacity-50">👈</span>
                    左のリストから児童を選択してください
                  </div>
                )}
              </div>
            </div>
          )}

          {/* 🗺️ 単元・設定 */}
          {activeTab === 'unit' && (
            <div className="max-w-3xl mx-auto w-full p-6 sm:p-10 space-y-8 animate-pop-in">
              <section>
                <h3 className="font-extrabold text-slate-800 text-lg mb-4 flex items-center gap-2"><span>💬</span> コミュニケーション設定</h3>
                <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                  <div>
                    <p className="font-bold text-slate-800">みんなのひろば（チャット）</p>
                    <p className="text-xs text-slate-500 mt-1">※オフにしてもピンに対するコメントは可能です</p>
                  </div>
                  <button onClick={()=>toggleSetting('chat_enabled')} className={`relative inline-flex h-8 w-14 shrink-0 items-center rounded-full transition-colors focus:outline-none ${state.activeUnit?.chat_enabled ? 'bg-brand-500' : 'bg-slate-200'}`}>
                    <span className={`inline-block h-6 w-6 transform rounded-full bg-white transition-transform shadow-sm ${state.activeUnit?.chat_enabled ? 'translate-x-7' : 'translate-x-1'}`}/>
                  </button>
                </div>
              </section>

              <section>
                <h3 className="font-extrabold text-slate-800 text-lg mb-4 flex items-center gap-2"><span>📍</span> ピンとスタンプの設定</h3>
                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                  <div className="flex items-center justify-between p-5 border-b border-slate-100">
                    <div>
                      <p className="font-bold text-slate-800">絵文字スタンプピンの使用</p>
                      <p className="text-xs text-slate-500 mt-1">オンにすると、ピンの代わりに指定したスタンプを使えます</p>
                    </div>
                    <button onClick={()=>toggleSetting('stamp_enabled')} className={`relative inline-flex h-8 w-14 shrink-0 items-center rounded-full transition-colors focus:outline-none ${state.activeUnit?.stamp_enabled !== false ? 'bg-brand-500' : 'bg-slate-200'}`}>
                      <span className={`inline-block h-6 w-6 transform rounded-full bg-white transition-transform shadow-sm ${state.activeUnit?.stamp_enabled !== false ? 'translate-x-7' : 'translate-x-1'}`}/>
                    </button>
                  </div>

                  {state.activeUnit?.stamp_enabled !== false && (
                    <div className="p-5 bg-slate-50 animate-pop-in">
                      <p className="font-bold text-slate-700 text-sm mb-1">使用するスタンプのカスタマイズ</p>
                      <p className="text-xs text-slate-500 mb-3">ピンの代わりに使える絵文字を連続で入力してください</p>
                      <div className="flex gap-3">
                        <input type="text" value={customStampsInput} onChange={e=>setCustomStampsInput(e.target.value)} placeholder="📍🐛🌸🚗⚠️🏠❓💡" className="flex-1 px-4 py-3 border border-slate-200 rounded-xl bg-white text-xl tracking-widest focus:ring-2 focus:ring-brand-500 focus:border-brand-500 outline-none transition shadow-inner-soft" />
                        <button onClick={handleSaveStamps} disabled={processing} className="bg-slate-800 text-white font-bold px-6 py-3 rounded-xl shadow-sm hover:bg-slate-700 transition whitespace-nowrap">更新</button>
                      </div>
                    </div>
                  )}
                </div>
              </section>

              <section>
                <h3 className="font-extrabold text-slate-800 text-lg mb-4 flex items-center gap-2"><span>🗺️</span> 地図の管理</h3>
                <div className="space-y-4">
                  {state.activeUnit && (
                    <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm border-l-4 border-l-accent-500 relative overflow-hidden group">
                      <div className="absolute top-0 right-0 bg-accent-50 text-accent-600 text-[10px] font-bold px-3 py-1 rounded-bl-xl border-b border-l border-accent-100">追加</div>
                      <h4 className="font-bold text-slate-800 mb-1">今の単元に「新しい地図」を追加</h4>
                      <p className="text-xs text-slate-500 mb-5">予想と結果の比較など、タブで切り替える地図を追加します。</p>

                      <div className="space-y-4">
                        <div>
                          <label className="block text-[11px] font-bold text-slate-500 mb-1 uppercase tracking-wider">タブの名前</label>
                          <input type="text" placeholder="例：探検後マップ" value={newMap.name} onChange={e=>setNewMap({...newMap, name: e.target.value})} className="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-sm font-bold focus:ring-2 focus:ring-brand-500 outline-none transition" />
                        </div>
                        <div>
                          <label className="block text-[11px] font-bold text-slate-500 mb-1 uppercase tracking-wider">背景画像</label>
                          <div className="flex gap-2 items-center">
                            <div className="flex-1 px-4 py-3 border border-slate-200 rounded-xl bg-slate-100 text-sm font-bold text-slate-500">
                              {newMap.url ? '✅ 画像を選択済み' : '未選択'}
                            </div>
                            <ImageUploadButton maxDim={1400} quality={0.75} onDone={(ref)=>setNewMap(p=>({...p, url: ref}))} className="bg-brand-50 text-brand-600 border border-brand-200 font-bold px-4 py-3 rounded-xl hover:bg-brand-500 hover:text-white transition whitespace-nowrap flex items-center gap-2">
                              <span>📁</span> 画像を選ぶ
                            </ImageUploadButton>
                          </div>
                        </div>
                        <label className="flex items-center gap-3 p-4 bg-accent-50/50 border border-accent-100 rounded-xl cursor-pointer hover:bg-accent-50 transition">
                          <input type="checkbox" checked={newMap.keepPins} onChange={e=>setNewMap({...newMap, keepPins: e.target.checked})} className="w-5 h-5 accent-accent-500 rounded cursor-pointer" />
                          <span className="text-sm font-bold text-accent-900">いま表示している地図の「ピン」をそのまま引き継ぐ</span>
                        </label>
                        <button onClick={handleAddMap} disabled={processing} className="w-full bg-accent-500 text-white py-3.5 rounded-xl font-bold hover:bg-accent-600 transition shadow-sm">この地図を追加する</button>
                      </div>
                    </div>
                  )}

                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm border-l-4 border-l-brand-500 opacity-80 hover:opacity-100 transition">
                    <h4 className="font-bold text-slate-800 mb-1">全てリセットして「新しい単元」を作る</h4>
                    <p className="text-xs text-slate-500 mb-5">児童の画面が完全に切り替わり、ゼロからのスタートになります。</p>
                    <div className="space-y-4">
                      <div>
                        <input type="text" placeholder="単元のなまえ（例：町探検）" value={newUnit.name} onChange={e=>setNewUnit({...newUnit, name: e.target.value})} className="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-sm font-bold focus:ring-2 focus:ring-brand-500 outline-none transition" />
                      </div>
                      <div className="flex gap-2 items-center">
                        <div className="flex-1 px-4 py-3 border border-slate-200 rounded-xl bg-slate-100 text-sm font-bold text-slate-500">
                          {newUnit.url ? '✅ 画像を選択済み' : '未選択'}
                        </div>
                        <ImageUploadButton maxDim={1400} quality={0.75} onDone={(ref)=>setNewUnit(p=>({...p, url: ref}))} className="bg-brand-50 text-brand-600 border border-brand-200 font-bold px-4 py-3 rounded-xl hover:bg-brand-500 hover:text-white transition whitespace-nowrap flex items-center gap-2">
                          <span>📁</span> 画像を選ぶ
                        </ImageUploadButton>
                      </div>
                      <button onClick={handleSaveUnit} disabled={processing} className="w-full bg-brand-500 text-white py-3.5 rounded-xl font-bold hover:bg-brand-600 transition shadow-sm">まっさらな状態で単元を開始</button>
                    </div>
                  </div>
                </div>
              </section>
            </div>
          )}

          {/* 👥 名簿登録 */}
          {activeTab === 'roster' && (
            <div className="max-w-3xl mx-auto w-full p-6 sm:p-10 animate-pop-in">
              <div className="bg-white p-8 rounded-3xl shadow-soft border border-slate-100">
                <h3 className="font-extrabold text-slate-800 text-xl mb-2 flex items-center gap-2"><span>👥</span> 児童の登録（名簿）</h3>
                <div className="bg-brand-50 border border-brand-100 rounded-xl p-4 mb-4 text-sm text-brand-900 font-medium">
                  💡 児童は先生が配った URL を開くだけで参加申請ができます。ここでの一括登録は<strong>承認なしで即参加</strong>になります。参加申請の承認は「学級の管理」タブで行ってください。
                </div>
                <p className="text-sm text-slate-500 mb-6 leading-relaxed">スプレッドシートやエクセルからコピーして貼り付けてください。<br/>書式: <code className="bg-slate-100 px-2 py-1 rounded text-brand-600 font-bold">メールアドレス, 表示名, 班名</code></p>

                <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 mb-4 font-mono text-xs text-slate-400">
                  001@school.ed.jp, 探検太郎, 1班<br/>
                  002@school.ed.jp, みっけ花子, 2班
                </div>

                <textarea value={bulkRoster} onChange={e=>setBulkRoster(e.target.value)} className="w-full h-64 p-5 border border-slate-200 rounded-2xl mb-6 bg-white focus:ring-2 focus:ring-brand-500 focus:border-brand-500 outline-none transition shadow-inner-soft font-mono text-sm leading-loose" placeholder="ここに貼り付け..."></textarea>
                <button onClick={handleSaveRoster} disabled={processing} className="w-full bg-brand-600 text-white py-4 rounded-xl font-bold text-lg hover:bg-brand-700 transition shadow-float">名簿を一括登録する</button>
              </div>
            </div>
          )}

          {/* 🏫 学級の管理（コンテナバインド版のみ） */}
          {activeTab === 'class' && <BoundClassPanel />}

          {/* ⚙️ AI設定 */}
          {activeTab === 'settings' && (
            <div className="max-w-3xl mx-auto w-full p-6 sm:p-10 animate-pop-in">
              <div className="bg-white p-8 rounded-3xl shadow-soft border border-slate-100">
                <h3 className="font-extrabold text-slate-800 text-xl mb-2 flex items-center gap-2"><span>🤖</span> AI機能（Gemini API）の設定</h3>
                <p className="text-sm text-slate-500 mb-6">学習分析ダッシュボードで児童の活動をAIに分析させるには、GoogleのGemini APIキーが必要です。キーはこのクラスのスプレッドシート（Settingsシート）にのみ保存されます。</p>

                <div className="mb-8 flex items-center gap-3 p-4 bg-slate-50 rounded-xl border border-slate-200">
                   <span className="text-sm font-bold text-slate-600">現在の状態:</span>
                   {state.hasApiKey ?
                     <span className="px-3 py-1 bg-emerald-100 text-emerald-800 rounded-full text-xs font-bold border border-emerald-200 flex items-center gap-1"><span>✅</span> 設定済み</span> :
                     <span className="px-3 py-1 bg-rose-100 text-rose-800 rounded-full text-xs font-bold border border-rose-200 animate-pulse flex items-center gap-1"><span>❌</span> 未設定</span>}
                </div>

                <div className="space-y-4 mb-8">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-500 mb-2 uppercase tracking-wider">APIキーを入力</label>
                    <input type="password" placeholder="AIzaSy..." value={apiKeyInput} onChange={e=>setApiKeyInput(e.target.value)} className="w-full p-4 border border-slate-200 rounded-xl bg-slate-50 focus:bg-white focus:ring-2 focus:ring-brand-500 outline-none transition font-mono text-lg" />
                  </div>
                  <button onClick={handleSaveApiKey} disabled={processing} className="w-full bg-slate-800 text-white py-4 rounded-xl font-bold hover:bg-slate-900 transition shadow-md">APIキーを保存して有効化</button>
                </div>

                <div className="p-6 bg-brand-50 rounded-2xl text-sm text-brand-900 border border-brand-100">
                   <p className="font-bold text-base mb-3 flex items-center gap-2"><span>💡</span> APIキーの取得方法（無料）</p>
                   <ol className="list-decimal ml-5 space-y-3 font-medium">
                     <li><a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noopener noreferrer" className="underline font-bold text-brand-600 hover:text-brand-800 transition">Google AI Studio</a> にアクセスします。</li>
                     <li>Googleアカウントでログインし、<strong>「Create API key」</strong>をクリックして新しいキーを作成します。</li>
                     <li>作成されたキーをコピーし、上の入力欄に貼り付けて保存してください。</li>
                   </ol>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

// ★ リアクションコンポーネント
const ReactionBar = ({ targetType, targetId }) => {
  const { state, dispatch, doAction } = useContext(AppContext);
  const myReactions = state.reactions.filter(r => r.target_id === targetId && r.target_type === targetType);
  const emojis = REACTION_EMOJIS;

  const toggleReaction = async (emoji) => {
    const existing = myReactions.find(r => r.email === state.user.email && r.emoji === emoji);
    if (existing) {
      dispatch(p => ({...p, reactions: p.reactions.filter(r => r !== existing)}));
    } else {
      const newReaction = { reaction_id: 'r_temp_'+Date.now(), unit_id: state.activeUnit.unit_id, email: state.user.email, target_type: targetType, target_id: targetId, emoji };
      dispatch(p => ({...p, reactions: [...p.reactions, newReaction]}));
    }
    try { await doAction('toggle_reaction', { unit_id: state.activeUnit.unit_id, target_type: targetType, target_id: targetId, emoji }); }
    catch(e) { /* 次回の同期で整合する */ }
  };

  const counts = {};
  myReactions.forEach(r => { counts[r.emoji] = (counts[r.emoji] || 0) + 1; });

  return (
    <div className="flex gap-2 items-center flex-wrap">
      <div className="flex gap-1 bg-slate-50 rounded-full px-2 py-1 border border-slate-200">
        {emojis.map(emoji => {
          const isPressed = myReactions.some(r => r.email === state.user.email && r.emoji === emoji);
          return (
            <button key={emoji} onClick={() => toggleReaction(emoji)} className={`w-8 h-8 flex items-center justify-center rounded-full transition-all duration-200 ${isPressed ? 'bg-brand-100 border border-brand-200 shadow-sm scale-110' : 'hover:bg-slate-200 opacity-60 hover:opacity-100 grayscale hover:grayscale-0'}`}>
              <span className="text-lg leading-none">{emoji}</span>
            </button>
          );
        })}
      </div>
      {Object.entries(counts).map(([emoji, count]) => count > 0 && (
         <div key={emoji} className="text-xs font-bold text-slate-600 bg-white border border-slate-200 px-2.5 py-1 rounded-full shadow-sm flex items-center gap-1">
           <span>{emoji}</span><span className="text-brand-600">{count}</span>
         </div>
      ))}
    </div>
  );
};

// モーダルをパネル等のスタッキングコンテキスト（backdrop-filter 等）の外に出すためのポータル
const Portal = ({ children }) => ReactDOM.createPortal(children, document.body);

// ピンの作成（pos 指定）と編集（pin 指定）を 1 つのフォームで行う
const PinFormModal = ({ pos, pin, onClose }) => {
  const { state, showToast, dispatch, doAction } = useContext(AppContext);
  const isEdit = !!pin;
  const [form, setForm] = useState(isEdit
    ? { color: pin.color, title: String(pin.title || ''), memo: String(pin.memo || ''), imageUrl: pin.image_url || null }
    : { color: PIN_COLORS[0], title: '', memo: '', imageUrl: null });
  const [loading, setLoading] = useState(false);

  const handleSubmit = async () => {
    if (!form.title.trim()) return showToast(<span><RubyText text="名前" kana="なまえ" />をかいてね</span>, 'error');
    setLoading(true);
    try {
      if (isEdit) {
        const changes = { pin_id: pin.pin_id, color: form.color, title: form.title, memo: form.memo, image_url: form.imageUrl || '' };
        await doAction('update_pin', changes);
        dispatch(p => ({...p, pins: p.pins.map(x => x.pin_id === pin.pin_id ? { ...x, ...changes } : x)}));
        showToast('きろくをなおしました！');
      } else {
        const payload = {
          pin_id: `p_${Date.now()}`, unit_id: state.activeUnit.unit_id,
          map_id: state.activeMapId, email: state.user.email, x: pos.x, y: pos.y, color: form.color,
          title: form.title, memo: form.memo, image_url: form.imageUrl || ''
        };
        await doAction('save_pin', payload);
        showToast('ピンをさしました！');
        dispatch(p => ({...p, pins: [...p.pins, payload]}));
      }
      onClose();
    } catch(e) { showToast(e.message || 'エラーがおきました', 'error'); }
    setLoading(false);
  };

  const customStamps = state.activeUnit?.custom_stamps || [];

  return (
    <Portal>
    <div className="fixed inset-0 bg-slate-900/60 z-modal flex items-center justify-center p-4 backdrop-blur-sm animate-pop-in"
         onPointerDown={e=>e.stopPropagation()} onPointerUp={e=>e.stopPropagation()}>
      <div className="bg-white w-full max-w-md rounded-[24px] shadow-float flex flex-col overflow-hidden max-h-[92vh] overflow-y-auto custom-scrollbar">
        <div className="bg-accent-500 px-6 py-4 text-white font-bold flex justify-between items-center relative shrink-0">
          <span className="flex items-center gap-2 text-lg">
            <span className="bg-white/20 p-1.5 rounded-lg leading-none">{isEdit ? '✏️' : '📍'}</span>
            <span>{isEdit ? <span>きろくをなおす</span> : <span>ピンをさす</span>}</span>
          </span>
          <button onClick={onClose} className="text-white/70 hover:text-white hover:bg-white/20 p-2 rounded-full transition"><SvgIcon.Close /></button>
        </div>

        <div className="p-6 space-y-6">
          <div>
            <p className="text-[11px] font-bold text-slate-500 mb-2 uppercase tracking-wider">アイコンをえらぶ</p>
            <div className="flex flex-wrap gap-2 justify-center bg-slate-50 p-4 rounded-2xl border border-slate-100 shadow-inner-soft">
              {PIN_COLORS.map(c => (
                <button key={c} onClick={() => setForm(p=>({...p, color: c}))} className={`w-10 h-10 rounded-full border-[3px] transition-transform ${form.color === c ? 'border-slate-800 scale-110 shadow-md' : 'border-white shadow-sm hover:scale-105'}`} style={{backgroundColor: c}} />
              ))}

              {state.activeUnit?.stamp_enabled !== false && customStamps.length > 0 && (
                <React.Fragment>
                  <div className="w-full h-px bg-slate-200 my-2"></div>
                  {customStamps.map(stamp => (
                    <button key={stamp} onClick={() => setForm(p=>({...p, color: stamp}))} className={`text-2xl w-11 h-11 flex items-center justify-center rounded-full border-[3px] transition-all ${form.color === stamp ? 'border-accent-500 bg-white scale-110 shadow-md' : 'border-transparent hover:bg-slate-200'}`}>
                      {stamp}
                    </button>
                  ))}
                </React.Fragment>
              )}
            </div>
          </div>

          <div className="space-y-4">
            <input type="text" placeholder="みつけたものの なまえ" value={form.title} onChange={e=>setForm(p=>({...p, title: e.target.value}))} className="w-full px-5 py-4 border border-slate-200 rounded-xl font-bold bg-slate-50 focus:bg-white focus:ring-2 focus:ring-accent-500 focus:border-accent-500 outline-none transition shadow-sm" />
            <textarea placeholder="メモ（みつけたこと）" value={form.memo} onChange={e=>setForm(p=>({...p, memo: e.target.value}))} className="w-full px-5 py-4 border border-slate-200 rounded-xl font-medium bg-slate-50 focus:bg-white focus:ring-2 focus:ring-accent-500 focus:border-accent-500 outline-none resize-none transition shadow-sm min-h-[100px] leading-relaxed" />
          </div>

          {!form.imageUrl ? (
            <ImageUploadButton maxDim={600} quality={0.7} onDone={(ref)=>setForm(p=>({...p, imageUrl: ref}))} className="block w-full p-5 border-2 border-dashed border-slate-300 rounded-2xl text-center cursor-pointer bg-slate-50 hover:bg-accent-50 hover:border-accent-300 transition group">
              <span className="flex flex-col items-center justify-center gap-2">
                <span className="text-slate-400 group-hover:text-accent-500 transition-colors"><SvgIcon.Image /></span>
                <span className="text-slate-500 font-bold text-sm group-hover:text-accent-600 transition-colors">
                  <RubyText text="写真" kana="しゃしん" /><span className="inline-block">をつける</span>
                </span>
              </span>
            </ImageUploadButton>
          ) : (
            <div className="relative rounded-2xl overflow-hidden shadow-sm border border-slate-200 group">
              <SmartImg src={form.imageUrl} className="w-full h-40 object-cover" />
              <div className="absolute inset-0 bg-black/20 opacity-0 group-hover:opacity-100 transition-opacity"></div>
              <button onClick={() => setForm(p=>({...p, imageUrl: null}))} className="absolute top-3 right-3 bg-white/90 backdrop-blur-sm text-slate-800 w-8 h-8 flex items-center justify-center rounded-full shadow-md font-bold hover:bg-rose-500 hover:text-white transition"><SvgIcon.Close /></button>
            </div>
          )}
          <button onClick={handleSubmit} disabled={loading} className="w-full bg-accent-500 text-white font-bold text-lg py-4 rounded-xl shadow-float hover:bg-accent-600 active:scale-95 transition-all flex items-center justify-center gap-2">
            {loading ? <span><RubyText text="送信中" kana="そうしんちゅう" />...</span> : (isEdit ? <span>ほぞんする！</span> : <span>けってい！</span>)}
          </button>
        </div>
      </div>
    </div>
    </Portal>
  );
};

const PinDetailModal = ({ pin, onClose, onEdit }) => {
  const { state, showToast, dispatch, doAction, canAdmin } = useContext(AppContext);
  const [isDeleting, setIsDeleting] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [comment, setComment] = useState('');

  const isOwner = state.user.email === pin.email;
  const isOwnerOrTeacher = isOwner || canAdmin;
  const authorInfo = state.users.find(u => u.email === pin.email) || {name: '不明', group_id: '?'};
  const pinComments = state.chats.filter(c => c.target_type === 'pin' && c.target_id === pin.pin_id);

  const handleDelete = async () => {
    setIsDeleting(true);
    try {
      await doAction('delete_pin', { pin_id: pin.pin_id });
      showToast('ピンを削除しました');
      dispatch(p => ({...p, pins: p.pins.filter(x => x.pin_id !== pin.pin_id)}));
      onClose();
    } catch(e) { showToast(e.message || 'エラー', 'error'); setIsDeleting(false); }
  };

  const handleSendComment = async () => {
    if(!comment.trim()) return;
    const payload = { chat_id: `c_${Date.now()}`, unit_id: state.activeUnit.unit_id, email: state.user.email, message: comment, target_type: 'pin', target_id: pin.pin_id };
    setComment('');
    dispatch(p => ({...p, chats: [...p.chats, payload]}));
    try { await doAction('save_chat', payload); }
    catch(e) { dispatch(p => ({...p, chats: p.chats.filter(c => c.chat_id !== payload.chat_id)})); showToast(<span><RubyText text="送信失敗" kana="そうしんしっぱい" /></span>,'error'); }
  };

  return (
    <div className="fixed inset-0 bg-slate-900/60 z-[9500] flex items-center justify-center p-4 backdrop-blur-sm animate-pop-in"
         onClick={onClose} onPointerDown={e=>e.stopPropagation()} onPointerUp={e=>e.stopPropagation()}>
      <div className="bg-white w-full max-w-lg rounded-[24px] shadow-float flex flex-col overflow-hidden relative max-h-[90vh]" onClick={e => e.stopPropagation()}>
        <div className="overflow-y-auto flex-1 custom-scrollbar">
          {pin.image_url ? (
            <div className="w-full h-64 bg-slate-100 relative shrink-0">
              <SmartImg src={pin.image_url} className="w-full h-full object-cover" />
              <button onClick={onClose} className="absolute top-4 right-4 bg-black/40 text-white rounded-full w-10 h-10 flex items-center justify-center hover:bg-black/60 backdrop-blur-md transition"><SvgIcon.Close /></button>
            </div>
          ) : (
            <div className="w-full flex justify-end p-4 absolute top-0 z-10 pointer-events-none">
               <button onClick={onClose} className="bg-slate-100/80 text-slate-500 rounded-full w-10 h-10 flex items-center justify-center hover:bg-slate-200 transition pointer-events-auto backdrop-blur-sm"><SvgIcon.Close /></button>
            </div>
          )}

          <div className={`p-6 sm:p-8 relative ${!pin.image_url ? 'pt-12' : ''}`}>

            <div className="flex items-center gap-2 mb-4">
              {pin.color.startsWith('#') ? <span className="w-4 h-4 rounded-full shadow-inner border border-black/10" style={{backgroundColor: pin.color}}></span> : <span className="text-xl leading-none drop-shadow-sm">{pin.color}</span>}
              <span className="bg-slate-100 text-slate-600 px-3 py-1 rounded-full text-xs font-bold border border-slate-200">{authorInfo.group_id} - {authorInfo.name}</span>
            </div>

            <h3 className="text-2xl sm:text-3xl font-extrabold text-slate-800 mb-4 leading-tight">{pin.title}</h3>

            <div className="bg-accent-50/50 p-5 rounded-2xl border border-accent-100 mb-3 shadow-inner-soft">
              <p className="text-slate-700 whitespace-pre-wrap font-medium text-[15px] leading-relaxed">{pin.memo || '（メモはありません）'}</p>
            </div>

            <div className="flex justify-end mb-8">
              <ReactionBar targetType="pin" targetId={pin.pin_id} />
            </div>

            <div className="border-t border-slate-100 pt-6">
              <h4 className="font-bold text-slate-700 mb-4 text-sm flex items-center gap-2">
                 <span className="bg-slate-100 p-1.5 rounded-lg text-xs">💬</span> コメント ({pinComments.length})
              </h4>
              <div className="space-y-4 mb-2">
                {pinComments.map(c => {
                  const cAuthor = state.users.find(u => u.email === c.email) || {name: '不明'};
                  return (
                    <div key={c.chat_id} className="bg-slate-50 p-4 rounded-2xl text-sm border border-slate-100 flex flex-col items-start">
                      <span className="font-bold text-slate-500 text-[11px] block mb-1">{cAuthor.name}</span>
                      <span className="text-slate-800 font-medium block leading-relaxed">{c.message}</span>
                      <div className="mt-2 scale-90 origin-left"><ReactionBar targetType="chat" targetId={c.chat_id} /></div>
                    </div>
                  );
                })}
                {pinComments.length === 0 && <p className="text-sm text-slate-400 text-center py-4 bg-slate-50 rounded-xl border border-dashed border-slate-200">まだコメントはありません。</p>}
              </div>
            </div>

            {isOwnerOrTeacher && (
              <div className="flex justify-between items-center gap-2 pt-6 mt-4 border-t border-slate-100">
                {isOwner && onEdit ? (
                  <button onClick={() => onEdit(pin)} className="text-brand-600 font-bold text-xs sm:text-sm bg-brand-50 px-4 py-2 rounded-full hover:bg-brand-100 transition flex items-center gap-1.5">
                    ✏️ なおす（<RubyText text="編集" kana="へんしゅう" />）
                  </button>
                ) : <span />}
                {!confirmDelete ? (
                  <button onClick={() => setConfirmDelete(true)} className="text-rose-500 font-bold text-xs sm:text-sm bg-rose-50 px-4 py-2 rounded-full hover:bg-rose-100 transition flex items-center gap-1.5">
                    <SvgIcon.Trash /> <RubyText text="削除" kana="さくじょ" />する
                  </button>
                ) : (
                  <div className="flex items-center gap-2 bg-rose-50 p-1.5 rounded-full border border-rose-100 animate-pop-in">
                    <span className="text-xs font-bold text-rose-600 pl-3 pr-1"><RubyText text="本当" kana="ほんとう" />に<RubyText text="消" kana="け" />す？</span>
                    <button onClick={() => setConfirmDelete(false)} className="px-3 py-1.5 bg-white text-slate-600 border border-slate-200 rounded-full text-xs font-bold hover:bg-slate-50">やめる</button>
                    <button onClick={handleDelete} disabled={isDeleting} className="px-4 py-1.5 bg-rose-500 text-white rounded-full text-xs font-bold shadow-sm hover:bg-rose-600">はい</button>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>

        <div className="p-4 sm:p-5 border-t border-slate-200 bg-surface shrink-0 flex gap-2">
          <input type="text" value={comment} onChange={e=>setComment(e.target.value)} onKeyDown={onEnterKey(handleSendComment)} placeholder="コメントをかく..." className="flex-1 bg-white border border-slate-200 rounded-full px-5 py-3 outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500 font-medium text-sm transition shadow-sm" />
          <button onClick={handleSendComment} className="bg-brand-500 text-white w-12 h-12 rounded-full flex items-center justify-center shadow-md hover:bg-brand-600 transition shrink-0"><SvgIcon.Send /></button>
        </div>
      </div>
    </div>
  );
};

const PinIcon = ({ color, onClick, isNew }) => {
  const isEmoji = !color.startsWith('#');
  // isNew のアニメーションは内側のラッパーに掛ける（外側の -translate-x/y による
  // 位置合わせが keyframes の transform に上書きされてズレるのを防ぐ）
  return (
    <div onClick={onClick} className={`pin-element absolute -translate-x-1/2 -translate-y-full cursor-pointer transition-transform duration-200 ${isNew ? '' : 'hover:scale-110 hover:-translate-y-full active:scale-95'} drop-shadow-xl hover:drop-shadow-2xl z-10`} style={{ width: '52px', height: '52px', pointerEvents: 'auto', color: isEmoji ? '#333' : color }}>
    <div className={`w-full h-full flex items-end justify-center ${isNew ? 'animate-pop-in' : ''}`}>
      {isEmoji ? (
        <div className="w-11 h-11 bg-white rounded-full shadow-md border-2 border-slate-200 flex items-center justify-center text-xl mb-1.5 relative">
          <span className="translate-y-[1px]">{color}</span>
          <div className="absolute -bottom-[9px] left-1/2 -translate-x-1/2 w-0 h-0 border-l-[7px] border-r-[7px] border-t-[10px] border-l-transparent border-r-transparent border-t-slate-200"></div>
          <div className="absolute -bottom-[6px] left-1/2 -translate-x-1/2 w-0 h-0 border-l-[5px] border-r-[5px] border-t-[8px] border-l-transparent border-r-transparent border-t-white"></div>
        </div>
      ) : (
        <svg viewBox="0 0 24 24" fill="currentColor" className="w-full h-full stroke-white stroke-[2px]">
          <path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z" />
        </svg>
      )}
    </div>
    </div>
  );
};

const ChatPanel = ({ isOpen, onClose }) => {
  const { state, showToast, dispatch, doAction, canAdmin } = useContext(AppContext);
  const [msg, setMsg] = useState('');
  const [replyTo, setReplyTo] = useState(null);
  const endRef = useRef(null);

  useEffect(() => { endRef.current?.scrollIntoView({ behavior: 'smooth' }); }, [state.chats, isOpen]);

  const handleSend = async () => {
    if(!msg.trim()) return;
    const payload = { chat_id: `c_${Date.now()}`, unit_id: state.activeUnit.unit_id, email: state.user.email, message: msg, target_type: replyTo ? 'chat' : 'general', target_id: replyTo ? replyTo.chat_id : '' };
    setMsg(''); setReplyTo(null);

    dispatch(p => ({...p, chats: [...p.chats, payload]}));
    try { await doAction('save_chat', payload); }
    catch(e) { dispatch(p => ({...p, chats: p.chats.filter(c => c.chat_id !== payload.chat_id)})); showToast(<span><RubyText text="送信失敗" kana="そうしんしっぱい" /></span>,'error'); }
  };

  if(!isOpen) return null;
  const isChatEnabled = state.activeUnit?.chat_enabled;

  return (
    <div className="absolute right-0 top-0 bottom-0 w-full md:w-96 bg-white/95 backdrop-blur-xl border-l border-white/40 shadow-[-10px_0_30px_rgba(0,0,0,0.05)] flex flex-col z-[8500] animate-slide-in-right"
         onPointerDown={e=>e.stopPropagation()} onPointerUp={e=>e.stopPropagation()}>
      <div className="px-5 py-4 border-b border-slate-100 flex justify-between items-center shrink-0 bg-white/50">
        <h3 className="font-bold text-slate-800 flex items-center gap-2">
          <span className="bg-brand-50 text-brand-600 p-1.5 rounded-lg text-sm">💬</span>
          みんなのひろば
        </h3>
        <button onClick={onClose} className="text-slate-400 hover:text-slate-600 hover:bg-slate-100 p-2 rounded-full transition"><SvgIcon.Close /></button>
      </div>

      <div className="flex-1 overflow-y-auto p-4 sm:p-5 space-y-5 custom-scrollbar">
        {state.chats.map(c => {
          const isMe = c.email === state.user.email;
          const author = state.users.find(u => u.email === c.email) || {name: '不明', group_id: '?'};
          let pinContext = null;
          if(c.target_type === 'pin') {
            const targetPin = state.pins.find(p => p.pin_id === c.target_id);
            pinContext = targetPin ? <span>📍 {targetPin.title} へのコメント</span> : <span>📍 <RubyText text="削除" kana="さくじょ" />されたピンへのコメント</span>;
          }
          let replyContext = null;
          if(c.target_type === 'chat') {
            const targetChat = state.chats.find(x => x.chat_id === c.target_id);
            const tAuthor = targetChat ? (state.users.find(u => u.email === targetChat.email)?.name || '誰か') : '誰か';
            replyContext = <span>↩️ {tAuthor}さんへの<RubyText text="返信" kana="へんしん" /></span>;
          }

          return (
            <div key={c.chat_id} className={`flex flex-col group ${isMe ? 'items-end' : 'items-start'}`}>
              <span className="text-[10px] text-slate-400 mb-1 font-bold">{author.name} ({author.group_id})</span>
              {(pinContext || replyContext) && (
                <div className={`text-[10px] font-bold mb-1 px-2.5 py-1 rounded-full border ${isMe ? 'bg-brand-50 border-brand-100 text-brand-700' : 'bg-slate-100 border-slate-200 text-slate-600'}`}>
                  {pinContext || replyContext}
                </div>
              )}
              <div className={`flex flex-col gap-1.5 ${isMe ? 'items-end' : 'items-start'}`}>
                <div className="flex items-end gap-2">
                  {isMe && (isChatEnabled || canAdmin) && <button onClick={()=>setReplyTo(c)} className="opacity-0 group-hover:opacity-100 text-xs font-bold text-slate-400 hover:text-brand-500 transition"><RubyText text="返信" kana="へんしん" /></button>}
                  <div className={`px-4 py-2.5 rounded-[20px] max-w-[240px] break-words shadow-sm font-medium text-[13px] leading-relaxed ${isMe ? 'bg-brand-500 text-white rounded-tr-sm' : 'bg-white border border-slate-200 text-slate-800 rounded-tl-sm'}`}>
                    {c.message}
                  </div>
                  {!isMe && (isChatEnabled || canAdmin) && <button onClick={()=>setReplyTo(c)} className="opacity-0 group-hover:opacity-100 text-xs font-bold text-slate-400 hover:text-brand-500 transition"><RubyText text="返信" kana="へんしん" /></button>}
                </div>
                <div className={`scale-[0.85] origin-top ${isMe ? 'origin-right' : 'origin-left'}`}><ReactionBar targetType="chat" targetId={c.chat_id} /></div>
              </div>
            </div>
          );
        })}
        <div ref={endRef} />
      </div>

      {isChatEnabled || canAdmin ? (
        <div className="p-4 bg-surface border-t border-slate-200 shrink-0 flex flex-col gap-2 relative">
          {replyTo && (
            <div className="flex justify-between items-center bg-white border border-slate-200 px-3 py-2 rounded-xl text-xs font-bold text-slate-600 shadow-sm animate-pop-in">
              <span>↩️ {(state.users.find(u=>u.email===replyTo.email)?.name || '誰か')} に<RubyText text="返信" kana="へんしん" />中</span>
              <button onClick={()=>setReplyTo(null)} className="hover:text-rose-500 bg-slate-50 rounded-full p-1"><SvgIcon.Close /></button>
            </div>
          )}
          <div className="flex gap-2">
            <input type="text" value={msg} onChange={e=>setMsg(e.target.value)} onKeyDown={onEnterKey(handleSend)} placeholder="メッセージ..." className="flex-1 bg-white border border-slate-200 rounded-full px-5 py-3 outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500 font-medium text-sm shadow-sm transition" />
            <button onClick={handleSend} className="bg-brand-500 text-white w-12 h-12 rounded-full flex items-center justify-center shadow-md hover:bg-brand-600 transition shrink-0"><SvgIcon.Send /></button>
          </div>
        </div>
      ) : (
        <div className="p-5 bg-slate-100 border-t border-slate-200 shrink-0 text-center">
          <span className="text-sm font-bold text-slate-500"><span><RubyText text="現在" kana="げんざい" /></span>、チャットはオフになっています🔇</span>
        </div>
      )}
    </div>
  );
};

// ── わたしのきろく（自分の気づきの一覧・整理・再編集）──
const MyNotesPanel = ({ isOpen, onClose }) => {
  const { state, dispatch } = useContext(AppContext);
  const [editingPin, setEditingPin] = useState(null);
  const [query, setQuery] = useState('');

  if (!isOpen) return null;
  const maps = state.activeUnit?.maps || [];
  const mapName = (id) => maps.find(m => m.id === id)?.name || '';
  const q = query.trim();
  const mine = state.pins
    .filter(p => p.email === state.user.email)
    .filter(p => !q || (String(p.title || '') + String(p.memo || '')).includes(q))
    .slice().reverse(); // 新しい順

  return (
    <div className="absolute right-0 top-0 bottom-0 w-full md:w-96 bg-white/95 backdrop-blur-xl border-l border-white/40 shadow-[-10px_0_30px_rgba(0,0,0,0.05)] flex flex-col z-[8500] animate-slide-in-right"
         onPointerDown={e=>e.stopPropagation()} onPointerUp={e=>e.stopPropagation()}>
      <div className="px-5 py-4 border-b border-slate-100 flex justify-between items-center shrink-0 bg-white/50">
        <h3 className="font-bold text-slate-800 flex items-center gap-2">
          <span className="bg-accent-50 text-accent-600 p-1.5 rounded-lg text-sm">📒</span>
          わたしのきろく <span className="text-xs text-slate-400 font-bold">({mine.length}件)</span>
        </h3>
        <button onClick={onClose} className="text-slate-400 hover:text-slate-600 hover:bg-slate-100 p-2 rounded-full transition"><SvgIcon.Close /></button>
      </div>

      <div className="px-4 pt-3 shrink-0">
        <input type="text" value={query} onChange={e=>setQuery(e.target.value)} placeholder="🔍 ことばでさがす..."
               className="w-full bg-white border border-slate-200 rounded-full px-4 py-2.5 outline-none focus:ring-2 focus:ring-accent-400 font-medium text-sm shadow-sm transition" />
      </div>

      <div className="flex-1 overflow-y-auto p-4 space-y-3 custom-scrollbar">
        {mine.map(pin => (
          <div key={pin.pin_id} className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-sm flex flex-col gap-2">
            <div className="flex gap-3 items-start">
              {pin.image_url ? (
                <SmartImg src={pin.image_url} className="w-14 h-14 rounded-xl object-cover border border-slate-100 shrink-0" />
              ) : (
                <div className="w-14 h-14 bg-slate-50 rounded-xl flex items-center justify-center text-2xl shrink-0 border border-slate-100 shadow-inner-soft">
                  {String(pin.color || '').startsWith('#') ? '📍' : pin.color}
                </div>
              )}
              <div className="flex-1 min-w-0">
                <h4 className="font-bold text-slate-800 text-sm truncate">{pin.title}</h4>
                <p className="text-xs text-slate-500 line-clamp-2 leading-relaxed">{pin.memo || 'メモはありません'}</p>
                {maps.length > 1 && <span className="inline-block mt-1 text-[10px] font-bold text-slate-400 bg-slate-100 px-2 py-0.5 rounded-full">🗺️ {mapName(pin.map_id)}</span>}
              </div>
            </div>
            <div className="flex gap-2 justify-end pt-1 border-t border-slate-50">
              <button onClick={() => { if (pin.map_id !== state.activeMapId) dispatch(p=>({...p, activeMapId: pin.map_id})); onClose(); }}
                      className="text-[11px] font-bold text-slate-500 bg-slate-100 hover:bg-slate-200 px-3 py-1.5 rounded-full transition">
                🗺️ <RubyText text="地図" kana="ちず" />で<RubyText text="見" kana="み" />る
              </button>
              <button onClick={() => setEditingPin(pin)}
                      className="text-[11px] font-bold text-brand-600 bg-brand-50 hover:bg-brand-100 px-3 py-1.5 rounded-full transition">
                ✏️ なおす
              </button>
            </div>
          </div>
        ))}
        {mine.length === 0 && (
          <div className="text-center py-12 text-slate-400 text-sm font-bold bg-slate-50 rounded-2xl border border-dashed border-slate-200">
            {q ? 'みつかりませんでした' : <span>まだきろくがないよ。<br/><RubyText text="地図" kana="ちず" />をタップしてピンをさそう！</span>}
          </div>
        )}
      </div>

      {editingPin && <PinFormModal pin={editingPin} onClose={() => setEditingPin(null)} />}
    </div>
  );
};

const MapArea = () => {
  const { state, dispatch, showToast, canAdmin } = useContext(AppContext);
  const [creatingPos, setCreatingPos] = useState(null);
  const [selectedPin, setSelectedPin] = useState(null);
  const [editingPin, setEditingPin] = useState(null);
  const [mode, setMode] = useState(state.user.role === 'teacher' ? 'view' : 'edit');

  const [isDrawing, setIsDrawing] = useState(false);
  const [lassoPoints, setLassoPoints] = useState([]);
  const [areaPins, setAreaPins] = useState(null);

  // 地図画像を表示領域いっぱいにフィットさせるための実測値
  const [viewport, setViewport] = useState({ w: 0, h: 0 });
  const [natural, setNatural] = useState(null);

  const outerRef = useRef(null);      // 地図の表示領域（画面いっぱい）
  const mapRef = useRef(null);        // 地図 <img>
  const containerRef = useRef(null);  // transform を掛けるコンテナ
  // パン/ズームは毎フレームの再レンダリングを避けるため ref + 直接 style 反映で管理する
  const view = useRef({ x: 0, y: 0, scale: 1 });
  const pointers = useRef(new Map());
  const gesture = useRef(null);       // { type: 'drag' | 'pinch', ... }
  const suppressTap = useRef(false);  // ピンチ後の指離しをタップ扱いにしない

  const maps = state.activeUnit?.maps || [];
  const currentMapInfo = maps.find(m => m.id === state.activeMapId) || maps[0];

  const visiblePins = state.pins.filter(p => {
    if(p.map_id !== state.activeMapId) return false;
    if(state.filter.scope === 'mine' && p.email !== state.user.email) return false;
    if(state.filter.scope === 'group' && state.users.find(u=>u.email === p.email)?.group_id !== state.user.group_id) return false;
    if(state.filter.color !== 'all' && p.color !== state.filter.color) return false;
    return true;
  });

  // 表示領域の実寸を追跡（端末の回転・リサイズ・パネル開閉に追従）
  useEffect(() => {
    const el = outerRef.current;
    if (!el) return;
    const update = () => setViewport({ w: el.clientWidth, h: el.clientHeight });
    update();
    if (typeof ResizeObserver !== 'undefined') {
      const ro = new ResizeObserver(update);
      ro.observe(el);
      return () => ro.disconnect();
    }
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, []);

  const url = currentMapInfo?.url || '';
  const resolvedUrl = useResolvedImage(url);
  const displayUrl = !resolvedUrl ? '' :
    resolvedUrl.startsWith('data:') ? resolvedUrl :
    resolvedUrl.includes('drive') ? resolvedUrl :
    `https://drive.google.com/thumbnail?sz=w1500&id=${resolvedUrl.match(/id=([^&]+)/)?.[1]||resolvedUrl}`;

  // 地図の切り替え時は寸法とビューをリセット。
  // 画像がこの effect より先に読み込み完了していると onLoad は再発火しないため、
  // 読み込み済みならここで寸法を確定する（フィットが効かないままになる競合の防止）
  useEffect(() => {
    view.current = { x: 0, y: 0, scale: 1 };
    applyView();
    const img = mapRef.current;
    if (img && img.complete && img.naturalWidth) {
      setNatural({ w: img.naturalWidth, h: img.naturalHeight });
    } else {
      setNatural(null);
    }
  }, [displayUrl]);

  // 表示領域に対して余白最小の contain フィット寸法（scale=1 で画面最大表示になる）
  const base = useMemo(() => {
    if (!natural || !viewport.w || !viewport.h) return null;
    const PAD = 10;
    const f = Math.min((viewport.w - PAD * 2) / natural.w, (viewport.h - PAD * 2) / natural.h);
    return { w: Math.max(Math.round(natural.w * f), 50), h: Math.max(Math.round(natural.h * f), 50) };
  }, [natural, viewport.w, viewport.h]);

  const applyView = () => {
    if (!containerRef.current) return;
    const v = view.current;
    containerRef.current.style.transform = `translate3d(${v.x}px, ${v.y}px, 0) scale(${v.scale})`;
  };

  const clampView = () => {
    const v = view.current;
    v.scale = Math.min(Math.max(v.scale, 1), 6);
    if (base) {
      // 地図が画面外へ完全に消えないようにパン量を制限
      const maxX = (base.w * v.scale + viewport.w) / 2 - 60;
      const maxY = (base.h * v.scale + viewport.h) / 2 - 60;
      v.x = Math.min(Math.max(v.x, -maxX), maxX);
      v.y = Math.min(Math.max(v.y, -maxY), maxY);
    }
  };

  // 指定した画面座標を支点にズーム（ピンチ・ホイール・ボタン共通）
  const zoomAt = (clientX, clientY, factor) => {
    const el = outerRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
    const v = view.current;
    const ns = Math.min(Math.max(v.scale * factor, 1), 6);
    const px = (clientX - cx - v.x) / v.scale;
    const py = (clientY - cy - v.y) / v.scale;
    view.current = { scale: ns, x: clientX - cx - ns * px, y: clientY - cy - ns * py };
    clampView(); applyView();
  };

  const zoomCenter = (factor) => {
    const el = outerRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    zoomAt(r.left + r.width / 2, r.top + r.height / 2, factor);
  };

  const resetView = () => {
    view.current = { x: 0, y: 0, scale: 1 };
    if (containerRef.current) containerRef.current.style.transition = 'transform 0.15s ease-out';
    applyView();
  };

  // トラックパッド・マウスホイールでのズーム（React の onWheel は preventDefault
  // できない環境があるため、passive:false のネイティブリスナーで登録する）
  useEffect(() => {
    const el = outerRef.current;
    if (!el) return;
    const onWheel = (e) => { e.preventDefault(); zoomAt(e.clientX, e.clientY, Math.exp(-e.deltaY * 0.0015)); };
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, [base, viewport.w, viewport.h]);

  const getPos = (e) => {
    if(!mapRef.current) return null;
    const r = mapRef.current.getBoundingClientRect();
    if (e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom) return null;
    return { x: ((e.clientX - r.left) / r.width) * 100, y: ((e.clientY - r.top) / r.height) * 100 };
  };

  const onPointerDown = (e) => {
    if(e.target.closest('button') || e.target.closest('.pin-element')) return;
    // 新しい操作の開始（primary pointer）時に、pointerup が届かず取りこぼした
    // 古いポインタ情報を掃除する（残っているとタップがピンチと誤判定され、
    // 以後ピンがさせなくなる）
    if (e.isPrimary) { pointers.current.clear(); gesture.current = null; suppressTap.current = false; }
    try { e.currentTarget.setPointerCapture(e.pointerId); } catch (err) {}
    pointers.current.set(e.pointerId, { x: e.clientX, y: e.clientY });

    if (pointers.current.size === 2) {
      // 2 本指 → ピンチズーム開始（なげなわ・ドラッグは中断）
      setIsDrawing(false); setLassoPoints([]);
      const [p1, p2] = [...pointers.current.values()];
      gesture.current = {
        type: 'pinch',
        dist: Math.hypot(p1.x - p2.x, p1.y - p2.y) || 1,
        scale: view.current.scale, x: view.current.x, y: view.current.y,
        mid: { x: (p1.x + p2.x) / 2, y: (p1.y + p2.y) / 2 }
      };
      suppressTap.current = true;
      if (containerRef.current) containerRef.current.style.transition = 'none';
      return;
    }

    if (mode === 'lasso') {
      setIsDrawing(true);
      const pos = getPos(e);
      setLassoPoints(pos ? [pos] : []);
    } else {
      gesture.current = { type: 'drag', x: view.current.x, y: view.current.y, mx: e.clientX, my: e.clientY, moved: false };
      if(containerRef.current) containerRef.current.style.transition = 'none';
    }
  };

  const onPointerMove = (e) => {
    if (pointers.current.has(e.pointerId)) {
      pointers.current.set(e.pointerId, { x: e.clientX, y: e.clientY });
    }
    const g = gesture.current;

    if (g && g.type === 'pinch') {
      if (pointers.current.size < 2) return;
      const [p1, p2] = [...pointers.current.values()];
      const el = outerRef.current;
      if (!el) return;
      const r = el.getBoundingClientRect();
      const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
      const dist = Math.hypot(p1.x - p2.x, p1.y - p2.y) || 1;
      const ns = Math.min(Math.max(g.scale * dist / g.dist, 1), 6);
      const px = (g.mid.x - cx - g.x) / g.scale;
      const py = (g.mid.y - cy - g.y) / g.scale;
      const mid = { x: (p1.x + p2.x) / 2, y: (p1.y + p2.y) / 2 };
      view.current = { scale: ns, x: mid.x - cx - ns * px, y: mid.y - cy - ns * py };
      clampView(); applyView();
      return;
    }

    if (mode === 'lasso') {
      if (isDrawing && pointers.current.has(e.pointerId)) {
        const pos = getPos(e);
        if(pos) setLassoPoints(prev => [...prev, pos]);
      }
      return;
    }

    if (g && g.type === 'drag') {
      const dx = e.clientX - g.mx;
      const dy = e.clientY - g.my;
      // タップ判定のしきい値。指は押している間に数px ぶれるため、タッチ/ペンは広めに取る
      const slop = e.pointerType === 'mouse' ? 5 : 10;
      if (!g.moved && (Math.abs(dx) > slop || Math.abs(dy) > slop)) g.moved = true;
      if (g.moved) {
        view.current.x = g.x + dx;
        view.current.y = g.y + dy;
        clampView(); applyView();
      }
    }
  };

  const onPointerUp = (e) => {
    try { e.currentTarget.releasePointerCapture(e.pointerId); } catch (err) {}
    pointers.current.delete(e.pointerId);
    const g = gesture.current;

    if (g && g.type === 'pinch') {
      if (pointers.current.size < 2) gesture.current = null;
      if (pointers.current.size === 0) suppressTap.current = false;
      return;
    }

    if (mode === 'lasso') {
      if (isDrawing) {
        setIsDrawing(false);
        if (e.type !== 'pointercancel' && lassoPoints.length > 5) {
          const containedPins = visiblePins.filter(p => isPointInPolygon({x: p.x, y: p.y}, lassoPoints));
          if (containedPins.length > 0) {
            setAreaPins(containedPins);
          } else {
            showToast(<span>かこんだ<RubyText text="中" kana="なか" />にピンはありませんでした</span>, 'error');
          }
        }
        setLassoPoints([]);
      }
    } else {
      if(containerRef.current) containerRef.current.style.transition = 'transform 0.1s ease-out';
      const wasTap = g && g.type === 'drag' && !g.moved && !suppressTap.current && e.type !== 'pointercancel';
      gesture.current = null;
      // ボタンやピンの上で指を離した場合はタップ扱いにしない
      // （ズームボタン等のタップでピン作成モーダルが誤って開くバグの修正）
      // 先生もピンをさせる（サーバー側は教員の save_pin に元々対応済み）
      if (wasTap && !e.target.closest('button') && !e.target.closest('.pin-element')
          && mode === 'edit') {
        const pos = getPos(e);
        if(pos) setCreatingPos(pos);
      }
    }
    if (pointers.current.size === 0) suppressTap.current = false;
  };

  return (
    <div ref={outerRef}
         className={`map-viewport flex-1 relative overflow-hidden bg-slate-100 flex items-center justify-center ${mode === 'lasso' ? 'cursor-crosshair' : 'cursor-grab active:cursor-grabbing'}`}
         onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp}>

      {/* センタリング用の -translate-x-1/2 と keyframes アニメーションの transform が
          打ち消し合わないよう、位置決めの外側とアニメーションの内側を分ける */}
      {maps.length > 1 && (
        <div className="absolute top-3 sm:top-5 left-1/2 -translate-x-1/2 z-overlay max-w-[95vw]">
          <div className="animate-pop-in flex gap-1.5 p-1.5 bg-white/80 backdrop-blur-lg rounded-full shadow-float border border-white/60 overflow-x-auto custom-scrollbar">
            {maps.map(m => (
               <button key={m.id} onClick={(e)=>{e.stopPropagation(); dispatch(p=>({...p, activeMapId: m.id}));}}
                       className={`px-4 sm:px-6 py-2 sm:py-2.5 rounded-full font-bold text-[13px] whitespace-nowrap transition-all ${state.activeMapId === m.id ? 'bg-slate-800 text-white shadow-md' : 'text-slate-500 hover:bg-white hover:text-slate-800'}`}>
                 {m.name}
               </button>
            ))}
          </div>
        </div>
      )}

      {mode === 'lasso' && (
        <div className="absolute top-16 sm:top-24 left-1/2 -translate-x-1/2 z-overlay pointer-events-none">
          <span className="animate-pop-in bg-brand-500 text-white px-6 py-3 rounded-full font-bold shadow-float flex items-center gap-2 border-[3px] border-white text-sm whitespace-nowrap">
            <span>✏️</span>
            <span><RubyText text="地図" kana="ちず" />をなぞって かこんでね！</span>
          </span>
        </div>
      )}

      <div className="absolute bottom-4 sm:bottom-8 left-1/2 -translate-x-1/2 z-overlay bg-white/90 backdrop-blur-xl p-1.5 sm:p-2 rounded-full shadow-float border border-white flex gap-1 sm:gap-1.5 no-print">
        <button onClick={(e)=>{ e.stopPropagation(); setMode('view'); }} className={`px-3 py-2.5 sm:px-7 sm:py-3.5 rounded-full font-bold text-xs sm:text-sm transition-all flex items-center gap-1.5 sm:gap-2 whitespace-nowrap ${mode==='view' ? 'bg-slate-800 text-white shadow-md' : 'text-slate-500 hover:bg-slate-100 hover:text-slate-800'}`}>
          <span className="text-base sm:text-lg shrink-0">👀</span>
          <span className="inline-block mt-1">さわる・<RubyText text="見" kana="み" />る</span>
        </button>

        <button onClick={(e)=>{ e.stopPropagation(); setMode('lasso'); }} className={`px-3 py-2.5 sm:px-7 sm:py-3.5 rounded-full font-bold text-xs sm:text-sm transition-all flex items-center gap-1.5 sm:gap-2 whitespace-nowrap ${mode==='lasso' ? 'bg-brand-500 text-white shadow-md' : 'text-slate-500 hover:bg-slate-100 hover:text-brand-600'}`}>
          <span className="text-base sm:text-lg shrink-0">✏️</span>
          <span className="inline-block mt-1">かこんで<RubyText text="見" kana="み" />る</span>
        </button>

        <button onClick={(e)=>{ e.stopPropagation(); setMode('edit'); }} className={`px-3 py-2.5 sm:px-7 sm:py-3.5 rounded-full font-bold text-xs sm:text-sm transition-all flex items-center gap-1.5 sm:gap-2 whitespace-nowrap ${mode==='edit' ? 'bg-accent-500 text-white shadow-md' : 'text-slate-500 hover:bg-slate-100 hover:text-accent-600'}`}>
          <span className="text-base sm:text-lg shrink-0">📍</span>
          <span className="inline-block mt-1">ピンをさす</span>
        </button>
      </div>

      <div className="absolute bottom-24 sm:bottom-8 left-3 sm:left-6 z-overlay flex flex-col gap-2">
        <div className="flex flex-col bg-white/90 backdrop-blur-md rounded-full shadow-float border border-white p-1">
          <button onClick={()=>zoomCenter(1.5)} className="w-10 h-10 sm:w-11 sm:h-11 bg-transparent rounded-full font-bold text-xl text-slate-600 hover:bg-slate-100 transition">+</button>
          <div className="w-full h-px bg-slate-200"></div>
          <button onClick={()=>zoomCenter(1/1.5)} className="w-10 h-10 sm:w-11 sm:h-11 bg-transparent rounded-full font-bold text-xl text-slate-600 hover:bg-slate-100 transition">-</button>
        </div>
        <button onClick={resetView} className="w-12 h-12 sm:w-[52px] sm:h-[52px] mt-1 bg-slate-800 text-white rounded-full shadow-float font-bold text-[11px] hover:bg-black transition"><RubyText text="戻" kana="もど" />す</button>
      </div>

      <div ref={containerRef} className="smooth-map-container relative shadow-[0_20px_50px_-12px_rgba(0,0,0,0.15)] rounded-2xl border-4 border-white bg-white transition-transform duration-100 ease-out"
           style={{ transform: `translate3d(${view.current.x}px, ${view.current.y}px, 0) scale(${view.current.scale})`, width: 'fit-content', height: 'fit-content' }}>

        {url && displayUrl ? <img ref={mapRef} src={displayUrl}
               onLoad={e => setNatural({ w: e.target.naturalWidth || 1, h: e.target.naturalHeight || 1 })}
               style={base ? { width: base.w + 'px', height: base.h + 'px', maxWidth: 'none', maxHeight: 'none' }
                           : { maxWidth: '88vw', maxHeight: '70vh', width: 'auto', height: 'auto' }}
               className="block pointer-events-none rounded-xl" draggable="false" />
             : url ? <div className="w-[60vw] h-[60vh] max-w-2xl max-h-2xl flex items-center justify-center text-slate-400 font-bold bg-slate-50 rounded-2xl animate-pulse">
                 <span className="text-4xl">🗺️</span>
               </div>
             : <div className="w-[60vw] h-[60vh] max-w-2xl max-h-2xl flex items-center justify-center text-slate-400 font-bold bg-slate-50 rounded-2xl border-2 border-dashed border-slate-200">
                 <div className="text-center">
                   <span className="text-6xl block mb-4">🗺️</span>
                   <span><RubyText text="先生" kana="せんせい" />が<RubyText text="地図" kana="ちず" />を<RubyText text="設定" kana="せってい" />するのを<RubyText text="待" kana="ま" />ってね</span>
                 </div>
               </div>}

        {mode === 'lasso' && lassoPoints.length > 0 && (
          <svg className="absolute inset-0 w-full h-full pointer-events-none rounded-2xl" style={{zIndex: 900}} viewBox="0 0 100 100" preserveAspectRatio="none">
            <path d={'M ' + lassoPoints.map(p => `${p.x} ${p.y}`).join(' L ')} stroke="#0ea5e9" strokeWidth="0.6" strokeDasharray="1,1" fill="rgba(14, 165, 233, 0.15)" vectorEffect="non-scaling-stroke" />
          </svg>
        )}

        {visiblePins.map(pin => (
          <div key={pin.pin_id} className="absolute" style={{ left: `${pin.x}%`, top: `${pin.y}%` }}>
            <PinIcon color={pin.color} onClick={(e) => {
              e.stopPropagation();
              // なげなわ描画中以外はいつでも詳細を開ける（編集モードでも見られるように）
              if(mode !== 'lasso') setSelectedPin(pin);
            }} />
          </div>
        ))}
        {creatingPos && <div className="absolute z-50" style={{ left: `${creatingPos.x}%`, top: `${creatingPos.y}%` }}><PinIcon color="#999" isNew /></div>}
      </div>

      {creatingPos && <PinFormModal pos={creatingPos} onClose={()=>setCreatingPos(null)} />}
      {editingPin && <PinFormModal pin={editingPin} onClose={()=>setEditingPin(null)} />}
      {selectedPin && <PinDetailModal pin={selectedPin} onClose={() => setSelectedPin(null)} onEdit={(p) => { setSelectedPin(null); setEditingPin(p); }} />}
      {areaPins && <PinListModal pins={areaPins} onClose={() => setAreaPins(null)} onSelectPin={(pin) => { setAreaPins(null); setSelectedPin(pin); }} />}
    </div>
  );
};

// ==========================================
// MainApp（アプリ本体。先生も児童も同じ画面で、出せるものだけが違う）
// ==========================================
const MainApp = ({ ctx, onExit }) => {
  const [state, dispatch] = useState({ user: null, users: [], activeUnit: null, activeMapId: null, pins: [], chats: [], reactions: [], filter: { scope: 'all', color: 'all' }, hasApiKey: false });
  const [sysState, setSysState] = useState({ loading: true, error: null, toast: null });
  const [uiState, setUiState] = useState({ chatOpen: false, notesOpen: false, teacherOpen: false });

  const showToast = useCallback((msg, type='success') => setSysState(p=>({...p, toast: {msg, type}})), []);
  const api = ctx.api;
  // 書き込み時刻を記録し、直後の定期同期が楽観的更新を巻き戻さないようにする
  const lastWriteAt = useRef(0);
  const doAction = useCallback(async (action, data) => {
    lastWriteAt.current = Date.now();
    const res = await ctx.doAction(action, data);
    lastWriteAt.current = Date.now();
    // 単元・地図の構成変更は次の定期同期を待たず即時反映する
    // （特に最初の単元作成時は同期が動いていないため、これがないと再読み込みが必要だった）
    if (action === 'save_unit' || action === 'add_map') { try { await initLoad(); } catch (e) {} }
    return res;
  }, [ctx]);
  // 管理操作（管理パネル・他人の記録の削除）は先生のときだけ。
  // 画面の出し分けは案内であって防御ではない（サーバー側でも必ず拒否する）
  const canAdmin = state.user?.role === 'teacher';

  const initLoad = async () => {
    try {
      const res = await api('GetInitData');
      if(res.status === 'unregistered') {
        setSysState(p=>({...p, loading: false, error: 'unregistered'}));
      } else {
        dispatch(p => {
          const mapIds = (res.activeUnit?.maps || []).map(m => m.id);
          return {
            ...p, user: res.user, users: res.users, activeUnit: res.activeUnit,
            // 表示中の地図タブは有効な限り維持する（地図追加後などにタブが勝手に戻らないように）
            activeMapId: mapIds.indexOf(p.activeMapId) >= 0 ? p.activeMapId : (mapIds[0] || null),
            pins: res.pins, chats: res.chats, reactions: res.reactions || [], hasApiKey: !!res.hasApiKey
          };
        });
        setSysState(p=>({...p, loading: false}));
      }
    } catch(e) { setSysState(p=>({...p, loading: false, error: e.message || 'エラーが発生しました'})); }
  };

  useEffect(() => {
    const unitId = state.activeUnit?.unit_id;
    if (sysState.loading && !unitId) initLoad();
    const intervalMs = state.user?.role === 'teacher' ? 10000 : 15000;
    const timer = setInterval(async () => {
      if(!unitId) return;
      if(document.hidden) return; // 非表示タブでは同期しない（サーバー負荷・クォータ対策）
      const startedAt = Date.now();
      try {
        const res = await api('SyncData', unitId);
        if(res.success) {
          // 先生が新しい単元を開始していたら、全データを取り直して自動で切り替える
          if (res.activeUnit && res.activeUnit.unit_id !== unitId) { initLoad(); return; }
          // 同期開始後に書き込みがあった場合、記録系の反映は次回に見送る
          // （保存直前に読まれた古いスナップショットで新しいピンが消えるのを防ぐ）
          const skipRecords = lastWriteAt.current > startedAt - 500;
          dispatch(p => {
             const newUnit = res.activeUnit || p.activeUnit;
             const mapIds = (newUnit?.maps || []).map(m => m.id);
             const mapId = mapIds.indexOf(p.activeMapId) >= 0 ? p.activeMapId : (mapIds[0] || null);
             return { ...p,
               pins: skipRecords ? p.pins : res.pins,
               chats: skipRecords ? p.chats : res.chats,
               reactions: skipRecords ? p.reactions : (res.reactions || []),
               users: res.users || p.users,
               activeUnit: newUnit, activeMapId: mapId,
               hasApiKey: res.hasApiKey !== undefined ? !!res.hasApiKey : p.hasApiKey };
          });
        }
      } catch(e) {}
    }, intervalMs);
    return () => clearInterval(timer);
  }, [state.activeUnit?.unit_id]);

  if (sysState.loading) return <LoadingOverlay />;
  if (sysState.error === 'unregistered') return (
    <div className="app-h flex items-center justify-center bg-surface p-4">
      <div className="bg-white p-12 rounded-[32px] shadow-float text-center max-w-md w-full border-t-[10px] border-brand-500 relative overflow-hidden">
        <div className="absolute top-0 left-0 w-full h-32 bg-brand-50/50"></div>
        <span className="text-7xl block mb-6 relative z-10 drop-shadow-sm">🏫</span>
        <h1 className="text-2xl font-extrabold text-slate-800 mb-4 relative z-10 tracking-tight">まだ <RubyText text="登録" kana="とうろく" />されていません</h1>
        <p className="text-slate-500 font-medium text-sm leading-relaxed relative z-10 bg-slate-50 p-4 rounded-xl border border-slate-100">
          <span><RubyText text="先生" kana="せんせい" />に メールアドレスを<RubyText text="登録" kana="とうろく" />してもらってから、もう<RubyText text="一度" kana="いちど" /> ひらいてね！</span>
        </p>
      </div>
    </div>
  );
  if (sysState.error) return <div className="app-h flex items-center justify-center font-bold text-rose-500 bg-surface p-6 text-center">{sysState.error}</div>;

  return (
    <AppContext.Provider value={{ state, dispatch, showToast, api, doAction, ctx, canAdmin }}>
      <div className="app-h w-full flex flex-col bg-surface overflow-hidden">

        <header className="flex-none h-14 sm:h-16 bg-white/90 backdrop-blur-md z-50 px-3 md:px-6 flex justify-between items-center border-b border-slate-200/60 shadow-sm no-print">
          <div className="flex items-center gap-3">
            {onExit && (
              <button onClick={onExit} className="bg-slate-100 text-slate-600 px-3 py-2 rounded-full font-bold text-xs hover:bg-slate-200 transition no-print">← ポータル</button>
            )}
            <span className="text-2xl bg-brand-50 p-1.5 rounded-xl shadow-inner-soft border border-brand-100">🗺️</span>
            <h1 className="text-xl font-extrabold text-slate-800 tracking-tight hidden sm:block">みっけ！</h1>
          </div>

          {state.user.role !== 'teacher' && (
            <div className="flex bg-slate-100/80 p-1 shadow-inner-soft border border-slate-200/80 rounded-full">
              <button onClick={()=>dispatch(p=>({...p, filter:{...p.filter, scope:'all'}}))} className={`px-4 sm:px-5 py-1.5 rounded-full text-xs sm:text-sm font-bold transition-all ${state.filter.scope==='all'?'bg-white shadow-sm text-brand-600':'text-slate-500 hover:text-slate-700'}`}>みんな</button>
              <button onClick={()=>dispatch(p=>({...p, filter:{...p.filter, scope:'group'}}))} className={`px-4 sm:px-5 py-1.5 rounded-full text-xs sm:text-sm font-bold transition-all ${state.filter.scope==='group'?'bg-white shadow-sm text-brand-600':'text-slate-500 hover:text-slate-700'}`}>
                <span className="mt-0.5 inline-block">じぶんの<RubyText text="班" kana="はん" /></span>
              </button>
              <button onClick={()=>dispatch(p=>({...p, filter:{...p.filter, scope:'mine'}}))} className={`px-4 sm:px-5 py-1.5 rounded-full text-xs sm:text-sm font-bold transition-all ${state.filter.scope==='mine'?'bg-white shadow-sm text-brand-600':'text-slate-500 hover:text-slate-700'}`}>じぶん</button>
            </div>
          )}

          <div className="flex items-center gap-2 sm:gap-4">
            <div className="text-right hidden md:block">
              <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">{state.user.group_id}</div>
              <div className="font-extrabold text-slate-700 text-sm">{state.user.name} <span className="text-[10px] font-normal text-slate-500">さん</span></div>
            </div>
            <button onClick={()=>setUiState(p=>({...p, notesOpen: !p.notesOpen, chatOpen: false}))} className={`relative p-2 sm:p-2.5 rounded-full shadow-sm hover:shadow-md border transition text-lg ${uiState.notesOpen ? 'bg-accent-50 border-accent-200' : 'bg-slate-50 border-slate-200'}`} title="わたしのきろく">
              📒 <span className="absolute -top-1.5 -right-1.5 bg-brand-500 text-white text-[10px] font-bold w-5 h-5 flex items-center justify-center rounded-full shadow-sm border-2 border-white">{state.pins.filter(p=>p.email===state.user.email).length}</span>
            </button>
            <button onClick={()=>setUiState(p=>({...p, chatOpen: !p.chatOpen, notesOpen: false}))} className={`relative p-2 sm:p-2.5 rounded-full shadow-sm hover:shadow-md border transition text-lg ${uiState.chatOpen ? 'bg-brand-50 border-brand-200' : 'bg-slate-50 border-slate-200'}`} title="みんなのひろば">
              💬 <span className="absolute -top-1.5 -right-1.5 bg-accent-500 text-white text-[10px] font-bold w-5 h-5 flex items-center justify-center rounded-full shadow-sm border-2 border-white">{state.chats.length}</span>
            </button>
            {canAdmin && (
              <button onClick={()=>setUiState(p=>({...p, teacherOpen: true}))} className="bg-slate-800 text-white px-3 sm:px-5 py-2 sm:py-2.5 rounded-full font-bold text-xs shadow-md hover:bg-black transition whitespace-nowrap">管理パネル</button>
            )}
          </div>
        </header>

        <div className="flex-1 relative flex overflow-hidden">
          <MapArea />
          <ChatPanel isOpen={uiState.chatOpen} onClose={()=>setUiState(p=>({...p, chatOpen: false}))} />
          <MyNotesPanel isOpen={uiState.notesOpen} onClose={()=>setUiState(p=>({...p, notesOpen: false}))} />
        </div>

        {/* モバイルでは地図領域を最大化するためフッターを表示しない */}
        <footer className="flex-none h-8 hidden md:flex justify-center items-center bg-white border-t border-slate-200 text-[10px] sm:text-xs text-slate-400 font-bold z-50 tracking-wider no-print">
          © 2026 みっけ！ <a href="https://giga-school.com" target="_blank" className="ml-1.5 text-slate-500 hover:text-brand-500 transition">GIGA山</a>
          {/* ⚠️ このフッターは hidden md:flex。地図を広く使うため、スマホでは出ない。
                 リンクもスマホでは出ないので、行き先を増やすときは他の場所も要る。 */}
          <a href="https://giga-school.com/apps/townmap-mikke/" target="_blank" rel="noopener noreferrer" className="ml-3 text-slate-500 hover:text-brand-500 transition">使い方を読む</a>
        </footer>

        {uiState.teacherOpen && <TeacherConsole onClose={()=>setUiState(p=>({...p, teacherOpen: false}))} />}
        <Toast msg={sysState.toast?.msg} type={sysState.toast?.type} onHide={()=>setSysState(p=>({...p, toast: null}))} />
      </div>
    </AppContext.Provider>
  );
};

// ==========================================
// 教員ポータル（デプロイ T / ?portal=1）
// ==========================================
// ==========================================
// 学級ゲート（この URL がその学級そのもの）
//
// サーバー（Bound.gs）が「開いている本人が誰か」「先生か児童か」を返す。
// 画面の出し分けは案内のためであって、防御ではない（防御はサーバー側）。
// ==========================================
const BoundGate = () => {
  const api = callApi;
  const doAction = doActionApi;
  const [phase, setPhase] = useState({ type: 'connecting' });
  const [joinForm, setJoinForm] = useState({ name: '', number: '' });
  const [joining, setJoining] = useState(false);

  const checkStatus = async () => {
    try {
      const res = await api('GetStatus');
      if (res.state === 'teacher' || res.state === 'active') setPhase({ type: 'app' });
      else if (res.state === 'pending') setPhase({ type: 'pending', className: res.className });
      else if (res.state === 'closed') setPhase({ type: 'closed', className: res.className });
      else if (res.state === 'setup') setPhase({ type: 'setup', message: res.message });
      else setPhase({ type: 'join', className: res.className, requireApproval: res.requireApproval });
    } catch (e) {
      setPhase({ type: 'error', message: e.message });
    }
  };

  useEffect(() => { checkStatus(); }, []);

  // 承認待ち・準備待ちの間は 20 秒ごとに状態を再確認（承認されたら自動で入室）
  useEffect(() => {
    if (phase.type !== 'pending' && phase.type !== 'setup') return;
    const t = setInterval(checkStatus, 20000);
    return () => clearInterval(t);
  }, [phase.type]);

  const handleJoin = async () => {
    if (!joinForm.name.trim()) return;
    setJoining(true);
    try {
      const res = await api('Join', joinForm.name.trim(), joinForm.number.trim());
      if (res.state === 'active') setPhase({ type: 'app' });
      else setPhase({ type: 'pending', className: phase.className });
    } catch (e) {
      setPhase({ type: 'error', message: e.message });
    }
    setJoining(false);
  };

  if (phase.type === 'connecting') return <LoadingOverlay label={<span>つないでいます...</span>} />;

  if (phase.type === 'app') {
    const ctx = { api, doAction };
    return <MainApp ctx={ctx} />;
  }

  const Card = ({ children, color = 'brand-500' }) => (
    <div className="app-h flex items-center justify-center bg-surface p-4">
      <div className={`bg-white p-10 sm:p-12 rounded-[32px] shadow-float text-center max-w-md w-full border-t-[10px] border-${color} relative overflow-hidden animate-pop-in`}>
        {children}
      </div>
    </div>
  );

  if (phase.type === 'join') return (
    <Card>
      <span className="text-6xl block mb-4">👋</span>
      <h1 className="text-2xl font-extrabold text-slate-800 mb-1">{phase.className || 'みっけ！'}</h1>
      <p className="text-sm text-slate-500 font-bold mb-6"><RubyText text="参加" kana="さんか" />の<RubyText text="申請" kana="しんせい" />をしよう</p>
      <div className="space-y-3 text-left">
        <div>
          <label className="block text-[11px] font-bold text-slate-500 mb-1"><RubyText text="表示" kana="ひょうじ" />する<RubyText text="名前" kana="なまえ" />（ニックネームでもOK）</label>
          <input type="text" value={joinForm.name} onChange={e=>setJoinForm(p=>({...p, name: e.target.value}))} placeholder="たろう" className="w-full px-5 py-4 border border-slate-200 rounded-xl bg-slate-50 font-bold focus:ring-2 focus:ring-brand-500 focus:bg-white outline-none transition" />
        </div>
        <div>
          <label className="block text-[11px] font-bold text-slate-500 mb-1"><RubyText text="出席番号" kana="しゅっせきばんごう" />（なくてもOK）</label>
          <input type="text" value={joinForm.number} onChange={e=>setJoinForm(p=>({...p, number: e.target.value}))} placeholder="12" className="w-full px-5 py-4 border border-slate-200 rounded-xl bg-slate-50 font-bold focus:ring-2 focus:ring-brand-500 focus:bg-white outline-none transition" />
        </div>
      </div>
      <button onClick={handleJoin} disabled={joining || !joinForm.name.trim()} className="w-full mt-6 bg-brand-500 text-white py-4 rounded-xl font-bold text-lg hover:bg-brand-600 transition shadow-float disabled:opacity-50">
        {joining ? <span><RubyText text="申請中" kana="しんせいちゅう" />...</span> : <span><RubyText text="参加" kana="さんか" />する！</span>}
      </button>
      {phase.requireApproval && <p className="text-[11px] text-slate-400 mt-4"><RubyText text="先生" kana="せんせい" />が<RubyText text="承認" kana="しょうにん" />すると<RubyText text="使" kana="つか" />えるようになるよ</p>}
    </Card>
  );

  if (phase.type === 'pending') return (
    <Card color="amber-400">
      <span className="text-6xl block mb-4 animate-bounce-pin">⏳</span>
      <h1 className="text-2xl font-extrabold text-slate-800 mb-3"><RubyText text="承認" kana="しょうにん" />を<RubyText text="待" kana="ま" />っています</h1>
      <p className="text-sm text-slate-500 leading-relaxed bg-slate-50 p-4 rounded-xl border border-slate-100">
        <RubyText text="先生" kana="せんせい" />が「OK」すると、<RubyText text="自動" kana="じどう" />でこの<RubyText text="画面" kana="がめん" />が<RubyText text="切" kana="き" />りかわるよ。<br/>このまま<RubyText text="待" kana="ま" />っていてね！
      </p>
    </Card>
  );

  if (phase.type === 'closed') return (
    <Card color="slate-400">
      <span className="text-6xl block mb-4">🚪</span>
      <h1 className="text-2xl font-extrabold text-slate-800 mb-3"><RubyText text="受付" kana="うけつけ" />が<RubyText text="閉" kana="と" />じています</h1>
      <p className="text-sm text-slate-500 leading-relaxed bg-slate-50 p-4 rounded-xl border border-slate-100">いまは<RubyText text="新" kana="あたら" />しく<RubyText text="参加" kana="さんか" />できません。<RubyText text="先生" kana="せんせい" />に<RubyText text="確認" kana="かくにん" />してね。</p>
    </Card>
  );

  // 先生がまだ「はじめの設定」をしていない。ここで開いた人を先生にはしない。
  if (phase.type === 'setup') return (
    <Card color="amber-400">
      <span className="text-6xl block mb-4">🛠️</span>
      <h1 className="text-xl font-extrabold text-slate-800 mb-3">じゅんびちゅうです</h1>
      <p className="text-sm text-slate-600 leading-relaxed bg-slate-50 p-4 rounded-xl border border-slate-100 text-left">
        {phase.message || '先生の準備がまだ終わっていません。'}
      </p>
      <p className="text-[11px] text-slate-400 mt-4">じゅんびができると、この画面はひとりでに切りかわります。</p>
    </Card>
  );

  return (
    <Card color="rose-400">
      <span className="text-6xl block mb-4">⚠️</span>
      <p className="font-bold text-slate-700 leading-relaxed text-sm whitespace-pre-wrap">{phase.message || 'エラーが発生しました'}</p>
      <button onClick={()=>{ setPhase({type:'connecting'}); checkStatus(); }} className="mt-6 bg-slate-800 text-white px-8 py-3 rounded-xl font-bold hover:bg-black transition">もう一度ためす</button>
    </Card>
  );
};

// 入口は 1 つだけ。誰が先生で誰が児童かはサーバーが返す。
const Root = () => <BoundGate />;

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<Root />);
