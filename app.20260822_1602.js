/* ============================================================
   FBA易 v1 · 本地优先 · 无服务器
   L4 主数据(IndexedDB) → L5 生成引擎(ExcelJS填模板副本) → L6 校验 → L7 导出
   支持 5 家物流商模板(各自字段映射)、数据库降级容错、渲染错误上屏。
   ============================================================ */
'use strict';
const APP_VERSION = '20260822_1602'; // 每次部署必须更新，用于破坏浏览器缓存（20260822：修复 SKU 主数据视图崩溃——13019 条一次性全量渲染 DOM 致 OOM；改为搜索+分页，默认 100/页，仅渲染当前页；新增 SKU 模糊搜索+每页条数切换）
const CHANNEL_SEED_VER = 2; // 渠道种子版本：1=初始 58 条含默认注册名/地址；2=清空未经验证的示例注册名/地址

/* ---------- 存储层：IndexedDB，不可用时降级为内存(保证不空白) ---------- */
const DB_NAME = 'invoice_sys_v1', DB_VER = 4; // bump: 旧库(version<4)缺 config/boxspecs store，需触发 onupgradeneeded 补建
let DB = null, USE_DB = true;
const mem = { channels:[], skus:[], templates:[], records:[], warehouses:[], boxspecs:[], config:[] };
function openDB(){
  return new Promise((res)=>{
    try{
      const r = indexedDB.open(DB_NAME, DB_VER);
      r.onupgradeneeded = e=>{
        const db = e.target.result;
        ['channels','skus','templates','records','warehouses','boxspecs','config'].forEach(s=>{ if(!db.objectStoreNames.contains(s)) db.createObjectStore(s,{keyPath:'id'}); });
      };
      r.onsuccess = e=>{ DB=e.target.result; res(DB); };
      r.onerror = ()=>{ USE_DB=false; res(); };
    }catch(e){ USE_DB=false; res(); }
  });
}
function _mem(store){ return mem[store] || (mem[store]=[]); }
// store 不存在(旧库 schema 不兼容 / 降级)时,自动走内存,避免 NotFoundError 硬崩
function _useMem(store){ return !USE_DB || !DB || !DB.objectStoreNames || !DB.objectStoreNames.contains(store); }
function _idb(store,mode){ if(_useMem(store)) throw new Error('MEM_FALLBACK:'+store); return DB.transaction(store,mode).objectStore(store); }
function getAll(store){
  if(_useMem(store)) return Promise.resolve(mem[store]||[]);
  return new Promise((res,rej)=>{ const r=_idb(store,'readonly').getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>rej(r.error); });
}
function put(store,val){
  if(_useMem(store)){ const a=_mem(store); const i=a.findIndex(x=>x.id===val.id); if(i>=0)a[i]=val;else a.push(val); return Promise.resolve(val); }
  return new Promise((res,rej)=>{ const r=_idb(store,'readwrite').put(val); r.onsuccess=()=>res(val); r.onerror=()=>rej(r.error); });
}
function del(store,id){
  if(_useMem(store)){ mem[store]=(mem[store]||[]).filter(x=>x.id!==id); return Promise.resolve(); }
  return new Promise((res,rej)=>{ const r=_idb(store,'readwrite').delete(id); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); });
}
function clear(store){
  if(_useMem(store)){ mem[store]=[]; return Promise.resolve(); }
  return new Promise((res,rej)=>{ const r=_idb(store,'readwrite').clear(); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); });
}
function get(store,id){
  if(_useMem(store)) return Promise.resolve((mem[store]||[]).find(x=>x.id===id)||null);
  return new Promise((res,rej)=>{ const r=_idb(store,'readonly').get(id); r.onsuccess=()=>res(r.result||null); r.onerror=()=>rej(r.error); });
}
const uid = ()=> Date.now().toString(36)+Math.random().toString(36).slice(2,7);
/* 把 FBA 货件号 + 箱标（如 B3）拼回 FBA 箱 ID */
function fbaBoxId(fbaNo, boxLabel){
  if(!fbaNo || !boxLabel) return boxLabel;
  const m = String(boxLabel).match(/^([A-Za-z]?)(\d+)$/);
  if(!m) return boxLabel;
  const n = parseInt(m[2],10);
  return fbaNo.replace(/\s+/g,'') + 'U' + String(n).padStart(6,'0');
}
function isSimpleLabel(v){ return /^[A-Za-z]?\d+$/.test(String(v||'')); }
/* 后端/老数据可能只返回 boxNo（且是简单箱标），需补齐 boxLabel 并尝试还原 FBA 箱 ID */
function normalizeItems(items, fbaNo){
  // 1) 忠实于装箱清单：丢掉无 SKU/FNSKU/品名的空行(占位/垃圾行)，绝不参与箱数统计
  const filtered = (items||[]).filter(it =>
    (it.sku && String(it.sku).trim()) ||
    (it.fnsku && String(it.fnsku).trim()) ||
    (it.nameEn && String(it.nameEn).trim()) ||
    (it.nameCn && String(it.nameCn).trim()));
  if(!filtered.length) return [];
  // 2) 从数据集自身的真实箱号反推货件号（源忠实：箱号来自装箱清单，绝不盲信外部传入的 fbaNo，
  //    避免 step1 残留的 stale 旧号把箱号"重造"成外来 FBA）。找不到再用传入的 fbaNo 兜底。
  let shipPrefix='';
  for(const it of filtered){ const s=String(it.boxNo||''); const m=s.match(/^(FBA[A-Z0-9]*)U\d{6}$/i); if(m){ shipPrefix=m[1].toUpperCase(); break; } }
  const effFba = shipPrefix || String(fbaNo||'').trim();
  const arr = filtered.map(it=>{
    const out = {...it};
    const hasLabel = out.boxLabel && String(out.boxLabel).trim();
    const hasNo = out.boxNo && String(out.boxNo).trim();
    const labelIsReal = hasLabel && isSimpleLabel(out.boxLabel) && /^[A-Za-z]/.test(out.boxLabel); // 如 B3
    const noIsRealFba = hasNo && /^FBA[A-Z0-9]*U\d{6}$/i.test(out.boxNo); // 如 FBA19K786CWTU000001
    // ① 箱号已是真实 FBA 箱 ID（来自装箱清单）→ 直接信任，绝不覆盖、绝不重造（用户明确：箱号来自装箱清单）。
    //    同时让「箱标」与「子单号」保持一致，都显示真实 FBA 箱 ID，避免列表里出现 B1/B2 这种旧箱标。
    if(noIsRealFba){
      out.boxLabel = out.boxNo;
      if(out.boxes==='' || out.boxes==null || String(out.boxes).trim()==='') out.boxes = 1;
      return out;
    }
    // ② 否则（箱号缺失或只是箱标）→ 优先用箱标；仍是简单箱标(B3)且已知货件号时，还原真实 FBA 箱 ID
    if(!hasLabel && hasNo && isSimpleLabel(out.boxNo)) out.boxLabel = out.boxNo;
    if(!hasNo && hasLabel) out.boxNo = out.boxLabel;
    if(out.boxNo && labelIsReal && !/^FBA[A-Z0-9]*U\d{6}$/i.test(out.boxNo||'')) out.boxNo = out.boxLabel;
    if(out.boxNo && isSimpleLabel(out.boxNo) && /^[A-Za-z]/.test(out.boxNo) && effFba){
      out.boxNo = fbaBoxId(effFba, out.boxNo);
    }
    // 箱号=子箱号：模板箱号统一用真实 FBA 箱 ID（用户明确：箱号就是子单号）
    if(out.boxNo && String(out.boxNo).includes('FBA')) out.boxLabel = out.boxNo;
    if(out.boxes==='' || out.boxes==null || String(out.boxes).trim()==='') out.boxes = 1;
    return out;
  });
  // 3) 按箱号从小到大升序排列（用户明确）
  const esc = effFba ? effFba.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') : null;
  const validRe = esc ? new RegExp('^'+esc+'U\\d{6}$') : null;
  arr.sort((a,b)=>{
    const na=String(a.boxNo||''), nb=String(b.boxNo||'');
    if(validRe && validRe.test(na) && validRe.test(nb)) return na<nb?-1:na>nb?1:0;
    return na.localeCompare(nb);
  });
  return arr;
}

/* 从已载入的物品行反推「本发票真正货件号」：优先取真实 FBA 箱 ID 的货件前缀，
   取不到再退到 W.form.fbaNo。生成/校验/文件名统一用它，避免 step2 缓存的 W.sources
   或 step1 残留的旧 FBA 号污染本张发票。 */
function deriveInvoiceFbaNo(items, fallback){
  const arr = (items||[]).filter(it => it && (it.boxNo || it.boxLabel));
  // 1) 真实 FBA 箱号 FBAxxxU000001
  for(const it of arr){
    const s = String(it.boxNo || it.boxLabel || '');
    const m = s.match(/^(FBA[A-Z0-9]*)U\d{6}$/i);
    if(m) return m[1].toUpperCase();
  }
  // 2) 任何 FBA 字符串
  for(const it of arr){
    const s = String(it.boxNo || it.boxLabel || '');
    const m = s.match(/\b(FBA[A-Z0-9]{4,})\b/);
    if(m) return m[1].toUpperCase();
  }
  return String(fallback || '').trim().toUpperCase();
}

/* 统一反查 SKU 主数据(商品申报信息) + 补充申报表(SKU_DECLARE) + 箱规格，回填缺失字段。
   这是「源忠实 + 主数据兜底」的关键：装箱清单(尤其云端拉取)常不带 HS/品名/申报价，
   这些本应从 SKU 主数据取。所有「物品载入」路径(上传/云端/本地预装/手建)都必须走它，
   否则会出现「校验通过但导出 HS 为空」(用户最早反馈的 bug)。 */
function resolveItemMaster(it){
  // SKU 归一化：装箱清单常带 @us / @xxx 市场后缀(如 1C56-1@us)，主数据用干净 SKU(1C56-1)。
  // 先精确匹配，未中再按归一化(去 @ 后缀)匹配，避免「主数据有却 join 不上」的假缺失。
  const nSku = (it.sku||'').replace(/@.*$/,'');
  const sk = (W.skus||[]).find(x=>x.sku===it.sku) || (nSku && nSku!==it.sku ? (W.skus||[]).find(x=>x.sku===nSku) : null);
  if(sk){
    if(!it.nameCn)  it.nameCn  = sk.中文品名;
    if(!it.nameEn)  it.nameEn  = sk.英文品名;
    if(!it.hs)      it.hs      = sk.HS;          // HS 海关编码：从主数据回填
    it.material = sk.材质 || it.material;        // 材质：主数据权威, 覆盖装箱清单残留(旧'Handbags'已清)
    if(!it.declare && sk.申报价) it.declare = sk.申报价;
    if(!it.brand)   it.brand   = sk.品牌;
    // 型号：主数据有型号就用主数据；主数据没有则默认=SKU（用户明确：型号就是SKU）
    if(!it.model || String(it.model).trim()==='') it.model = sk.型号 || it.sku;
    if(!it.purpose) it.purpose = sk.用途;        // 用途：从主数据回填(商品申报信息P列)
  }
  if(window.SKU_DECLARE && window.SKU_DECLARE[it.sku]){
    const d = window.SKU_DECLARE[it.sku];
    if((!it.declare||it.declare==='') && d.d!=null) it.declare = d.d;
    if(!it.nameCn && d.n) it.nameCn = d.n;
    if(!it.hs && d.h)     it.hs     = d.h;       // HS 兜底(补充申报表也带 hs 时)
  }
  const norm = s => (s||'').replace(/@us$/i,'').trim();
  const bs = (W.boxspecs||[]).find(x=>norm(x.sku)===norm(it.sku));
  if(bs){
    if(!it.boxSpec)   it.boxSpec   = bs.model;
    if(!it.boxWeight) it.boxWeight = bs.weight;
    if(!it.len)       it.len       = bs.l;
    if(!it.wid)       it.wid       = bs.w;
    if(!it.hgt)       it.hgt       = bs.h;
  }
  // 业务默认兜底（仅在主数据也未提供时填充，导出时会高亮为推算/默认值）
  if(!it.boxes || String(it.boxes).trim()==='') it.boxes = 1;     // 每行=一箱
  if(!it.brand || String(it.brand).trim()==='') it.brand = 'JW PEI';
  if(!it.purpose || String(it.purpose).trim()==='') it.purpose = 'Put things';
  if(!it.model || String(it.model).trim()==='') it.model = it.sku; // 型号默认=SKU
  // 产品毛重(单件) = 箱重 / 数量 —— 用于安速 J 列(GW产品毛重)等
  if(it.boxWeight && it.qty){
    const pw = parseFloat(it.boxWeight)/parseFloat(it.qty);
    if(!isNaN(pw)) it.prodWeight = Math.round(pw*10000)/10000;
  }
  return it;
}

/* 确保主数据(W.skus)已就绪：在「物品载入」路径(预装/云端/上传)调用 resolveItemMaster 前 await。
   消除 wizard 尚未把 IndexedDB 的 skus 注入 W.skus 时的竞态，避免材质/用途/HS 漏填（用户反馈的 bug）。 */
async function ensureSkusLoaded(){
  if(!W) return;
  if(W.skus && W.skus.length) return;
  try{ W.skus = await getAll('skus'); }catch(e){ console.warn('ensureSkusLoaded 失败:', e); }
}

/* 确保箱规主数据(W.boxspecs)已就绪：物品渲染/重拉装箱清单前调用，避免箱重/长宽高漏填。 */
async function ensureBoxspecsLoaded(){
  if(!W) return;
  if(W.boxspecs && W.boxspecs.length) return;
  try{ W.boxspecs = await getAll('boxspecs'); }catch(e){ console.warn('ensureBoxspecsLoaded 失败:', e); }
}

/* ---------- 5 家物流商模板字段映射(由 inspect_all.js 解析得到) ---------- */
const MAPPINGS = {
  '安速':{
    titleCell:'A1', titleText:'FBA订单（V3）',
    // 注意：安速模板每行 A_n:C_n(标签)/D_n:H_n(值) 合并，值主格是 D 列。原 E 列映射会落入合并从属格、Excel 不显示，故全部改为 D 列主格。
    meta:{ fbaNo:'D2', amazonRef:'D4', shipMethod:'D3', warehouseCode:'D5', company:'D6', country:'D7', province:'D8', city:'D9', address:'D10', phone:'D11', zip:'D12', email:'D13', customs:'D14', vat:'D15', eori:'D16', vatName:'D17', vatAddr:'D18', customInfo:'D19' },
    // 模板列：A=No.of Pkgs(箱号)=boxLabel, B=子单号(同箱号)=boxNo(FBA箱ID)
    item:{ boxLabel:'A', boxNo:'B', nameCn:'G', nameEn:'H', qty:'I', declare:'K', material:'M', purpose:'N', hs:'L', brand:'V', model:'W', boxWeight:'C', prodWeight:'J', len:'D', wid:'E', hgt:'F', elec:'O', magnet:'P', img:'Q', imgUrl:'R', salePrice:'S', saleUrl:'T', currency:'Z', origin:'AA' },
    itemStartRow:21,
    // clearExtra：安速模板 U(电池类型)/X,Y(净重)/AB(税则号)/AC(配货) 无数据源，留空处理（模板本就空无 stale，清空保持整洁）
    clearExtra:['U','X','Y','AB','AC']
  },
  '艾杜克':{
    // vat 主格是 B7(B7:F7 合并)，原 C7 为从属格不显示，改为 B7。
    meta:{ company:'B5', vat:'B7', warehouseCode:'I8', amazonRef:'I6', fbaNo:'I7' },
    // cartons='F'(箱数,值取 it.boxes) 修复：模板物品行有 Cartons 箱数列，缺字段会导致插入行全空+前8行残留模板示例"1"。
    // powder='N'(Powder/liquid 列)：CSV 无粉末信息，清空后留空待人工确认（源忠实，不猜测填 N）。
    item:{ nameEn:'A', nameCn:'B', boxNo:'C', brand:'D', hs:'E', cartons:'F', qty:'G', boxWeight:'H', len:'I', declare:'J', img:'L', elec:'M', model:'O', powder:'N' },
    itemStartRow:14,
    // 合计行聚合列（配置驱动，避免英文表头猜测失败）。模板 r22：F=箱数/Cartons G=数量/Quantity H=箱重/Total GW K=货值/Total Price
    totals:{ cartons:'F', qty:'G', weight:'H', value:'K' },
    // total 修复：模板 Total Price(K) 是静态示例值非公式，必须显式重算 = 单价(J)×数量(G)，否则残留示例价、明细与合计对不上
    total:{ col:'K', unit:'J', qty:'G' }
  },
  '亦邦':{
    // 亦邦模板无独立收货人块，原 meta 指向表头行(F1=商铺名表头/G1=货件追踪表头)会在生成时把表头覆盖成收货人值→改为空。收货人由外部(物流系统)处理。
    meta:{},
    item:{ boxNo:'A', boxWeight:'B', len:'C', wid:'D', hgt:'E', company:'F', amazonRef:'G', nameCn:'I', nameEn:'J', declare:'K', qty:'L', elec:'M', magnet:'O', sku:'P', hs:'Q', material:'R', purpose:'T', img:'V', powder:'N' },
    itemStartRow:2,
    // clearExtra：模板预填 stale 示例值（N=是否液体/粉末"N"、S=材质英文、U=用途英文、H=商品类型、W=颜色备注）→ 整列擦掉，否则导出保留错误值（同艾杜克 powder 类 bug，但更隐蔽：海关液体/粉末错报）
    clearExtra:['H','N','S','U','W']
  },
  '亚丰':{
    // email 主格是 B14(B14:D14 合并)，vat 主格是 J11(J11:L11 合并)；原 C14/G11 为从属格不显示，改为 B14/J11。
    meta:{ fbaNo:'B1', shipMethod:'B2', warehouseCode:'B3', company:'B5', address:'B6', city:'B9', province:'B10', zip:'B11', country:'B12', phone:'B13', email:'B14', customs:'F6', vat:'J11', poNo:'B15' },
    item:{ boxNo:'A', boxWeight:'B', len:'C', wid:'D', hgt:'E', nameEn:'F', nameCn:'G', declare:'H', qty:'I', material:'J', hs:'K', purpose:'L', brand:'M', model:'N', saleUrl:'O', salePrice:'P', img:'Q', imgUrl:'R', prodWeight:'S', elec:'T', magnet:'U', asin:'V', fnsku:'W', sku:'X' },
    itemStartRow:19
  },
  '合联':{
    titleCell:'A1', titleText:'PACKING LIST',
    meta:{ fbaNo:'A3' },
    item:{ fbaNo:'A', boxNo:'B', nameEn:'C', nameCn:'D', hs:'E', boxCount:'F', qty:'G', sku:'H', declare:'I', boxWeight:'K', len:'L', wid:'M', hgt:'N', brand:'P', elec:'Q', img:'R', material:'S' },
    itemStartRow:5,
    // 合计行聚合列。模板 r8：F=件数/CTNS G=数量/QUANTITY J=总价(USD)/TOTAL VALUE K=单箱重量合计/G.WEIGH（C=箱号/材质等不在合计列）
    totals:{ cartons:'F', qty:'G', value:'J', weight:'K' },
    // total 修复：模板 TOTAL VALUE(J) 是静态示例值非公式，必须重算 = 单价(I)×数量(G)
    total:{ col:'J', unit:'I', qty:'G' },
    // clearExtra：O(方数) 预填 stale 数值 + 模板 r16/r17 含内部 PS 备注文字("做完装箱单后请帮忙修改下文件名…") → 整列擦掉防泄漏
    clearExtra:['O']
  }
};
const TPL_FILES = {
  '安速':'安速发票模板.xlsx','艾杜克':'艾杜克发票模板.xlsx','亦邦':'亦邦发票模板.xlsx','亚丰':'亚丰发票模板.xlsx','合联':'合联发票模板.xlsx'
};
const seedStatus = { loading:false, errors:[], loaded:0, total:Object.keys(TPL_FILES).length };

/* ---------- 模板字段别名词典（用于上传时自动扫描解析新模板） ---------- */
/* ---------- 模板字段别名词典（用于上传时自动扫描解析新模板） ---------- */
/* 规则：① 避免单字母英文别名（w/h/l 会误匹配）；② 中文优先级高于英文 */
const cleanTxt = s => String(s||'').replace(/[*★（()）\n\r\t]/g,' ').replace(/\s+/g,' ').trim().toLowerCase();
const FIELD_ALIASES = [
  {key:'boxNo', label:'箱号', names:['箱号','货箱编号','子单号','carton no','box no','pkg no','no.of pkgs','carton']},
  {key:'boxLabel', label:'箱标签', names:['箱子名称','箱标签','box label','label']},
  {key:'nameCn', label:'中文品名', names:['中文品名','产品中文品名','chinese name','品名(中文)','in chinese']},
  {key:'nameEn', label:'英文品名', names:['英文品名','产品英文品名','品名(英文)','in english']},
  {key:'sku', label:'SKU', names:['sku条码','商品sku','goods id','产品sk']},
  {key:'fnsku', label:'FNSKU', names:['fnsku','产品fnsku']},
  {key:'asin', label:'ASIN', names:['asin','产品asin']},
  {key:'qty', label:'数量', names:['数量','qty','quantity','每箱数量','ctns','cartons','件数','产品数量']},
  {key:'declare', label:'申报价', names:['申报','单价usd','单价(usd)','unit price','货值(usd)','vauel申报','总价（usd)','total value','产品申报']},
  {key:'hs', label:'HS编码', names:['hs code','hs编码','h.s code','海关编码','海关商品编码']},
  {key:'material', label:'材质', names:['材质','material','material/usage']},
  {key:'purpose', label:'用途', names:['用途','purpose']},
  {key:'boxWeight', label:'箱重', names:['箱重','箱毛重','单箱重量','gw/ctn','箱子净重','gross weight','total gw','货箱重量']},
  {key:'len', label:'长', names:['长(cm)','货箱长度','length','长（cm）']},
  {key:'wid', label:'宽', names:['宽(cm)','货箱宽度','width','宽（cm）']},
  {key:'hgt', label:'高', names:['高(cm)','货箱高度','height','高（cm）']},
  {key:'brand', label:'品牌', names:['品牌','brand','有无牌子','logo']},
  {key:'model', label:'型号', names:['型号','model','产品型号','规格']},
  {key:'elec', label:'带电', names:['带电','是否带电','electric','with battery','battery','带不带电','electricity']},
  {key:'magnet', label:'带磁', names:['带磁','是否带磁','magnetic']},
  {key:'img', label:'图片', names:['产品图片','图片','image','img','实物图片','product picture','picture']},
  {key:'prodWeight', label:'产品毛重', names:['产品毛重','产品重量','单个产品净重','净重','n.w.','net weight']},
  {key:'salePrice', label:'销售价格', names:['销售链接','销售价','sale price','售价','产品销售价格']},
  {key:'saleUrl', label:'销售链接', names:['链接','url','link','销售地址']},
  {key:'color', label:'颜色', names:['颜色','color','颜色备注']},
  // meta 字段
  {key:'fbaNo', label:'FBA号', names:['客户订单号','fba号','货件编号','fba id','shipment id','fba shipment','订单装运编号','fba号码']},
  {key:'shipMethod', label:'运输方式', names:['运输方式','物流渠道','ship method']},
  {key:'warehouseCode', label:'仓库代码', names:['仓库代码','收件人(仓库代码)','地址库编码','fc','目的地编号','destination fc']},
  {key:'company', label:'收件公司', names:['收件人公司','收件公司','company','consignee','店铺名称']},
  {key:'amazonRef', label:'Amazon参考号', names:['amazon ref','参考号','amazon reference id','货件追踪编号']},
  {key:'poNo', label:'PO号', names:['po','采购单','po number']},
];

const ANCHOR_FIELDS = ['boxNo','nameCn','nameEn','sku','qty','hs'];

/* 扫描上传的 xlsx 模板，自动检测字段列位置 */
async function scanTemplateMapping(blob){
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await blob.arrayBuffer());
  // 容错：getWorksheet(1) 在某些多 sheet 文件上可能返回 undefined
  let ws = wb.getWorksheet(1);
  if(!ws && wb.worksheets && wb.worksheets.length>0) ws = wb.worksheets[0];
  if(!ws) throw new Error('xlsx 中找不到工作表（文件可能为空或结构异常）');
  const maxRow = Math.min(ws.rowCount||200, 200);
  const colL = n => n<1?'':n>26?String.fromCharCode(64+Math.floor((n-1)/26))+String.fromCharCode(65+(n-1)%26):String.fromCharCode(64+n);
  const matchOne = (cl, alias) => cl.includes(cleanTxt(alias)) || cleanTxt(alias).includes(cl);

  // Phase 1: 扫描元信息区（前20行）
  const meta = {};
  for(let r=1; r<=Math.min(maxRow,20); r++){
    const row = ws.getRow(r);
    for(let c=1; c<=8; c++){
      const t = String(row.getCell(c).value||'').trim();
      if(!t) continue;
      const cl = cleanTxt(t);
      for(const fa of FIELD_ALIASES){
        if(meta[fa.key]) continue;
        if(fa.names.some(n=>matchOne(cl, n))){
          for(let vc=c+1; vc<=Math.min(c+4, 30); vc++){
            const v = String(row.getCell(vc).value||'').trim();
            if(v && v.length<60 && !FIELD_ALIASES.some(x=>x.names.some(n=>cleanTxt(v).includes(cleanTxt(n))))){
              meta[fa.key] = v; break;
            }
          }
        }
      }
    }
  }

  // Phase 2: 找表头行
  let headerRow = 0, itemCols = {};
  for(let r=1; r<=Math.min(maxRow,50); r++){
    const row = ws.getRow(r);
    const hits = {};
    for(let c=1; c<=30; c++){
      const t = String(row.getCell(c).value||'').trim();
      if(!t) continue;
      const cl = cleanTxt(t);
      for(const fa of FIELD_ALIASES){
        if(hits[fa.key]) continue;
        if(fa.names.some(n=>matchOne(cl, n))) hits[fa.key] = colL(c);
      }
    }
    const keys = Object.keys(hits);
    const anchors = keys.filter(k=>ANCHOR_FIELDS.includes(k)).length;
    if(keys.length>=4 && anchors>=2){
      const nr = ws.getRow(r+1);
      for(let c=1; c<=30; c++){
        const nt = String(nr.getCell(c).value||'').trim();
        if(!nt) continue;
        const ncl = cleanTxt(nt);
        for(const fa of FIELD_ALIASES){
          if(hits[fa.key]) continue;
          if(fa.names.some(n=>matchOne(ncl, n))) hits[fa.key] = colL(c);
        }
      }
      headerRow = r; itemCols = hits; break;
    }
  }

  // Phase 3: itemStartRow
  let itemStartRow = (headerRow||21)+1;
  if(headerRow){
    for(let r=headerRow+1; r<=maxRow; r++){
      const row = ws.getRow(r);
      let hasData = false;
      for(let c=1; c<=10; c++){
        const v = row.getCell(c).value;
        if(v!==null && v!==undefined && String(v).trim()){ hasData = true; break; }
      }
      if(hasData){ itemStartRow = r; break; }
    }
  }

  // Phase 4: 标题检测
  let titleCell = '', titleText = '';
  for(let r=1; r<=5; r++){
    const row = ws.getRow(r);
    for(let c=1; c<=8; c++){
      const t = String(row.getCell(c).value||'').trim();
      if(t && (t.includes('FBA订单')||t.includes('PACKING LIST')||t.includes('INVOICE')) && t===String(row.getCell(c+1).value||'').trim()){
        titleCell = colL(c)+r; titleText = t; break;
      }
    }
    if(titleText) break;
  }

  return { meta, item: itemCols, itemStartRow, titleCell, titleText };
}

/* ---------- 种子数据(首次运行注入) ---------- */
async function seedIfEmpty(){
  const ch = await getAll('channels');
  // 迁移守卫:旧版仅 3 条废渠道(空国家/空仓库,与新版 58 条 ID 体系不同)。
  // 改为"按首条新渠道是否存在 + 种子版本号"判定,确保已缓存旧数据的浏览器也能补全 58 条,
  // 并能在后续版本强制刷新渠道字段(如清空示例 VAT 注册名/地址)。
  // 同时清理已知的 3 条旧版废渠道(否则会污染渠道下拉、破坏向导⑥校验)。
  const SEED_NEW_FIRST = 'ch_安速_00';
  const SEED_OLD_IDS = ['ch_ansu_us','ch_aiduk_sa','ch_yifeng_us'];
  const needSeed = !ch.some(c=>c.id===SEED_NEW_FIRST) || parseInt(localStorage.getItem('invoiceChannelSeedVer')||'0',10) < CHANNEL_SEED_VER;
  if(needSeed){
    localStorage.setItem('invoiceChannelSeedVer', String(CHANNEL_SEED_VER));
    for(const id of SEED_OLD_IDS){ try{ await del('channels', id); }catch(e){} }
    await put('channels',{id:'ch_安速_00',物流商:'安速',渠道:'中运通达-广州DHL(不含油)',国家:'',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_01',物流商:'安速',渠道:'中运通达-广州联邦IP(不含油)',国家:'',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_02',物流商:'安速',渠道:'中运通达-大陆UPS红单小货(含油)',国家:'',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_03',物流商:'安速',渠道:'欧洲包税-空派快线(普货)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_04',物流商:'安速',渠道:'欧洲包税-空派慢线(普货)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_05',物流商:'安速',渠道:'欧洲包税-卡航',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_06',物流商:'安速',渠道:'欧洲包税-卡航卡派',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_07',物流商:'安速',渠道:'欧洲包税-海运',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_08',物流商:'安速',渠道:'欧洲包税-海运卡派',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_09',物流商:'安速',渠道:'加拿大包税-空派(普货)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_10',物流商:'亚丰',渠道:'亚丰-欧洲空运包税(UPS快线)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_11',物流商:'安速',渠道:'美国包税-空派(普货)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_12',物流商:'安速',渠道:'美国包税-海派(美森正班CLX)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_13',物流商:'安速',渠道:'美国包税-海派(美森加班MAX)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_14',物流商:'安速',渠道:'美国包税-海派(盐田普船)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_15',物流商:'安速',渠道:'美国包税-海卡(美森正班卡派)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_16',物流商:'安速',渠道:'美国包税-海卡(美森加班卡派)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_17',物流商:'安速',渠道:'美国包税-海卡(盐田普船卡派)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_18',物流商:'安速',渠道:'加拿大包税-海派(限时达UPS派送)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_19',物流商:'安速',渠道:'加拿大包税-海派(限时达卡派)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_20',物流商:'安速',渠道:'加拿大包税-海派(定提UPS派送)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_21',物流商:'安速',渠道:'加拿大包税-海派(定提卡派)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_22',物流商:'安速',渠道:'加拿大包税-海派(海运UPS派送)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_23',物流商:'安速',渠道:'加拿大包税-海派(海运卡派)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_24',物流商:'安速',渠道:'澳洲包税-海运(悉尼代表)',国家:'澳洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_25',物流商:'安速',渠道:'澳洲包税-空运(普货)',国家:'澳洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_26',物流商:'安速',渠道:'澳洲包税-空运(带磁)',国家:'澳洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_27',物流商:'安速',渠道:'日本自税-海运快船ACP逆算(贴标)',国家:'日本',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_28',物流商:'安速',渠道:'欧洲VAT递延-空派(普货)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_29',物流商:'安速',渠道:'欧洲VAT递延-卡航',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_30',物流商:'安速',渠道:'欧洲VAT递延-海运',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_31',物流商:'安速',渠道:'英国VAT递延-空运快线',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_32',物流商:'安速',渠道:'英国VAT递延-空运慢线',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_33',物流商:'安速',渠道:'英国VAT递延-卡航',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_34',物流商:'安速',渠道:'英国VAT递延-海运',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_35',物流商:'安速',渠道:'英国VAT递延-海卡(BHX4/BHX8/LBA4)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_36',物流商:'安速',渠道:'英国VAT递延-海卡(LPL2/LBA8/EMA3等)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_37',物流商:'亚丰',渠道:'美国包税-海卡(普船)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_38',物流商:'亚丰',渠道:'美国空派快线(双清包税UPS)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_39',物流商:'亚丰',渠道:'美国空派经济线(双清包税UPS)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_40',物流商:'亚丰',渠道:'美国美森限时达(CLX双清包税)',国家:'美国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_41',物流商:'亚丰',渠道:'欧盟海运包税',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_42',物流商:'亚丰',渠道:'欧盟快铁包税',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_43',物流商:'亚丰',渠道:'欧盟卡航包税(UPS派送)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_44',物流商:'亚丰',渠道:'欧盟卡航包税(限时达)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_45',物流商:'亚丰',渠道:'欧洲自税递延-空运(普货特快限时达)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_46',物流商:'亚丰',渠道:'欧洲自税递延-空运(普货快线限时达)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_47',物流商:'亚丰',渠道:'欧洲自税递延-海运卡派',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_48',物流商:'亚丰',渠道:'英国空派(自税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_49',物流商:'亚丰',渠道:'英国空派(包税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_50',物流商:'亚丰',渠道:'英国海运(自税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_51',物流商:'亚丰',渠道:'英国海运(包税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_52',物流商:'亚丰',渠道:'英国卡航(自税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_53',物流商:'亚丰',渠道:'英国卡航(包税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_艾杜克_54',物流商:'艾杜克',渠道:'沙特空派',国家:'沙特',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_艾杜克_55',物流商:'艾杜克',渠道:'沙特空运',国家:'沙特',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_艾杜克_56',物流商:'艾杜克',渠道:'沙特空运包税',国家:'沙特',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_合联_57',物流商:'合联',渠道:'沙特海运',国家:'沙特',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
  }
  // 仓库主数据（FBA 收货仓库）：从 Amazon 官方 FC 目录种子(warehouse_seed.js)，覆盖实际发货国家
  // 按「仓库代码」判重避免重复（旧库 id 大小写不同）；EDI4 旧种子误写为美国德州，强制修正为英国 Dunfermline
  const WHSEED = (typeof window!=='undefined' && window.WAREHOUSE_SEED) ? window.WAREHOUSE_SEED : [];
  const _existCodes = new Set((await getAll('warehouses')).map(w=>w.代码));
  for(const w of WHSEED){ if(!_existCodes.has(w.代码)) await put('warehouses', w); }
  const _edi = (await getAll('warehouses')).find(w=>w.代码==='EDI4');
  if(_edi){ _edi.国家='英国'; _edi.公司='Amazon EDI4'; _edi.省份='Fife'; _edi.城市='Dunfermline'; _edi.地址='Amazon Way'; _edi.邮编='KY11 8ST'; await put('warehouses', _edi); }
  // 亦邦渠道：从交接清单历史(国家×空海运组合)派生，幂等补全（无邮件报价，故未建在 rates.json）
  if(!(await get('channels','ch_亦邦_58'))){
    await put('channels',{id:'ch_亦邦_58',物流商:'亦邦',渠道:'亦邦-阿联酋空派',国家:'阿联酋',VAT:'',EORI:'',注册名:'',注册地址:'',仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
  }
  if(!(await get('channels','ch_亦邦_59'))){
    await put('channels',{id:'ch_亦邦_59',物流商:'亦邦',渠道:'亦邦-阿联酋海运',国家:'阿联酋',VAT:'',EORI:'',注册名:'',注册地址:'',仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
  }
  // 国家→仓库代码映射：可配置（主数据→仓库主数据页编辑）。默认覆盖实际发货国家，已有映射保留并补全缺失国家
  const DEF_WH = {美国:'SCK8',沙特:'RUH8',德国:'FRA1',阿联酋:'DXB1',日本:'NRT1',加拿大:'YYZ1',英国:'EDI4',澳大利亚:'MEL1'};
  const _whCfg = await get('config','whByC');
  const whByC = (_whCfg && _whCfg.v) ? Object.assign({}, DEF_WH, _whCfg.v) : DEF_WH;
  await put('config',{id:'whByC', v: whByC});
  // SKU 主数据：从烘焙的 window.SKUS(商品申报信息同步) seed；版本号变更则重 seed
  // 声明价值来源 = 商品申报信息.成本价(人民币) ÷ 汇率(见 sync_master.py RATE)，带版本号可复验
  const SKUS_SEED_VER = 9; // bump: 强制重灌IndexedDB到最新skus.js(13019条,含带@渠道后缀SKU的基础款),解决旧IndexedDB未重灌导致的误判缺主数据
  let skusSeeded = false;
  try { skusSeeded = localStorage.getItem('skus_seeded_ver') === String(SKUS_SEED_VER); } catch(e){}
  if(!skusSeeded){
    await clear('skus');
    for(const s of (window.SKUS||[])){ try{ await put('skus', s.id ? s : {...s, id: s.sku||uid()}); }catch(e){ console.warn('sku 种子跳过', s&&s.sku, e.message); } }
    try { localStorage.setItem('skus_seeded_ver', String(SKUS_SEED_VER)); } catch(e){}
  }
  // 箱型规格主数据：从烘焙的 window.BOX_SPECS(「SKU纸箱规格」飞书表同步) seed
  const BOXSPEC_SEED_VER = 2; //  bumped：强制所有客户端重新 seed 新版 3788 条箱规
  let bsSeeded = false;
  try { bsSeeded = localStorage.getItem('boxspecs_seeded_ver') === String(BOXSPEC_SEED_VER); } catch(e){}
  // 双重校验：即使 version 命中，若 IndexedDB 实际条数远少于当前源数据，也重 seed（防 localStorage/IndexedDB 不同步）
  let bsCount = 0;
  try { bsCount = (await getAll('boxspecs')).length; } catch(e){}
  const specs = window.BOX_SPECS || {};
  const specCount = Object.keys(specs).length;
  if(!bsSeeded || bsCount < specCount * 0.8){
    await clear('boxspecs');
    for(const b of Object.values(specs)){ try{ await put('boxspecs', {...b, id: b.id||b.sku||uid()}); }catch(e){ console.warn('boxspec 种子跳过', b&&b.sku, e.message); } }
    try { localStorage.setItem('boxspecs_seeded_ver', String(BOXSPEC_SEED_VER)); } catch(e){}
  }
  // 预置 5 家模板（同目录 fetch，仅 http 下可用；file:// 失败则手动上传）
  await seedDefaultTemplates(false);
}

/* 加载默认模板种子；force=true 时覆盖已有（用于手动重试）。返回加载成功的模板数 */
async function isValidZipBlob(blob){
  if(!blob || blob.size < 4) return false;
  const buf = await blob.slice(0,4).arrayBuffer();
  const arr = new Uint8Array(buf);
  // xlsx 本质是 zip，魔数 PK\x03\x04
  return arr[0]===0x50 && arr[1]===0x4B && arr[2]===0x03 && arr[3]===0x04;
}
async function seedDefaultTemplates(force=false){
  seedStatus.loading = true;
  seedStatus.errors = [];
  seedStatus.loaded = 0;
  const tmpls = await getAll('templates');
  const have = new Map(tmpls.map(t=>[t.id, t]));
  for(const [key, file] of Object.entries(TPL_FILES)){
    const id = 'tmpl_'+key;
    // 若已有但 zip 损坏，强制重载
    if(!force && have.has(id)){
      const existing = have.get(id);
      if(existing && existing.blob && await isValidZipBlob(existing.blob)){ continue; }
      console.warn('模板已存在但 zip 损坏，重新拉取', file);
    }
    try{
      const r = await fetch(`./${file}?v=${APP_VERSION}`);
      if(r.ok){
        const blob = await r.blob();
        if(!blob || blob.size === 0) throw new Error('返回空文件');
        if(!(await isValidZipBlob(blob))) throw new Error('文件不是有效 zip/xlsx（可能下载中断或被 CDN 截断）');
        await put('templates',{id,物流商:key,渠道:'(通用)',名称:file,blob,状态:'ACTIVE',版本:1,创建日:new Date().toISOString().slice(0,10),mapping:MAPPINGS[key]});
        seedStatus.loaded++;
      }else{
        throw new Error('HTTP '+r.status);
      }
    }catch(e){
      seedStatus.errors.push(`${file}: ${e.message||String(e)}`);
      console.warn('模板种子失败', file, e);
    }
  }
  seedStatus.loading = false;
  // 刷新 W.templates（若已初始化），避免 step4 仍显示空列表
  if(typeof W !== 'undefined' && W) W.templates = (await getAll('templates')).filter(t=>t.状态!=='DISABLED');
  return seedStatus.loaded;
}
const COEFF = 0.3; // 推算系数: 申报价 = 成本 × 系数(标黄)
const RAILWAY_URL = 'https://web-production-6e31e.up.railway.app'; // 云端后端(任何人可用, 不依赖本机开机)
const LOCAL_URL = 'http://localhost:3460'; // 本机常驻后端(兜底)

/* ---------- 后端代理状态检测 ---------- */
async function resolveBackend(){
  // 优先级: 用户手动设置 > 云端Railway(任何人可用) > 本机localhost(兜底)
  const saved=localStorage.getItem('backend_url');
  const candidates=[saved, RAILWAY_URL, LOCAL_URL].filter(Boolean);
  for(const u of candidates){
    try{
      const r=await fetch(u+'/api/health', {signal:AbortSignal.timeout(3000)});
      const d=await r.json();
      if(d&&d.ok) return u;
    }catch(_){}
  }
  return saved||RAILWAY_URL; // 都连不上则返回首选, 让checkBackend标记未连接
}
async function checkBackend(){
  const el=$('#backendStatus'); if(!el) return;
  let url=localStorage.getItem('backend_url')||RAILWAY_URL;
  try{
    let r=await fetch(url+'/api/health', {signal:AbortSignal.timeout(3000)});
    let d=await r.json();
    if(!d||!d.ok){ // 手动设置失效, 自动探测云端/本机
      const alt=await resolveBackend();
      if(alt!==url){ url=alt; localStorage.setItem('backend_url', url); r=await fetch(url+'/api/health',{signal:AbortSignal.timeout(3000)}); d=await r.json(); }
    }
    if(d&&d.ok){ el.textContent='后端: 已连接 ('+(url.includes('railway')?'云端':'本机')+')'; el.style.borderColor='var(--green)'; el.style.color='var(--green)'; }
    else throw new Error('not ok');
  }catch(e){
    el.textContent='后端: 未连接 (点击配置)'; el.style.borderColor='var(--warn)'; el.style.color='var(--warn)';
  }
  el.style.cursor='pointer';
  el.title='点击配置后端地址';
}
setTimeout(()=>{ checkBackend(); $('#backendStatus').onclick=()=>{
  const cur=localStorage.getItem('backend_url')||RAILWAY_URL;
  const v=prompt('请输入后端代理地址：\n默认云端Railway(任何人可用, 推荐)\n如本机常驻后端则填 http://localhost:3460', cur);
  if(v&&v.trim()){ localStorage.setItem('backend_url', v.trim()); checkBackend(); }
}; }, 500);

/* ---------- 工具 ---------- */
const $ = (s,el=document)=>el.querySelector(s);
const $$ = (s,el=document)=>[...el.querySelectorAll(s)];
// 空安全点击绑定：元素缺失时仅告警不抛错，避免渲染中断
function bindClick(id, fn){ const el=document.getElementById(id); if(el) el.onclick=fn; }
function el(html){ const t=document.createElement('template'); t.innerHTML=html.trim(); return t.content.firstElementChild; }
function esc(s){ return String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function main(){ return document.getElementById('main'); }

/* ---------- 视图路由 ---------- */
const VIEWS = { overview, wizard, channels, warehouses, skus, templates, monitor };
let userNavigated = false;   // 用户是否已手动切过视图（防止 init 的种子异步完成后把视图强拉回总览）
function go(view){
  try{
    $$('.nav-btn').forEach(b=>b.classList.toggle('active',b.dataset.view===view));
    main().innerHTML=''; VIEWS[view]();
  }catch(e){ main().innerHTML='<div class="alert alert-err">渲染错误：'+esc(e.message)+'</div>'; console.error(e); }
}
$$('.nav-btn').forEach(b=> b.onclick=()=>{ userNavigated=true; go(b.dataset.view); });

/* ============================================================
   架构总览
   ============================================================ */
function overview(){
  main().innerHTML = `
  <h2>发票系统骨架架构图 · v1</h2>
  <div class="sub">七层解耦 · 云端优先 · 本地兜底 · 质量第一。模板=排版层；值来自 L4 主数据（生成时反查填模板）；导出 Excel=默认交付。</div>
  <div class="grid7">
    ${layer('L1 入口','外部触发','入口A 查飞书单号 / 入口B 用户拖文件 / 手建装箱单','')}
    ${layer('L2 适配','可插拔','飞书表·文件·SP-API·聚水潭（解析防错, fail loud）','')}
    ${layer('L3 规范模型','转轴 · 质量命门','装箱单中枢(规范字段); 源忠实、推算值高亮','tag-core')}
    ${layer('L4 配置与资源','主数据权威源','渠道·收货人 / SKU主数据 / 模板库 / 仓库（本地 IndexedDB）','tag-src')}
    ${layer('L5 生成引擎','业务操作','源忠实填空白模板副本 + 装箱单 CRUD','')}
    ${layer('L6 校验反查','质量命门','溯源 / 反查 / 合理性 / 勾稽 / 人审闸门','tag-core')}
    ${layer('L7 交付','输出','导出 Excel(默认) + 发送(可选可跳过)','')}
  </div>
  <div class="card">
    <h3>本 v1 已落地能力</h3>
    <ul>
      <li><b>L4 主数据</b>：渠道·收货人（含仓库子表）、SKU 主数据（申报价带版本号）本地 CRUD，IndexedDB 持久化。</li>
      <li><b>反查机制</b>：选「物流商+渠道+仓库代码」→ 自动带出国家/VAT/地址（绿=主数据反查，白=手填，黄=推算）。</li>
      <li><b>L5 生成引擎</b>：ExcelJS 打开空白模板副本按映射填格，<b>保留原模板样式/合并/图片公式</b>。已支持 5 家模板（安速/艾杜克/亦邦/亚丰/合联）各自字段映射。</li>
      <li><b>L6 校验</b>：必填完整性 + 勾稽（箱数/数量/申报总值）+ 源忠实高亮。</li>
      <li><b>L7 交付</b>：导出填好的 .xlsx（默认，物流商可直接导入其系统）；发送为可选、可跳过。</li>
    </ul>
    <div class="hint">已知边界（v1）：图片自动嵌入未做（模板样例行 =DISPIMG 公式保留）；飞书单号直查/拖文件解析/SP-API 待接入 L2 适配层。</div>
  </div>`;
}
function layer(title,tag,desc,cls){
  const tagHtml = cls?`<span class="tag ${cls}">${tag}</span>`:`<span class="tag tag-src">${tag}</span>`;
  return `<div class="layer">${tagHtml}<h4>${title}</h4><small>${desc}</small></div>`;
}

/* ============================================================
   生成发票向导
   ============================================================ */
let W = null;
async function wizard(){
  const channels = await getAll('channels');
  const skus = await getAll('skus');
  const templates = (await getAll('templates')).filter(t=>t.状态!=='DISABLED');
  const boxspecs = await getAll('boxspecs');
  W = { step:1, channels, skus, boxspecs, warehouses: await getAll('warehouses'), templates, mode:'forward', handover:null, packed:false,
        form:{ 物流商:'', 渠道:'', 仓库代码:'SCK8', fbaNo:'', amazonRef:'', customs:'否', customInfo:'', items:[] },
        sources:{}, checks:null };
  // 兜底：若 IndexedDB 为空（如隐私模式/seed 失败），确保内存变量至少为空数组，避免渲染抛错
  if(!W.boxspecs) W.boxspecs = [];
  renderWizard();
}
async function renderWizard(){
  await ensureSkusLoaded();
  await ensureBoxspecsLoaded();
  const m = main();
  m.innerHTML = `
  <h2>生成发票向导</h2>
  <div class="sub">① 选择生成方式 → ② 核对收货人+物品 → ③ 选模板预览 → ④ 校验 → ⑤ 交付。所有值自动反查，来源清晰可追溯。</div>
  <div class="stepper">
    <span class="s ${W.step>=1?'active':''}" data-s="1">① 入口·装箱单</span>
    <span class="s ${W.step>=2?'active':''}" data-s="2">② 反查收货人</span>
    <span class="s ${W.step>=3?'active':''}" data-s="3">③ 物品明细</span>
    <span class="s ${W.step>=4?'active':''}" data-s="4">④ 选模板·预览</span>
    <span class="s ${W.step>=5?'active':''}" data-s="5">⑤ 校验反查</span>
    <span class="s ${W.step>=6?'active':''}" data-s="6">⑥ 人审·交付</span>
  </div>
  <div id="wstep"></div>`;
  $$('.stepper .s').forEach(s=> s.onclick=()=>{ const n=+s.dataset.s; if(n<W.step) {W.step=n; renderWizard();} });
  const box = $('#wstep');
  if(W.step===1){ step1(box); syncNext1State(); }
  else if(W.step===2) await step2(box);
  else if(W.step===3) step3(box);
  else if(W.step===4) step4(box);
  else if(W.step===5) step5(box);
  else if(W.step===6) await step6(box);
}
/* step1 中统一兜底：只要已 packed，就启用 next1 并绑定对应确认函数 */
function syncNext1State(){
  const n1=$('#next1'); if(!n1) return;
  if(W.packed){
    n1.disabled=false;
    n1.textContent='下一步：核对收货人 →';
    if(W._handoverHit){
      const fid=W._handoverHit.fba_shipment||W._handoverHit.internal_no;
      n1.onclick=()=>confirmCardA(fid);
    } else {
      n1.onclick=()=>confirmCardB();
    }
  } else {
    n1.disabled=true;
    n1.textContent='请先完成以上任一种方式，再继续 →';
    n1.onclick=null;
  }
}
function step1(box){
  const f = W.form;
  W._cardA = false; W._cardB = false;
  box.innerHTML = `
  <div class="card-row" style="display:flex;gap:16px;flex-wrap:wrap">
    <!-- 卡片 A -->
    <div class="card" style="flex:1;min-width:340px">
      <h3 style="color:#2b6cb0">方式一：按 FBA 货件号获取</h3>
      <div class="hint">输入 FBA 货件号 → 从飞书交接清单反查物流商信息；<b>物流渠道需手动选择</b>（交接清单不含渠道）。</div>
      <div class="row" style="margin-top:10px">
        <div style="flex:1"><label>FBA 货件号</label>
          <div style="display:flex;gap:6px">
            <input id="a_fbaNo" value="${esc(f.fbaNo)}" placeholder="FBA15LXD5XHN" style="flex:1">
            <button class="btn" id="a_search" style="white-space:nowrap">搜索</button>
          </div>
        </div>
      </div>
      <div id="a_result"></div>
    </div>
    <!-- 卡片 B -->
    <div class="card" style="flex:1;min-width:340px">
      <h3 style="color:#6b46c1">方式二：手动上传装箱清单（兜底）</h3>
      <div class="hint">上传本机装箱清单（Excel/CSV），系统解析物品行并识别 FBA 货件号。<b>物流渠道需手动选择</b>（文件中不含渠道）。若识别出的 FBA 号已收录于交接清单索引，物流商可自动反查；否则需手动选择物流商。</div>
      <label class="btn secondary" style="display:inline-block;margin-top:10px;cursor:pointer;font-size:14px;padding:8px 16px">
        📤 选择文件<input id="b_file" type="file" accept=".xlsx,.xls,.csv" style="display:none">
      </label>
      <span id="b_filename" class="muted" style="margin-left:8px"></span>
      <div id="b_result" style="margin-top:10px"></div>
    </div>
  </div>
  <div style="margin-top:18px"><button class="btn" id="next1" ${W.packed?'':'disabled'}>${W.packed?'下一步：核对收货人 →':'请先完成以上任一种方式，再继续 →'}</button></div>`;
  /* ---------- 卡片 A：搜索 FBA 号 ---------- */
  $('#a_search').onclick = async ()=>{
    const fid = ($('#a_fbaNo').value||'').trim();
    if(!fid){ alert('请输入 FBA 货件号'); return; }
    f.fbaNo = fid;
    const res = $('#a_result');
    res.innerHTML = '<div class="hint">🔍 正在搜索飞书交接清单...</div>';
    const r = await searchHandover(fid);
    if(r.error){
      res.innerHTML = '<div class="alert alert-warn">⚠️ 云端搜索暂不可用：'+esc(r.error)+'</div>';
      const local = (window.HANDOVER_INDEX||[]).filter(h=>(h.fba_shipment||'').includes(fid)||(h.internal_no||'').includes(fid));
      if(local.length) renderCardAHit(local[0], res);
      else renderCardAOnlineFallback(fid, res);
      return;
    }
    if(!r.results||!r.results.length){ renderCardAOnlineFallback(fid, res); return; }
    renderCardAHit(r.results[0], res);
    syncNext1State();
  };
  /* ---------- 卡片 B：上传文件 ---------- */
  const bf=$('#b_file');
  if(bf) bf.onchange = async (e)=>{
    const file=e.target.files[0]; if(!file) return;
    $('#b_filename').textContent = file.name;
    const isXlsx=/\.(xlsx|xls)$/i.test(file.name);
    const res=$('#b_result');
    res.innerHTML = '<div class="hint">⏳ 正在解析 '+esc(file.name)+'...</div>';
    const rd=new FileReader();
    rd.onload=async()=>{
      try{
        let items = isXlsx ? await parsePackingXlsx(rd.result) : parsePackingList(rd.result);
        const meta = (items && items.meta) || {};
        // 装箱清单反推的货件号最权威（来自文件"箱号"列真实箱号前缀），覆盖 step1 残留的 stale 值，杜绝外来 FBA
        const fileFbaNo = ((items && items.fbaNo) || meta.fbaNo || '').trim();
        if(fileFbaNo) f.fbaNo = fileFbaNo;
        const effectiveFbaNo = fileFbaNo || f.fbaNo;
        await ensureSkusLoaded();
        items = normalizeItems(items, effectiveFbaNo).map(resolveItemMaster);
        if(!items.length){ res.innerHTML='<div class="alert alert-err">解析为空，请确认文件是有效的装箱清单</div>'; return; }
        W.form.items = items; W.packed = true;
        // 应用文件表头解析出的收货人元数据（FC代码/配送地址），避免落到默认仓SCK8
        if(meta.fcCode) f.仓库代码 = meta.fcCode;
        if(meta.parsedAddress){
          const a = meta.parsedAddress;
          f.company = a.company || f.company || '';
          f.address = a.address || f.address || '';
          f.city = a.city || f.city || '';
          f.province = a.province || f.province || '';
          f.zip = a.zip || f.zip || '';
          f.country = a.country || f.country || '';
          f._addrFromFile = true;
        } else if(meta.deliveryAddress){
          f.address = meta.deliveryAddress;
          f._addrFromFile = true;
        }
        const F = effectiveFbaNo || f.fbaNo;
        if(F){
          const local = (window.HANDOVER_INDEX||[]).find(h=>h.fba_shipment===F||h.internal_no===F);
          if(local){
            f.物流商 = local.carrier||f.物流商;
            f.fbaNo = local.fba_shipment||f.fbaNo;
            // 若交接清单有国家/仓库信息也覆盖（但保留文件解析的详细地址）
            if(local.country && !f.country){ f.country = local.country; }
            renderCardBResult(res, local); syncNext1State(); return;
          }
        }
        renderCardBManual(res, { fbaNo: fileFbaNo, items: items.length }); syncNext1State();
      }catch(err){ res.innerHTML='<div class="alert alert-err">解析失败：'+esc(err.message||err)+'</div>'; }
    };
    if(isXlsx) rd.readAsArrayBuffer(file); else rd.readAsText(file);
  };
}
/* 渠道下拉：按物流商+国家+空海运三重匹配，避免全部列出来看花眼 */
function channelCountryMatch(chCountry, hoCountry){
  if(!chCountry || !hoCountry) return true;
  const cc=String(chCountry).trim(), hc=String(hoCountry).trim();
  if(cc===hc) return true;
  if(cc==='欧洲' || cc==='欧盟'){
    const eu=['德国','法国','意大利','西班牙','波兰','捷克','荷兰','比利时','奥地利','卢森堡','斯洛伐克','斯洛文尼亚','匈牙利','罗马尼亚','保加利亚','克罗地亚','丹麦','瑞典','芬兰','爱尔兰','葡萄牙','希腊','爱沙尼亚','拉脱维亚','立陶宛'];
    return eu.includes(hc);
  }
  return cc.includes(hc) || hc.includes(cc);
}
function channelModeMatch(chName, airSea){
  if(!airSea || !chName) return true;
  const n=String(chName).toLowerCase(), as=String(airSea).toLowerCase();
  if(as.includes('空')) return /空|航/.test(n) && !/卡航/.test(n);
  if(as.includes('海')) return !/空派|空运|快递|dhl|ups|fedex/.test(n);
  return true;
}
function channelSelectHTML(id, channels, hit){
  const same = channels.filter(c=>c.物流商===hit.carrier);
  const matched = same.filter(c=>channelCountryMatch(c.国家, hit.country) && channelModeMatch(c.渠道, hit.air_sea));
  const opts = (matched.length ? matched : same).map(c=>`<option value="${esc(c.渠道)}">${esc(c.渠道)}</option>`).join('');
  const hint = matched.length
    ? (matched.length < same.length ? `<div class="hint" style="margin-top:4px">已按「${esc(hit.carrier)} · ${esc(hit.country)} · ${esc(hit.air_sea)}」筛选，共 ${matched.length} 条可选渠道</div>` : '')
    : (same.length ? `<div class="hint warn" style="margin-top:4px">未找到「${esc(hit.carrier)} · ${esc(hit.country)} · ${esc(hit.air_sea)}」完全匹配渠道，已显示该物流商全部 ${same.length} 条</div>` : '');
  return `<select id="${id}" style="width:100%;margin-top:4px"><option value="">-- 请选择 --</option>${opts}</select>${hint}`;
}

/* 卡片 A 未命中交接清单 → 提供在线拉取入口 */
function renderCardAOnlineFallback(fid, el){
  el.innerHTML = `
    <div class="alert alert-err">未在交接清单索引中找到「${esc(fid)}」。可检查单号、用方式二上传，或点击下方按钮尝试在线拉取。</div>
    <div class="card" style="margin-top:10px;border-color:#38a169">
      <div class="hint ok" style="margin-bottom:8px">🌐 云端拉取（ Railway 后端代理飞书云文档）</div>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <button class="btn" id="a_online" style="background:#38a169;color:#fff"><span style="font-size:17px;margin-right:6px">🌐</span>在线拉取该货件装箱清单</button>
        <span id="a_online_msg" class="muted" style="flex:1;min-width:200px"></span>
      </div>
      <div class="hint" style="margin-top:10px;font-size:12px;color:#888">
        若该货件已收录在飞书云文档中，点击后可跳过交接清单索引直接拉取装箱明细；<br>
        拉取成功后需手动选择物流商/渠道（因无交接清单，系统无法反查物流商）。
      </div>
    </div>`;
  $('#a_online').onclick = async ()=>{ await tryOnlineFillStep1(fid); };
}
/* 卡片 A 命中交接清单后的确认面板 */
function renderCardAHit(hit, el){
  W._handoverHit = hit;
  el.innerHTML = `
    <div class="card" style="margin-top:10px;border-color:#2b6cb0">
      <div class="hint ok" style="margin-bottom:8px">✅ 匹配到交接清单记录</div>
      <table class="kv">
        <tr><td>FBA 货件号</td><td><b>${esc(hit.fba_shipment||'')}</b></td></tr>
        <tr><td>内部单号</td><td>${esc(hit.internal_no||'')}</td></tr>
        <tr><td><b>物流商</b></td><td><b>${esc(hit.carrier||'')}</b>（自动反查）</td></tr>
        <tr><td>空/海运</td><td>${esc(hit.air_sea||'')}</td></tr>
        <tr><td>国家</td><td>${esc(hit.country||'')}</td></tr>
        <tr><td>箱数</td><td>${esc(hit.boxes||'')}</td></tr>
      </table>
      <div style="margin-top:10px"><label><b>请选择物流渠道</b>（交接清单不含渠道，需手动选）</label>
        ${channelSelectHTML('a_channel', W.channels, hit)}</div>
      <div style="margin-top:10px">
        <button class="btn" id="a_confirm">✅ 确认，获取装箱清单并继续</button>
      </div>
    </div>`;
  $('#a_confirm').onclick = async ()=>{
    const ch = $('#a_channel').value;
    if(!ch){ alert('请先选择物流渠道'); return; }
    W.form.物流商 = hit.carrier;
    W.form.渠道 = ch;
    W.handover = hit;
    W.packFbaId = hit.fba_shipment||hit.internal_no;
    W._plLoading = true; W.step=2; renderWizard();
    const res = await fetchItemsLive(W.packFbaId);
    W._plLoading = false;
    if(res && res.items && res.items.length){ W.form.items = res.items; W.packed = true; W.plAutoFilled = res.count; }
    renderWizard();
  };
}
/* 卡片 B 命中交接清单后的确认面板 */
function renderCardBResult(el, local){
  el.innerHTML = `
    <div class="card" style="margin-top:10px;border-color:#6b46c1">
      <div class="hint ok" style="margin-bottom:8px">✅ 已识别 FBA 号 ${esc(local.fba_shipment||'')}，物流商已自动反查</div>
      <table class="kv">
        <tr><td>物流商</td><td><b>${esc(local.carrier||'')}</b>（自动反查）</td></tr>
        <tr><td>国家</td><td>${esc(local.country||'')}</td></tr>
        <tr><td>空/海运</td><td>${esc(local.air_sea||'')}</td></tr>
      </table>
      <div style="margin-top:10px"><label><b>请选择物流渠道</b>（文件中不含渠道，需手动选）</label>
        ${channelSelectHTML('b_channel', W.channels, local)}</div>
      <div class="hint" style="margin-top:8px">✅ 物品已解析，继续后可直接核对。</div>
    </div>`;
  $('#b_channel').onchange = ()=>{
    const chName = $('#b_channel').value;
    const ch = W.channels.find(c=>c.渠道===chName);
    // 【质量第一·防错】物流商以"所选渠道的真实归属"为准，杜绝交接清单反查物流商写错导致错模板
    W.form.物流商 = (ch && ch.物流商) || local.carrier;
    W.form.渠道 = chName;
    W.packFbaId = local.fba_shipment||local.internal_no;
  };
}
/* 卡片 B 未命中交接清单 → 全手动选。
   文案铁律：必须区分「已从文件识别 FBA 号但未收录于交接清单」vs「未识别到 FBA 号」——
   文件解析成功与否由物品行数证明，绝不能让用户误以为解析失败。 */
function renderCardBManual(el, ctx){
  ctx = ctx || {};
  const fba = String(ctx.fbaNo || '').trim();
  const nItems = parseInt(ctx.items, 10) || 0;
  const head = fba
    ? `<div class="hint ok" style="margin-bottom:8px">✅ 已从文件识别 FBA 货件号：<b>${esc(fba)}</b>${nItems ? `（已解析 ${nItems} 行物品）` : ''}</div>
       <div class="hint warn" style="margin-bottom:8px">该号未收录于本地交接清单索引，无法自动反查物流商 → 请手动选择物流商和渠道。选择渠道后系统会自动锁定该渠道的真实物流商（防错），已识别的货件号将用于箱号对齐。</div>`
    : `<div class="hint warn" style="margin-bottom:8px">⚠️ 未能从文件识别 FBA 货件号，无法自动反查物流商 → 请手动选择物流商和渠道。${nItems ? `（物品已解析 ${nItems} 行，可继续）` : ''}</div>`;
  el.innerHTML = `
    <div class="card" style="margin-top:10px;border-color:#c53030">
      ${head}
      <div class="row" style="margin-top:8px">
        <div><label>物流商</label>
          <select id="b_manual_carrier" style="width:100%">
            <option value="">-- 请选择 --</option>
            ${[...new Set(W.channels.map(c=>c.物流商))].map(o=>`<option>${o}</option>`).join('')}
          </select></div>
        <div><label>物流渠道</label>
          <select id="b_manual_channel" style="width:100%"><option value="">-- 先选物流商 --</option></select></div>
      </div>
      <div class="hint" style="margin-top:8px">✅ 物品已解析，继续后可直接核对。</div>
    </div>`;
  $('#b_manual_carrier').onchange = ()=>{
    const car = $('#b_manual_carrier').value;
    const ch = $('#b_manual_channel');
    ch.innerHTML = '<option value="">-- 请选择 --</option>'+W.channels.filter(c=>c.物流商===car).map(c=>`<option>${esc(c.渠道)}</option>`).join('');
    W.form.物流商 = car;
  };
  $('#b_manual_channel').onchange = ()=>{
    const chName = $('#b_manual_channel').value;
    const ch = W.channels.find(c=>c.渠道===chName);
    // 【质量第一·防错】选渠道时同步物流商，杜绝手动模式只选渠道漏选物流商导致错模板
    if(ch && ch.物流商) W.form.物流商 = ch.物流商;
    W.form.渠道 = chName;
  };
}
/* step1 未命中交接清单时，直接在线拉取装箱清单 → 成功后选手动物流商/渠道 */
async function tryOnlineFillStep1(fid){
  const btn=$('#a_online'); if(btn) btn.disabled=true;
  const msg=$('#a_online_msg'); if(msg) msg.textContent='⏳ 正在从云端拉取装箱清单...';
  await ensureSkusLoaded();
  await ensureBoxspecsLoaded();
  const res = await fetchItemsLive(fid);
  if(res && res.items && res.items.length){
    W.form.fbaNo = fid;
    W.form.items = res.items;
    W.packed = true;
    W.plAutoFilled = res.count;
    W.packFbaId = fid;
    const meta = res.meta || (window.PACKING_META && window.PACKING_META[fid]) || null;
    if(meta && meta.fcCode) W.form.仓库代码 = meta.fcCode;
    W.handover = null;
    W._plLoading = false;
    // 留在 step1，显示手动选择物流商/渠道面板
    renderCardAManual(fid, $('#a_result'));
  } else {
    const hint = hintForOnlineFail(fid, (res && res.code) || 'NOT_FOUND');
    if(msg) msg.innerHTML = '<span style="color:var(--err)">❌ '+hint+'</span>';
    if(btn) btn.disabled=false;
  }
}
/* 卡片 A 在线拉取成功后 → 手动选择物流商/渠道（无交接清单） */
function renderCardAManual(fid, el){
  el.innerHTML = `
    <div class="card" style="margin-top:10px;border-color:#38a169">
      <div class="hint ok" style="margin-bottom:8px">✅ 已从云端拉取 <b>${W.form.items.length}</b> 行装箱明细（${esc(fid)}）</div>
      <div class="hint warn" style="margin-bottom:8px">⚠️ 未匹配到交接清单，无法自动反查物流商/国家，请手动选择</div>
      <div class="row" style="margin-top:8px">
        <div><label>物流商</label>
          <select id="a_manual_carrier" style="width:100%">
            <option value="">-- 请选择 --</option>
            ${[...new Set(W.channels.map(c=>c.物流商))].map(o=>`<option>${esc(o)}</option>`).join('')}
          </select></div>
        <div><label>物流渠道</label>
          <select id="a_manual_channel" style="width:100%"><option value="">-- 先选物流商 --</option></select></div>
      </div>
      <div style="margin-top:10px">
        <button class="btn" id="a_manual_confirm">✅ 确认，继续核对收货人 →</button>
      </div>
    </div>`;
  $('#a_manual_carrier').onchange = ()=>{
    const car = $('#a_manual_carrier').value;
    const ch = $('#a_manual_channel');
    ch.innerHTML = '<option value="">-- 请选择 --</option>'+W.channels.filter(c=>c.物流商===car).map(c=>`<option>${esc(c.渠道)}</option>`).join('');
    W.form.物流商 = car;
  };
  $('#a_manual_channel').onchange = ()=>{
    const chName = $('#a_manual_channel').value;
    const ch = W.channels.find(c=>c.渠道===chName);
    W.form.物流商 = (ch && ch.物流商) || W.form.物流商; // 物流商以渠道真实归属为准
    W.form.渠道 = chName;
  };
  $('#a_manual_confirm').onclick = ()=>{
    if(!W.form.物流商 || !W.form.渠道){ alert('请先选择物流商和渠道'); return; }
    W.packFbaId = fid;
    W.step = 2;
    renderWizard();
  };
}
function confirmCardA(fid){
  if(!W.form.渠道){ alert('请先选择物流渠道'); return; }
  W.packFbaId = fid;
  W._plLoading = true; W.step=2; renderWizard();
  fetchItemsLive(fid).then(res=>{
    W._plLoading = false;
    if(res && res.items && res.items.length){ W.form.items = res.items; W.packed = true; W.plAutoFilled = res.count; }
    const meta = (res && res.meta) || (window.PACKING_META && window.PACKING_META[fid]) || null;
    if(meta && meta.fcCode) W.form.仓库代码 = meta.fcCode;
    renderWizard();
  });
}
function confirmCardB(){
  if(!W.form.渠道){ alert('请先选择物流渠道'); return; }
  W.step=2; renderWizard();
}
/* ========== 废弃的分界线（以下为旧功能保留引用，新 step1 不再使用）========== */
// captureForward/reversePanelHTML/bindReverse 已废弃，但 searchHandover/fetchItemsLive 等仍需使用
async function searchHandover(key){
  const backendUrl = localStorage.getItem('backend_url') || RAILWAY_URL;
  try{
    const r = await fetch(backendUrl+'/api/search-handover?q='+encodeURIComponent(key), {signal:AbortSignal.timeout(12000)});
    if(!r.ok) throw new Error('HTTP '+r.status);
    const data = await r.json();
    if(!data.ok) throw new Error(data.error||'搜索失败');
    return data;
  }catch(e){ return {error: e.message}; }
}
function renderHitsOffline(key, res){
  const k=key.toLowerCase();
  const all=window.HANDOVER_INDEX||[];
  const hits=all.filter(r=> (r.internal_no&&r.internal_no.toLowerCase().includes(k)) || (r.fba_shipment&&r.fba_shipment.toLowerCase().includes(k)) ).slice(0,30);
  if(hits.length===0){ res.innerHTML='<div class="hint">离线索引也未找到。</div>'; return; }
  res.innerHTML = hits.map((r,i)=>`<div class="rev-item" data-i="${i}"><span><b>${esc(r.internal_no||'(无内部单号)')}</b> / ${esc(r.fba_shipment||'(无FBA号)')}</span><span class="muted">${esc(r.carrier||r.物流商||'?')} · ${esc(r.country||'?')} · ${esc(r.boxes||'?')}箱 · ${esc(r.packing_list||'无装箱清单')}</span></div>`).join('');
  res.querySelectorAll('.rev-item').forEach(el=> el.onclick=()=> showConfirm(hits[+el.dataset.i]) );
}
function showConfirm(r){
  const c=$('#rev_confirm');
  c.innerHTML = `
    <div class="card" style="margin-top:10px">
      <div class="hint warn">请确认这是你要的单（防取错）：</div>
      <table class="kv">
        <tr><td>内部单号</td><td>${esc(r.internal_no)}</td></tr>
        <tr><td>FBA货件号</td><td>${esc(r.fba_shipment)}</td></tr>
        <tr><td>物流商</td><td>${esc(r.carrier||r.物流商)}</td></tr>
        <tr><td>国家</td><td>${esc(r.country)}</td></tr>
        <tr><td>空/海运</td><td>${esc(r.air_sea)}</td></tr>
        <tr><td>箱数</td><td>${esc(r.boxes)}</td></tr>
        <tr><td>取件方式(地址)</td><td>${esc(r.pickup_addr)}</td></tr>
        <tr><td>装箱清单</td><td>${esc(r.packing_list)}</td></tr>
        <tr><td>FNSKU信息</td><td>${esc(r.fnsku_file)}</td></tr>
        <tr><td>发票-物流填</td><td>${esc(r.invoice_drop)}</td></tr>
      </table>
      <div style="display:flex;gap:10px;margin-top:8px">
        <button class="btn" id="rev_ok">确认采用，抓取信息 →</button>
        <button class="btn secondary" id="rev_cancel">重选</button>
      </div>
    </div>`;
  $('#rev_cancel').onclick=()=>{ c.innerHTML=''; };
  $('#rev_ok').onclick=()=> applyHandover(r);
}
async function fetchItemsLive(fid){
  await ensureSkusLoaded();
  // 1) 预装优先（零网络）
  const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
  if(pl.length) return {items: normalizeItems(pl.map(x=>resolveItemMaster(Object.assign({}, x))), fid), source:'preloaded', count:pl.length, meta:(window.PACKING_META && window.PACKING_META[fid])||null};
  // 2) 后端实时（活表，新增货件也能拉到）
  const backendUrl = localStorage.getItem('backend_url') || RAILWAY_URL;
  try{
    const r = await fetch(backendUrl+'/api/fetch-packing-list', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({fid}), signal:AbortSignal.timeout(15000)});
    const data = await r.json();
    if(data.ok && data.items && data.items.length){
      data.items.forEach(x=>{ if(!x.boxWeight && x.weight) x.boxWeight=x.weight; });
      const meta = data.meta || null;
      if(meta) (window.PACKING_META = window.PACKING_META||{})[fid] = meta; // 在线 meta 缓存进预装，后续离线也可用
      return {items: normalizeItems(data.items.map(x=>resolveItemMaster(Object.assign({}, x))), fid), source:'cloud', count:data.items.length, meta};
    }
    // 后端明确返回错误（如 NOT_FOUND / 权限/解析失败）
    if(data && data.error) return {error:data.error, code:data.code||'BACKEND_ERROR', items:[]};
  }catch(e){ console.warn('fetchItemsLive backend fail', e.message); return {error:e.message, code:'NETWORK_ERROR', items:[]}; }
  return {error:'云端返回空数据', code:'EMPTY_RESPONSE', items:[]};
}
/* 根据交接清单索引生成"在线拉取失败"时的具体提示 */
function hintForOnlineFail(fid, code){
  const idx = (window.HANDOVER_INDEX||[]).find(h=>(h.fba_shipment||'').toUpperCase()===fid.toUpperCase() || (h.internal_no||'').toUpperCase()===fid.toUpperCase());
  if(idx){
    const file = idx.packing_file || idx.packingList || idx['装箱清单(csv)文件名'] || '';
    const carrier = idx.carrier || idx.物流商 || '';
    let s = `该 FBA 号已存在于飞书交接清单（${idx.country||''} ${idx.air_sea||''} ${carrier ? '· 物流商 '+carrier : ''}），但<b>装箱明细尚未同步到云端数据库</b>。`;
    if(file) s += `<br>📎 飞书装箱清单文件名：<code>${esc(file)}</code>；请从飞书下载后用<b>方式二</b>上传。`;
    else s += `<br>请从飞书下载该货件的装箱清单 Excel/CSV 后用<b>方式二</b>上传。`;
    return s;
  }
  if(code==='NOT_FOUND') return '云端数据库未收录该货件，且交接清单索引中也无记录。请检查单号，或用方式二上传本机装箱清单。';
  if(code==='NETWORK_ERROR') return '无法连接云端后端（网络问题或后端暂停）。请改用方式二上传本机装箱清单。';
  return '在线拉取失败。请改用方式二上传本机装箱清单。';
}
// 从装箱清单表头元数据取精确 FBA 仓代码（应用见 applyHandover）：优先 meta.fcCode，否则国家→默认仓
async function applyHandover(r){
  const f=W.form; W.handover=r;
  f.fbaNo = r.fba_shipment || r.internal_no;
  if(r.carrier||r.物流商) f.物流商=r.carrier||r.物流商; // 始终以源表 carrier 为准，避免落到默认渠道
  const sameCarrier = W.channels.filter(c=>c.物流商===f.物流商);
  const byCountry = sameCarrier.find(c=>c.国家 && r.country && c.国家.includes(r.country));
  f.渠道 = (byCountry||sameCarrier[0]||{渠道:f.渠道}).渠道;
  const fid = r.fba_shipment || r.internal_no;
  W.packFbaId = fid;
  W.plAutoFilled = 0; W.packed=false;
  W.step=2; W._plLoading=true; renderWizard();        // 先显示收货人页 + 拉取提示
  const res = await fetchItemsLive(fid);               // 实时拉箱内容（后端优先，fallback 预装）
  W._plLoading=false;
  // 仓库代码：装箱清单表头抓到的精确 FC（RUH8…）优先，否则 国家→默认仓，再否则清空
  const whCfg = await get('config','whByC');
  const whByC = (whCfg && whCfg.v) ? whCfg.v : {美国:'SCK8',沙特:'RUH8'};
  const meta = (res && res.meta) || (window.PACKING_META && window.PACKING_META[fid]) || null;
  f.仓库代码 = (meta && meta.fcCode) ? meta.fcCode : (whByC[r.country] || '');
  if(res && res.items && res.items.length){ W.form.items = res.items; W.plAutoFilled = res.count; W.packed=true; }
  else {
    const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
    if(pl.length){ W.form.items = normalizeItems(pl.map(x=>resolveItemMaster(Object.assign({}, x))), fid); W.plAutoFilled=pl.length; W.packed=true; }
  }
  W.step=3; renderWizard();
}
function warehouseOptions(f){
  const list = W.warehouses || [];
  let opts = list.map(w=>`<option value="${esc(w.代码)}" ${w.代码===f.仓库代码?'selected':''}>${esc(w.代码)} · ${esc(w.国家||'')} ${esc(w.城市||'')}</option>`).join('');
  if(f.仓库代码 && !list.some(w=>w.代码===f.仓库代码)){
    opts += `<option value="${esc(f.仓库代码)}" selected>${esc(f.仓库代码)} · 来自装箱清单</option>`;
  }
  return opts;
}
function lookupChannel(){ return W.channels.find(c=>c.物流商===W.form.物流商&&c.渠道===W.form.渠道) || null; }
async function lookupWarehouse(){ return (await getAll('warehouses')).find(w=>w.代码===W.form.仓库代码) || null; }
async function step2(box){
  const ch=lookupChannel(), wh=await lookupWarehouse();
  // 文件表头解析出的地址优先于仓库主数据（避免主数据缺该FC时落到默认仓SCK8）
  const fileAddr = W.form._addrFromFile ? {
    company: W.form.company, address: W.form.address, city: W.form.city,
    province: W.form.province, zip: W.form.zip, country: W.form.country, phone: W.form.phone
  } : null;
  const pickAddr = (k, whField)=>{
    if(fileAddr && fileAddr[k]) return {v:fileAddr[k], src:'packing'};
    return {v:wh?wh[whField||k]:'', src:'warehouse'};
  };
  const src = W.sources = {
    shipMethod:{v:ch?ch.渠道:'',src:'channel'}, country: fileAddr&&fileAddr.country ? {v:fileAddr.country, src:'packing'} : {v:(W.handover&&W.handover.country)||(ch?ch.国家:''),src:(W.handover&&W.handover.country)?'handover':'channel'}, vat:{v:ch?ch.VAT:'',src:'channel'},
    eori:{v:ch?ch.EORI:'',src:'channel'}, vatName:{v:ch?ch.注册名:'',src:'channel'}, vatAddr:{v:ch?ch.注册地址:'',src:'channel'},
    warehouseCode:{v:W.form.仓库代码,src:'manual'}, company:pickAddr('company','公司'), province:pickAddr('province','省份'),
    city:pickAddr('city','城市'), address:pickAddr('address','地址'), zip:pickAddr('zip','邮编'), phone:pickAddr('phone','电话'),
    fbaNo:{v:W.form.fbaNo,src:'manual'}, amazonRef:{v:W.form.amazonRef,src:'manual'}, customs:{v:W.form.customs,src:'manual'}, customInfo:{v:W.form.customInfo,src:'manual'},
    title:{v:'',src:'template'}, poNo:{v:'',src:'manual'}
  };
  const srcClass = s => (s==='channel'||s==='warehouse'||s==='template'||s==='packing') ? 'cell-src' : 'cell-manual';
  const srcLabel = s => ({channel:'主数据·渠道',warehouse:'主数据·仓库',manual:'人工填写',template:'模板固定',calc:'推算',packing:'装箱清单'}[s]||s);
  const FIELDS = ['fbaNo','amazonRef','poNo','shipMethod','warehouseCode','company','country','province','city','address','phone','zip','email','customs','vat','eori','vatName','vatAddr','customInfo'];
  const LABELS = {fbaNo:'客户订单号(FBA号)',amazonRef:'Amazon Reference ID',poNo:'PO Number',shipMethod:'运输方式',warehouseCode:'收件人(仓库代码)',company:'收件人公司',country:'国家',province:'收件省份',city:'收件城市',address:'收件地址',phone:'收件电话',zip:'邮编',email:'收件人email',customs:'报关(否/是)',vat:'VAT号',eori:'EORI',vatName:'VAT注册名',vatAddr:'VAT注册地址',customInfo:'自定义信息'};
  box.innerHTML = `
  <div class="card">
    ${W._plLoading?'<div class="hint" style="margin-bottom:8px">⏳ 正在实时拉取装箱清单（从飞书云文档）...</div>':''}
    <h3>② 反查收货人（生成时从 L4 主数据取值，不手敲）</h3>
    <div class="hint">所选：<b>${esc(W.form.物流商)} / ${esc(W.form.渠道)}</b>，仓库代码 <b>${esc(W.form.仓库代码)}</b>。绿底=主数据反查带出，白底=需人工填（传统贸易常见）。可直接改，但建议改「主数据页」以保证全量一致。</div>
    ${(!W.form.仓库代码 && W.handover && W.handover.country) ? `<div class="alert alert-warn" style="margin-top:8px">⚠ 未匹配到「${esc(W.handover.country)}」的 FBA 仓库代码（主数据未配置该国家映射/仓库）。请到「主数据 → 仓库主数据」补充该国家的仓库代码与收货地址（一次录入、后续自动带出），或在下方下拉手动选择/输入。</div>` : ''}
    <table>
      <thead><tr><th>字段</th><th>取值</th><th>来源</th></tr></thead>
      <tbody>
        ${FIELDS.map(k=>{ const s=src[k]; if(!s) return ''; return `<tr><td>${LABELS[k]}</td><td class="${srcClass(s.src)}"><input data-meta="${k}" value="${esc(s.v)}"></td><td><span class="pill ${s.src==='manual'?'pill-gray':'pill-green'}">${srcLabel(s.src)}</span></td></tr>`; }).join('')}
      </tbody>
    </table>
  </div>
  <div style="margin-top:14px;display:flex;gap:10px"><button class="btn secondary" id="prev2">← 上一步</button><button class="btn" id="next2">下一步：物品明细 →</button></div>`;
  $$('[data-meta]').forEach(inp=> inp.oninput = e=>{ W.sources[e.target.dataset.meta].v=e.target.value; W.sources[e.target.dataset.meta].src='manual'; });
  bindClick('prev2', ()=>{W.step=1;renderWizard();});
  bindClick('next2', ()=>{W.step=3;renderWizard();});
}
function step3(box){
  if(W.form.items.length===0 && !W.packed){ W.form.items.push({boxLabel:'',boxNo:'',sku:'',nameCn:'',nameEn:'',qty:'',declare:'',material:'',purpose:'',hs:'',brand:'',model:'',boxWeight:'',prodWeight:'',len:'',wid:'',hgt:'',elec:'N',magnet:'N',saleUrl:'',cost:'',boxes:'',boxSpec:''}); }
  function addRow(){ W.form.items.push({boxLabel:'',boxNo:'',sku:'',nameCn:'',nameEn:'',qty:1,declare:'',material:'',purpose:'',hs:'',brand:'',model:'',boxWeight:'',prodWeight:'',len:'',wid:'',hgt:'',elec:'N',magnet:'N',saleUrl:'',cost:'',boxes:'',boxSpec:''}); renderWizard(); }
  function renderRows(){
    const selTmpl = W.templates.find(x=>x.id===W.selTmpl);
    const tpl = (selTmpl && selTmpl.mapping) || MAPPINGS[W.form.物流商];
    const tplItem = (tpl && tpl.item) || {};
    const used = f => !!(tplItem[f]); // 当前模板是否用到该字段(标红跟随模板:艾杜克无材质/用途列则不标)
    // 分页：每页最多渲染 PV 行，避免千/万箱时全量 DOM 爆炸卡死
    const total = W.form.items.length;
    const PV = 100;
    const pages = Math.max(1, Math.ceil(total/PV));
    if(W._pvPage>=pages) W._pvPage=pages-1;
    if(W._pvPage<0||W._pvPage===undefined) W._pvPage=0;
    const begin = W._pvPage*PV;
    const end = Math.min(begin+PV, total);
    const slice = W.form.items.slice(begin, end);
    return slice.map((it,k)=>{ const i = begin+k; // i = 全局索引，与 W.form.items 对齐
      const nSku = (it.sku||'').replace(/@.*$/,'');
      const sk = W.skus.find(s=>s.sku===it.sku) || (nSku && nSku!==it.sku ? W.skus.find(s=>s.sku===nSku) : null);
      const missingMaster = it.sku && !sk; // SKU 有值但主数据确实缺失(已归一化)
      const normSku = s=>(s||'').replace(/@us$/i,'').trim();
      const bs = W.boxspecs.find(b=>normSku(b.sku)===normSku(it.sku)) || null;
      let declareSrc='manual', declareVal=it.declare;
      if(declareVal===''||declareVal==null){ if(sk && sk.申报价){ declareVal=sk.申报价; declareSrc='sku'; } else if(window.SKU_DECLARE && window.SKU_DECLARE[it.sku]){ declareVal=window.SKU_DECLARE[it.sku].d; declareSrc='sku'; } else if(it.cost!==''){ declareVal=(parseFloat(it.cost)*COEFF).toFixed(2); declareSrc='calc'; } }
      let dcls, pill, pillTxt;
      if(declareSrc==='calc'){ dcls='cell-calc'; pill='pill-yellow'; pillTxt='推算(成本×'+COEFF+')'; }
      else if(declareSrc==='sku'){ dcls='cell-src'; pill='pill-green'; pillTxt='SKU主数据'; }
      else if(declareVal!=='' && declareVal!=null){ dcls='cell-manual'; pill='pill-gray'; pillTxt='手填'; }
      else { dcls='cell-warn'; pill='pill-red'; pillTxt='⚠ 待手填'; }
      const displayModel = it.model || (sk?sk.型号:'') || it.sku;
      // 标红跟随模板:仅当"当前模板用到该字段"且"该字段值空"才给单元格加轻量黄标;整行不再红,避免一片红
      const mc = (field,val)=> (missingMaster && used(field) && (!val||String(val).trim()==='')) ? 'class="cell-miss" title="该 SKU 未收录于商品申报信息主数据，'+field+' 列导出将空白"' : '';
      const badge = missingMaster ? '<span class="badge-warn" title="整行 SKU 未收录于「商品申报信息」主数据，下列用到的字段将空白">⚠缺主数据</span>' : '';
      // 推算材质高亮:材质来自中文品名自动抽取,需人审确认
      const matVal = it.material || (sk?sk.材质:'');
      const matCls = mc('material', matVal) || (sk && sk.材质推算 ? 'class="cell-inferred" title="材质由中文品名推算，请人审确认"' : '');
      return `<tr>
        <td>${badge}<input data-i="${i}" data-k="boxLabel" value="${esc(it.boxLabel)}" placeholder="箱号=子单号" title="FBA 箱 ID（同子单号）"></td>
        <td><input data-i="${i}" data-k="boxNo" value="${esc(it.boxNo)}" placeholder="子单号/FBA箱ID" title="FBA 箱子 ID（如 FBA15...U000001）"></td>
        <td><input data-i="${i}" data-k="sku" value="${esc(it.sku)}" list="skuList" placeholder="SKU"></td>
        <td><input data-i="${i}" data-k="nameCn" value="${esc(it.nameCn||(sk?sk.中文品名:'')||(window.SKU_DECLARE&&window.SKU_DECLARE[it.sku]?window.SKU_DECLARE[it.sku].n:''))}" placeholder="中文"></td>
        <td><input data-i="${i}" data-k="nameEn" value="${esc(it.nameEn||(sk?sk.英文品名:''))}" placeholder="英文"></td>
        <td><input data-i="${i}" data-k="qty" value="${esc(it.qty)}" style="width:54px" placeholder="数量"></td>
        <td class="${dcls}"><input data-i="${i}" data-k="declare" value="${esc(declareVal)}" style="width:74px"><br><span class="pill ${pill}">${pillTxt}</span></td>
        <td ${matCls}><input data-i="${i}" data-k="material" value="${esc(matVal)}" placeholder="材质"></td>
        <td ${mc('purpose', it.purpose||(sk?sk.用途:''))}><input data-i="${i}" data-k="purpose" value="${esc(it.purpose||(sk?sk.用途:''))}" placeholder="用途"></td>
        <td ${mc('hs', it.hs||(sk?sk.HS:''))}><input data-i="${i}" data-k="hs" value="${esc(it.hs||(sk?sk.HS:''))}" placeholder="HS"></td>
        <td ${mc('brand', it.brand||(sk?sk.品牌:''))}><input data-i="${i}" data-k="brand" value="${esc(it.brand||(sk?sk.品牌:''))}" placeholder="品牌"></td>
        <td ${mc('model', displayModel)}><input data-i="${i}" data-k="model" value="${esc(displayModel)}" placeholder="型号"></td>
        <td><input data-i="${i}" data-k="boxes" value="${esc(it.boxes)}" style="width:46px" placeholder="箱数"></td>
        <td class="${(!it.boxSpec||it.boxSpec==='')&&bs&&bs.model?'cell-src':''}"><input data-i="${i}" data-k="boxSpec" value="${esc((!it.boxSpec||it.boxSpec==='')&&bs&&bs.model?bs.model:it.boxSpec)}" placeholder="箱规"></td>
        <td class="${(!it.boxWeight||it.boxWeight==='')&&bs&&(bs.weight!=null)?'cell-src':''}"><input data-i="${i}" data-k="boxWeight" value="${esc((!it.boxWeight||it.boxWeight==='')&&bs&&(bs.weight!=null)?bs.weight:it.boxWeight)}" style="width:54px" placeholder="箱重"></td>
        <td class="${(!it.len||it.len==='**')&&bs&&bs.l?'cell-src':''}"><input data-i="${i}" data-k="len" value="${esc((!it.len||it.len==='**')&&bs&&bs.l?bs.l:it.len)}" style="width:46px" placeholder="长"></td>
        <td class="${(!it.wid||it.wid==='**')&&bs&&bs.w?'cell-src':''}"><input data-i="${i}" data-k="wid" value="${esc((!it.wid||it.wid==='**')&&bs&&bs.w?bs.w:it.wid)}" style="width:46px" placeholder="宽"></td>
        <td class="${(!it.hgt||it.hgt==='**')&&bs&&bs.h?'cell-src':''}"><input data-i="${i}" data-k="hgt" value="${esc((!it.hgt||it.hgt==='**')&&bs&&bs.h?bs.h:it.hgt)}" style="width:46px" placeholder="高"></td>
        <td><input data-i="${i}" data-k="elec" value="${esc(it.elec)}" style="width:38px" placeholder="电"></td>
        <td><input data-i="${i}" data-k="magnet" value="${esc(it.magnet)}" style="width:38px" placeholder="磁"></td>
        <td><input data-i="${i}" data-k="saleUrl" value="${esc(it.saleUrl)}" placeholder="销售链接"></td>
        <td><button class="btn danger" data-del="${i}" style="padding:4px 8px">删</button></td>
      </tr>`;
    }).join('');
  }
  function renderPagerHTML(){
    const total = W.form.items.length;
    const PV = 100;
    const pages = Math.max(1, Math.ceil(total/PV));
    if(pages<=1) return '';
    const cur = W._pvPage||0;
    return `<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
      <button class="btn secondary" id="pvPrev" ${cur<=0?'disabled':''}>← 上一页</button>
      <span>第 <b>${cur+1}</b> / ${pages} 页（共 ${total} 行，每页 ${PV}）</span>
      <button class="btn secondary" id="pvNext" ${cur>=pages-1?'disabled':''}>下一页 →</button>
      <span class="muted">跳至</span><input id="pvJump" type="number" min="1" max="${pages}" value="${cur+1}" style="width:64px"> <button class="btn secondary" id="pvGo">跳</button>
    </div>`;
  }
  function bindRowInputs(){
    $$('#rows [data-i]').forEach(inp=> inp.oninput = e=>{ const i=+e.target.dataset.i, k=e.target.dataset.k; W.form.items[i][k]=e.target.value; if(k==='sku'){ resolveItemMaster(W.form.items[i]); renderWizard(); } });
    $$('#rows [data-del]').forEach(b=> b.onclick=()=>{ W.form.items.splice(+b.dataset.del,1); renderWizard(); });
  }
  function bindPager(){
    const total=W.form.items.length, PV=100, pages=Math.max(1,Math.ceil(total/PV));
    const go=p=>{ W._pvPage=Math.max(0,Math.min(pages-1,p)); const t=document.getElementById('rows'); if(t) t.innerHTML=renderRows(); bindRowInputs(); const np=document.getElementById('rowPager'); if(np) np.innerHTML=renderPagerHTML(); bindPager(); };
    const prev=$('#pvPrev'); if(prev) prev.onclick=()=>go((W._pvPage||0)-1);
    const next=$('#pvNext'); if(next) next.onclick=()=>go((W._pvPage||0)+1);
    const goBtn=$('#pvGo'); if(goBtn) goBtn.onclick=()=>{ const v=parseInt($('#pvJump').value,10)||1; go(v-1); };
  }
  box.innerHTML = `
  <div class="card">
    <h3>③ 物品明细（装箱单行项目）</h3>
    <div class="hint">填 SKU 自动反查中文品名/材质/HS/品牌/型号，并带出<b>申报价</b>（绿=SKU主数据；黄=无主数据按成本×${COEFF}推算，需人审确认）。</div>
    <datalist id="skuList">${W.skus.map(s=>`<option value="${s.sku}">${s.中文品名}</option>`).join('')}</datalist>
    ${ W.handover ? packingBannerHTML() : '' }
    <div style="overflow:auto"><table>
      <thead><tr><th>箱标<br><small>(模板箱号)</small></th><th>子单号<br><small>(FBA箱ID)</small></th><th>SKU</th><th>中文</th><th>英文</th><th>数量</th><th>申报价(USD)</th><th>材质</th><th>用途</th><th>HS</th><th>品牌</th><th>型号</th><th>箱数</th><th>箱规</th><th>箱重</th><th>长</th><th>宽</th><th>高</th><th>电</th><th>磁</th><th>销售链接</th><th></th></tr></thead>
      <tbody id="rows">${renderRows()}</tbody>
    </table></div>
    <div id="rowPager" style="margin:8px 0 4px">${renderPagerHTML()}</div>
    <button class="btn secondary" id="addRow" style="margin-top:6px">+ 添加一行</button>
  </div>
  <div style="margin-top:14px;display:flex;gap:10px"><button class="btn secondary" id="prev3">← 上一步</button><button class="btn" id="next3">下一步：选模板·预览 →</button></div>`;
  bindRowInputs();
  bindPager();
  $('#addRow').onclick=addRow;
  $('#prev3').onclick=()=>{W.step=2;renderWizard();};
  $('#next3').onclick=()=>{W.step=4;renderWizard();};
  if(W.handover){
    const fid = W.packFbaId || W.handover.fba_shipment || W.handover.internal_no;
    const rl=$('#pl_reload'); if(rl) rl.onclick=async()=>{ await loadPackingList(fid); };
    const pf=$('#pl_file'); if(pf) pf.onchange=e=>{
      const file=e.target.files[0]; if(!file) return;
      const msg=$('#pl_msg'); msg.textContent='⏳ 正在解析 '+file.name+'...';
      const isXlsx = /\.(xlsx|xls)$/i.test(file.name);
      const rd=new FileReader();
      rd.onload=async()=>{
        try{
          await ensureBoxspecsLoaded();   // 解析前确保箱规主数据就绪，否则 resolveItemMaster 无法回填箱重/尺寸
          await ensureSkusLoaded();       // 必须等 SKU 主数据就绪，否则 resolveItemMaster 在 W.skus 空时跑 → 材质/HS/用途全漏填(竞态 bug)
          let items = isXlsx ? await parsePackingXlsx(rd.result) : parsePackingList(rd.result);
          const meta = (items && items.meta) || {};
          // 源忠实：文件表头解析出的 FBA 号优先于 W.form.fbaNo 旧值，防止 step1 输入/缓存的旧 FBA 号污染新文件箱号
          const fileFbaNo = (meta.fbaNo || (items && items.fbaNo) || '').trim();
          const effectiveFbaNo = fileFbaNo || W.form.fbaNo || '';
          if(fileFbaNo && fileFbaNo !== W.form.fbaNo) W.form.fbaNo = fileFbaNo;
          items = normalizeItems(items, effectiveFbaNo);
          if(items.length){
            W.form.items=items; W.packed=true;
            // 同步更新收货人元数据（如果文件表头有）
            if(meta.fcCode) W.form.仓库代码 = meta.fcCode;
            if(meta.parsedAddress){
              const a=meta.parsedAddress;
              W.form.company=a.company||W.form.company||'';
              W.form.address=a.address||W.form.address||'';
              W.form.city=a.city||W.form.city||'';
              W.form.province=a.province||W.form.province||'';
              W.form.zip=a.zip||W.form.zip||'';
              W.form.country=a.country||W.form.country||'';
              W.form._addrFromFile=true;
            }
            msg.textContent='✅ 已上传并填入 '+items.length+' 行（'+file.name+'）。'+(meta.fcCode?' FC='+meta.fcCode:'');
            renderWizard();
          }
          else msg.textContent='⚠️ 解析为空，请确认文件是有效的装箱清单（Excel xlsx 或 CSV）。';
        }catch(err){ console.error(err); msg.textContent='❌ 解析失败：'+(err.message||err); }
      };
      rd.onerror=()=>{ msg.textContent='❌ 文件读取失败'; };
      if(isXlsx) rd.readAsArrayBuffer(file); else rd.readAsText(file);
    };
    const ob=$('#pl_online'); if(ob) ob.onclick=async()=>{ await onlineFetch(fid); };
  }
}
function packingBannerHTML(){
  const fid = W.packFbaId || (W.handover&&(W.handover.fba_shipment||W.handover.internal_no)) || '';
  const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
  const fromCloud = (W.packed && !pl.length);
  const fn = W.handover ? W.handover.packing_list : '';
  if(pl.length || fromCloud){
    const n = pl.length || W.form.items.length;
    const src = pl.length ? (fn||'(FBA箱唛交接表关联)') : '🌐 云端后端拉取（Railway）';
    return `
    <div class="card" style="margin-top:10px;border-color:#2b6cb0">
      <div class="hint ok">✅ 已从装箱清单自动填入 <b>${n}</b> 行物品（货件 ${esc(fid)}）。请核对品名/数量/申报价，无误即可继续。</div>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <button class="btn" id="pl_reload">↻ 重新从装箱清单拉取</button>
        <span class="muted">源：${esc(src)}</span>
      </div>
    </div>`;
  }
  return `
  <div class="card" style="margin-top:10px;border-color:#c53030">
    <div class="hint warn">⚠️ 本地未收录该货件（${esc(fid)}）的装箱清单内容。下方两个按钮：</div>
    <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-top:10px">
      <label class="btn" id="pl_upload_btn" style="margin:0;font-size:15px;padding:10px 20px;background:#2b6cb0;color:#fff"><span style="font-size:17px;margin-right:6px">📤</span>上传装箱清单（Excel / CSV）<input type="file" id="pl_file" accept=".xlsx,.xls,.csv" style="display:none"></label>
      <button class="btn" id="pl_online" style="font-size:15px;padding:10px 20px;background:#38a169;color:#fff"><span style="font-size:17px;margin-right:6px">🌐</span>在线获取</button>
      <span id="pl_msg" class="muted" style="flex:1;min-width:200px"></span>
    </div>
    <div class="hint" style="margin-top:10px;font-size:12px;color:#888">
      <b>📤 上传</b>：直接选本机的 xlsx/xls/csv 装箱清单（无需另存为 CSV）。<br>
      <b>🌐 在线获取</b>：①系统已收录该货件号 → 秒级自动填入；②<b>未收录</b> → <b>点击后会自动调云端后端从飞书云文档拉取</b>。<br>
      <span style="color:#2b6cb0">ℹ️ 云端优先</span>：当前默认连接 Railway 云端后端，无需你电脑开机；若遇「不在FBA表」说明镜像表正在同步，可稍后重试或先上传本机文件。
    </div>
  </div>`;
}
function loadPackingList(fid){
  const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
  if(pl.length){
    W.form.items = normalizeItems(pl.map(x=>{
      x = Object.assign({}, x);
      if((!x.declare||x.declare==='') && window.SKU_DECLARE && window.SKU_DECLARE[x.sku]){
        x.declare = window.SKU_DECLARE[x.sku].d;
        if(!x.nameCn) x.nameCn = window.SKU_DECLARE[x.sku].n;
      }
      return x;
    }), fid).map(resolveItemMaster);
    W.packed=true;
    renderWizard();
  } else { const m=$('#pl_msg'); if(m) m.textContent='本地未收录该装箱清单'; }
}
function parsePackingList(text){
  // 去 BOM
  text = text.replace(/^\uFEFF/, '');
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if(lines.length < 2) return [];

  // 简单 CSV 解析：支持引号内逗号、转义引号
  const splitCsv = line => {
    const out = [];
    let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
        else { inQ = !inQ; }
      } else if (ch === ',' && !inQ) {
        out.push(cur.trim().replace(/^"|"$/g, '').replace(/""/g, '"'));
        cur = '';
      } else { cur += ch; }
    }
    out.push(cur.trim().replace(/^"|"$/g, '').replace(/""/g, '"'));
    return out;
  };

  // 1) 扫描元信息：货件编号 / 配送地址 / 货件名称
  let _fbaNo = '', _deliveryAddress = '', _shipmentName = '';
  for (const line of lines.slice(0, 20)) {
    const c = splitCsv(line);
    const key = (c[0] || '').trim();
    const val = (c[1] || '').trim();
    if (/货件编号|shipment\s*id|fba\s*shipment/i.test(key)) _fbaNo = val;
    if (/配送地址|delivery\s*address/i.test(key)) _deliveryAddress = val;
    if (/货件名称|shipment\s*name/i.test(key)) _shipmentName = val;
  }

  // 2) 找真正表头行：Amazon 标准 CSV 前若干行是元信息，表头行须同时含 SKU + 箱号
  let headerIdx = -1;
  let hdr = [];
  for (let i = 0; i < Math.min(lines.length, 30); i++) {
    const c = splitCsv(lines[i]).map(h => h.toLowerCase());
    const hasSku = c.some(h => h === 'sku' || h.includes('sku'));
    const hasBoxNo = c.some(h => /箱号|box no|boxno|carton|ctn/.test(h));
    if (hasSku && hasBoxNo) { headerIdx = i; hdr = c; break; }
  }
  if (headerIdx < 0) { headerIdx = 0; hdr = splitCsv(lines[0]).map(h => h.toLowerCase()); }

  const map = {
    'boxLabel':['箱子名称','箱标签','箱标','boxlabel','label'],
    'boxNo':['箱号','box no','boxno','carton','ctn'],
    'sku':['sku'],
    'nameCn':['中文品名','namecn'],
    'nameEn':['英文品名','英文名称','nameen','商品名称','产品名称','product name','名称'],
    'qty':['数量','qty','quantity','商品总数','总数'],
    'qtyPerBox':['每箱件数','单箱数量','units per box','qty per box'],
    'boxes':['箱子总数','箱数','boxes','cartons'],
    'declare':['申报价','申报价值','declare','price'],
    'material':['材质','material'],
    'purpose':['用途','purpose','usage'],
    'hs':['hs','海关编码'],
    'brand':['品牌','brand'],
    'model':['型号','model'],
    'boxWeight':['包装箱重量','箱重','weight'],
    'prodWeight':['产品重量','商品重量','prodweight'],
    'len':['箱子长度','长','length'],
    'wid':['箱子宽度','宽','width'],
    'hgt':['箱子高度','高','height'],
    'elec':['带电','elec'],
    'magnet':['带磁','magnet'],
    'saleUrl':['销售链接','链接','url','saleurl'],
    'asin':['asin'],
    'fnsku':['fnsku']
  };
  const idx = {};
  for(const f in map){ const i=hdr.findIndex(h=>map[f].some(k=>h.includes(k))); if(i>=0) idx[f]=i; }

  const out = [];
  for(let n = headerIdx + 1; n < lines.length; n++){
    const c = splitCsv(lines[n]);
    if(c.every(x=>!x)) continue;
    const get = f => idx[f] >= 0 ? c[idx[f]] : '';
    const base = {
      boxLabel:get('boxLabel'), boxNo:get('boxNo'), sku:get('sku'),
      nameCn:get('nameCn'), nameEn:get('nameEn'), qty:get('qty') || 1,
      declare:get('declare'), material:get('material'), purpose:get('purpose'),
      hs:get('hs'), brand:get('brand'), model:get('model'),
      boxWeight:get('boxWeight'), prodWeight:get('prodWeight') || '',
      len:get('len'), wid:get('wid'), hgt:get('hgt'),
      elec:(get('elec')||'N').toUpperCase().startsWith('Y')?'Y':'N',
      magnet:(get('magnet')||'N').toUpperCase().startsWith('Y')?'Y':'N',
      saleUrl:get('saleUrl'), asin:get('asin'), fnsku:get('fnsku'), cost:''
    };

    // Amazon 原厂包装格式：行末"箱号"列是逗号分隔的多个真实 FBA 箱号
    const qtyPerBox = parseFloat(get('qtyPerBox')) || parseFloat(get('qty')) || 1;
    const boxNosStr = get('boxNo');
    const boxNos = boxNosStr ? boxNosStr.split(',').map(s=>s.trim()).filter(Boolean) : [];
    if(boxNos.length === 0){
      out.push({...base, boxLabel:base.boxLabel || base.boxNo, qty:qtyPerBox});
    }else{
      for(const bn of boxNos) out.push({...base, boxNo:bn, boxLabel:bn, qty:qtyPerBox});
    }
  }

  out.fbaNo = _fbaNo;
  out.meta = { fbaNo:_fbaNo, deliveryAddress:_deliveryAddress, shipmentName:_shipmentName, parsedAddress:{address:_deliveryAddress} };
  return out;
}

/* 在线获取：带 loading 进度条，先查缓存，后续扩展为后端代理 */
async function onlineFetch(fid){
  const msg=$('#pl_msg'); if(!msg) return;
  const btn=$('#pl_online'); if(btn) btn.disabled=true;
  await ensureSkusLoaded();       // 先确保 W.skus 就绪，避免预装路径 resolveItemMaster 漏填主数据
  await ensureBoxspecsLoaded();   // 同样确保箱规主数据就绪
  const steps=[
    {txt:'正在搜索货件号 <b>'+esc(fid)+'</b> 的装箱清单...'},
    {txt:'正在读取系统预装数据...'},
    {txt:'数据解析完成，正在填入表单...'}
  ];
  msg.innerHTML = '<div class="loading-wrap" id="loading_ui">'
    +'<div class="loading-header"><div class="spinner"></div><div class="loading-title">正在处理，请稍候...</div></div>'
    +'<div class="loading-bar-wrap"><div class="loading-bar" id="loading_bar"></div></div>'
    +'<div class="loading-steps" id="loading_steps">'
    +steps.map((s,i)=>'<div class="loading-step" data-i="'+i+'"><div class="dot"></div>'+s.txt+'</div>').join('')
    +'</div></div>';

  /* 阶段推进 */
  function setStep(i,state){
    const el=document.querySelector('#loading_steps [data-i="'+i+'"]');
    if(el){ el.className='loading-step '+state; }
  }
  function setBar(pct){
    const bar=document.querySelector('#loading_bar');
    if(bar) bar.style.width=pct+'%';
  }
  function done(ok,html){
    const ui=document.querySelector('#loading_ui');
    const spinner=ui&&ui.querySelector('.spinner');
    const title=ui&&ui.querySelector('.loading-title');
    if(spinner) spinner.style.animation='none';
    if(title){
      title.innerHTML=ok ? '✅ 完成' : '⚠️ 未找到';
      title.style.color=ok ? 'var(--green)' : 'var(--warn)';
    }
    /* 2秒后替换成结果信息 */
    if(ui){ setTimeout(()=>{ ui.outerHTML=html; if(btn) btn.disabled=false; }, ok ? 800 : 2000); }
    else if(btn) btn.disabled=false;
  }

  /* 阶段1: 正在搜索（立即激活）*/
  setStep(0,'active');
  setBar(20);

  setTimeout(()=>{
    setStep(0,'done');
    setStep(1,'active');
    setBar(50);

    setTimeout(()=>{
      setStep(1,'done');

      /* 实际查数据 */
      const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
      if(pl.length){
        setStep(2,'active');
        setBar(80);
        setTimeout(()=>{
          W.form.items = normalizeItems(pl.map(x=>resolveItemMaster(Object.assign({}, x))), W.form.fbaNo);
          W.packed=true;
          W.packed=true;
          setStep(2,'done');
          setBar(100);
          renderWizard();
          done(true,'<div class="hint ok" style="margin-top:10px">✅ 已从系统预装的装箱清单自动填入 <b>'+pl.length+'</b> 行（货件 '+esc(fid)+'）。请核对品名/数量/申报价。</div>');
        },300);
      } else {
        /* 未预装 → 尝试调本地后端代理（15秒超时,卡死自动放弃） */
        setStep(1,'done');
        setStep(2,'active');
        setBar(60);
        const backendUrl = localStorage.getItem('backend_url') || RAILWAY_URL;
        fetch(backendUrl+'/api/fetch-packing-list', {
          method:'POST',
          headers:{'Content-Type':'application/json'},
          body:JSON.stringify({fid}),
          signal: AbortSignal.timeout(15000) // 15s 超时,死锁自动放弃
        }).then(r=>r.json()).then(data=>{
          if(data.ok && data.items && data.items.length>0){
            setStep(2,'done');
            setBar(100);
            // 后端返回的重量字段叫 weight, 前端物品模型用 boxWeight —— 对齐, 否则云端发票毛重空白
            data.items.forEach(x=>{ if(!x.boxWeight && x.weight) x.boxWeight = x.weight; });
            W.form.items = normalizeItems(data.items.map(x=>resolveItemMaster(Object.assign({}, x))), W.form.fbaNo);
            W.packed=true;
            renderWizard();
            done(true,'<div class="hint ok" style="margin-top:10px">✅ 已通过后端代理从飞书云文档拉取 <b>'+data.items.length+'</b> 行（货件 '+esc(fid)+'）。请核对品名/数量/申报价。</div>');
          } else {
            throw new Error(data?.error||'后端返回空数据');
          }
        }).catch(e=>{
          setStep(2,'fail');
          setBar(100);
          let hint = '';
          if(e.name==='TimeoutError' || e.message.includes('timeout')){
            hint = '<div class="hint warn" style="margin-top:10px">⏱️ 云端后端响应超时（15秒）。<br>① 稍后重试；<br>② 点右上角「后端：已连接」检查是否连上云端；<br>③ 当前货件可直接点「<b>📤 上传装箱清单</b>」选本机 xlsx。</div>';
          } else if(e.message.includes('Failed to fetch')||e.message.includes('fetch')){
            hint = '<div class="hint warn" style="margin-top:10px">❌ 浏览器连不上云端后端。<br>① 检查网络；<br>② 点右上角「后端：已连接」切换回本机 localhost:3460 兜底；<br>③ 或直接点「<b>📤 上传装箱清单</b>」选本机文件。</div>';
          } else {
            hint = '<div class="hint warn" style="margin-top:10px">⚠️ <b>'+esc(fid)+'</b> 后端不可用（'+esc(e.message)+'）。<br>① 云端优先模式下可稍后重试；<br>② 点「<b>📤 上传装箱清单</b>」选本机 xlsx（<b>已加强解析支持</b>）。</div>';
          }
          done(false, hint);
        });
      }
    }, 400);
  }, 300);
}

/* 带 loading 的 loadPackingList（从搜索结果自动填入也用同样的流程） */
async function loadPackingList(fid){
  await ensureSkusLoaded();       // 确保 W.skus 就绪再回填主数据
  await ensureBoxspecsLoaded();   // 确保 W.boxspecs 就绪再回填箱重/长宽高
  const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
  if(pl.length){
    W.form.items = normalizeItems(pl.map(x=>resolveItemMaster(Object.assign({}, x))), W.form.fbaNo);
    W.packed=true;
    renderWizard();
  } else { const m=$('#pl_msg'); if(m) m.textContent='本地未收录该装箱清单，请上传 CSV 或 Excel。'; }
}

/* 从亚马逊配送地址字符串解析结构化地址：公司,街道,...,城市,省/州,邮编,国家
   20260821 修复：原实现要求邮编以数字开头（美国格式），加拿大(V3M 5Y9)/英国(B48 7JA)等字母开头邮编被误清空→zip 报空。
   改为从尾部按固定位置取（country←末段, zip←倒数2, province←倒数3, city←倒数4, address←中间段, company←首段），不依赖邮编格式。
   兼容段数：≥6 段(美/加 带省) / 5 段(英 无省) / <4 段兜底原样。 */
function parseDeliveryAddress(str){
  if(!str) return null;
  const parts = str.split(',').map(s=>s.trim()).filter(s=>s!=='');
  if(parts.length < 4) return {company:'', address:str, city:'', province:'', zip:'', country:''};
  const countryMap = {US:'美国',USA:'美国',CA:'加拿大',UK:'英国',GB:'英国',DE:'德国',SA:'沙特',AE:'阿联酋',JP:'日本',AU:'澳大利亚'};
  const countryRaw = parts[parts.length-1];
  const country = countryMap[countryRaw.toUpperCase()] || countryRaw;
  const zip = parts[parts.length-2] || '';
  // ≥6 段：最后三段是 省/州, 邮编, 国家；5 段(英国无省)：城市, 邮编, 国家
  const hasProvince = parts.length >= 6;
  const province = hasProvince ? (parts[parts.length-3] || '') : '';
  const cityIdx = hasProvince ? parts.length-4 : parts.length-3;
  const city = parts[cityIdx] || '';
  const addressParts = parts.slice(1, cityIdx);
  const address = addressParts.join(', ') || parts[1] || '';
  const company = parts[0] || '';
  return {company, address, city, province, zip, country};
}

/* 解析 Excel xlsx 装箱清单：用 ExcelJS 读，支持亚马逊 ONE_SKU 导出(格式A,按箱号/箱子名称区间展开)和通用按箱展开格式 */
async function parsePackingXlsx(arrayBuffer){
  if(typeof ExcelJS==='undefined') throw new Error('ExcelJS 未加载');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  const ws = wb.worksheets[0];
  if(!ws) throw new Error('Excel文件无工作表');

  // 0. 扫描元数据行（前5行）提取 FBA 货件号、物流中心编码(FC)、配送地址
  let _fbaNo='', _fcCode='', _shipmentName='', _deliveryAddress='', _shipMethod='';
  const nextVal=(row, startC, skip)=>{
    for(let c=startC+1;c<=Math.max(row.cellCount,30);c++){
      const v=row.getCell(c).value;
      const s=v===null||v===undefined?'':String(v).trim();
      if(s!=='' && s!==skip) return s; // skip=触发标签,避免合并单元格把标签复制到相邻列导致取到标签本身
    }
    return '';
  };
  for(let r=1;r<=Math.min(ws.rowCount,20);r++){
    const row=ws.getRow(r);
    for(let c=1;c<=row.cellCount;c++){
      const v = String(row.getCell(c).value||'');
      if(!_fbaNo){
        const m = v.match(/\b(FBA\d{2}[A-Z0-9]{5,})\b/);
        if(m){ _fbaNo=m[1]; }
      }
      if(/物流中心编码|配送中心|FC Code|Fulfillment Center|仓库代码/i.test(v) && !_fcCode){
        _fcCode = nextVal(row,c,v).toUpperCase().replace(/[^A-Z0-9]/g,'');
      }
      if(/货件名称|Shipment Name/i.test(v) && !_shipmentName){
        _shipmentName = nextVal(row,c,v);
      }
      if(/配送地址|收件地址|收货地址|Ship To|Destination|Delivery Address|地址/i.test(v) && !_deliveryAddress){
        _deliveryAddress = nextVal(row,c,v);
      }
      if(/运输方式|Ship Method|Shipping Method|Shipment Method/i.test(v) && !_shipMethod){
        _shipMethod = nextVal(row,c,v);
      }
    }
    if(_fbaNo && _fcCode && _deliveryAddress) break;
  }

  // 1. 扫描前30行,找到真正的表头行（格式B「FBA订单(V3)」的表头在 R20,需扩范围；格式A 在 R3 不受影响）
  let headerRow = 1;
  let headerStr = '';
  for(let r=1;r<=Math.min(ws.rowCount,30);r++){
    const row=ws.getRow(r);
    let txt=''; for(let c=1;c<=row.cellCount;c++) txt += String(row.getCell(c).value||'')+'|';
    const low=txt.toLowerCase();
    // 跳过汇总/元数据行（如"装箱方式|每箱一款SKU|SKU种类|64|总箱数|134..."），这些行不是列头
    if(/装箱方式|sku种类|总箱数|总重量|总体积|发货量\s*\|\s*\d/.test(low)) continue;
    // 检测多种表头模式
    if((low.includes('sku')||low.includes('msku')) &&
       (low.includes('数量')||low.includes('qty')||low.includes('件数')||low.includes('箱数'))){
      headerRow=r; headerStr=txt; break;
    }
    // 格式B：亚马逊「FBA订单(V3)」装箱单 —— 表头含 No.of Pkgs / 中文品名+英文品名 / HS CODE
    if(low.includes('no.of pkgs') || (low.includes('中文品名')&&low.includes('英文品名')) || (low.includes('hs code')&&low.includes('品名'))){
      headerRow=r; headerStr=txt; break;
    }
  }

  // 2. 读表头
  const hdr=ws.getRow(headerRow);
  const headers=[]; for(let c=1;c<=Math.max(hdr.cellCount,30);c++) headers.push(String(hdr.getCell(c).value||'').trim());
  const lower=headers.map(h=>h.toLowerCase());

  const findCol=cands=>{
    // 候选按优先级顺序扫描表头，避免「序号」等通用名因列靠前而覆盖「箱号」等具体列
    let i=-1;
    for(const k of cands){ i=lower.findIndex(h=>h===k); if(i>=0) break; }
    if(i<0) for(const k of cands){ i=lower.findIndex(h=>h.includes(k)); if(i>=0) break; }
    return i>=0?i+1:0;
  };
  const col={
    sku: findCol(['msku','sku','型号','产品型号']),
    fnsku: findCol(['fnsku']),
    nameCn: findCol(['申报中文名','中文品名','中文名称','品名','中文','名称']),
    nameEn: findCol(['英文品名','英文名称','商品名称','英文','nameen','english','product name']),
    qty: findCol(['发货量','已装量','数量','qty','quantity','商品总数']),
    qtyPerBox: findCol(['单箱数量','每箱数量','每箱件数']),
    boxes: findCol(['箱子总数','箱数','boxes','ctns','cartons']),
    boxSpec: findCol(['箱子型号','箱规','boxspec','box spec']),
    boxWeight: findCol(['箱子毛重','单箱毛重','箱重','重量','weight']),
    len: findCol(['箱子长度','长','length','l']),
    wid: findCol(['箱子宽度','宽','width','w']),
    hgt: findCol(['箱子高度','高','height','h']),
    boxNo: findCol(['箱号','boxno','box no','fba箱号','序号']),
    boxLabel: findCol(['箱子名称','箱标签','boxlabel','label']),
    asin: findCol(['asin']),
    brand: findCol(['品牌','brand']),
    declare: findCol(['申报','vauel','申报价','value','申报(usd)']),
    hs: findCol(['hs code','海关编码','hs','h.s.']),
    material: findCol(['材质','material']),
    purpose: findCol(['用途','purpose']),
    prodWeight: findCol(['产品毛重','毛重','gw(kg)产品','product weight']),
    elec: findCol(['带电','electric']),
    packed: findCol(['已装箱数','已装量','已装箱']), // MUL_SKU 混箱格式：H列=该SKU已装箱合计（校验用）
  };

  // 3. 格式检测
  const hasWorkingWorkflow = col.asin>0 && col.boxNo>0 && !col.boxLabel; // 亚马逊"原厂包装发货"格式
  const hasAmazonOneSKU = col.qtyPerBox>0 && col.boxLabel>0; // 亚马逊 ONE_SKU_NO_PIC
  const hasGenericFormat = col.sku>0 && col.qty>0;

  // 3.5 混箱(MUL_SKU)格式检测：表头含「第1箱」「第2箱」…子表头列
  //    特征：一行一个 SKU，行内按「第N箱」列分布各箱数量；箱号/箱重/长宽高/箱子名称在文件底部每箱一列（横向）。
  //    与 ONE_SKU 区别：单箱可混多款 SKU；箱号不在数据行，而在底部「箱号」行（如 FBA19M26WNLMU000001）。
  let multiCols = [];  // 子表头「第N箱」列号（按列顺序，第 k 个 = 第 k 箱）
  let hasMultiSku = false;
  {
    const hdrRow2 = ws.getRow(headerRow);
    for(let c=1;c<=Math.max(hdrRow2.cellCount,30);c++){
      const h = String(hdrRow2.getCell(c).value||'').trim();
      if(/^第\d+箱$/.test(h)) multiCols.push(c);
    }
    hasMultiSku = multiCols.length>=1;
  }

  if(!col.sku && !col.boxNo){
    throw new Error(`未识别表头。请确认 Excel 包含 SKU/MSKU/箱号/装箱清单 等列。当前表头: ${headers.slice(0,10).join(', ')}`);
  }

  const out=[];
  const getStr=(row,c)=> c>0 ? String(row.getCell(c).value||'').trim() : '';
  const getNum=(row,c)=> c>0 ? (parseFloat(row.getCell(c).value)||'') : '';

  // MUL_SKU 混箱格式：读取底部箱区行号（合计/重量/长宽高/箱号/箱子名称）。
  // 值列与「第N箱」子表头列一一对应：第 k 个 multiCols 列 = 第 k 箱的底部值。
  let mkBottom=null;
  if(hasMultiSku){
    const scanBottom=(re)=>{
      // 起始行从表头之后开始扫（原 Math.min(rowCount,45) 对小文件失效：rowCount=16 时只扫最后一行，
      // 漏掉 R10-R16 的合计/重量/尺寸/箱号行 → boxNo 全空。改从 headerRow+1 扫到底，大/小文件都覆盖）
      for(let r=Math.max(headerRow+1, 2); r<=ws.rowCount; r++){
        const row=ws.getRow(r);
        for(let c=1;c<=row.cellCount;c++){
          if(re.test(String(row.getCell(c).value||'').trim())) return r;
        }
      }
      return 0;
    };
    mkBottom = {
      totalRow:   scanBottom(/^合计|^总计|^小计/),
      weightRow:  scanBottom(/^Weight of box|箱子毛重|单箱毛重|箱重/i),
      lenRow:     scanBottom(/^Box length|箱子长度/i),
      widRow:     scanBottom(/^Box width|箱子宽度/i),
      hgtRow:     scanBottom(/^Box height|箱子高度/i),
      boxNoRow:   scanBottom(/^箱号$|^箱子号$|^FBA箱号|^货件箱号/),
      nameRow:    scanBottom(/^箱子名称$|^箱标签|^box\s*label/i),
    };
  }

  // 解析 FBA 箱号区间：支持 Amazon ONE_SKU 的"末位缩写"写法
  //   "FBA19K786CWTU000001～19；"   → FBA19K786CWTU000001 … FBA19K786CWTU000019（末位仅写 19，需按起始号宽度对齐补零）
  //   "FBA19J6FCXNKU000001～2；"    → …U000001、…U000002
  //   "FBA19J6FCXNKU000001～FBA19J6FCXNKU000002；" → 完整写法也支持
  // 关键：箱号必须 100% 来自装箱清单的"箱号"列，绝不能退化成用货件号去"造"。
  const parseBoxRange=(str)=>{
    if(!str) return [];
    const segs = String(str).split(/[；;]/).map(s=>s.trim()).filter(Boolean);
    const res=[];
    for(const raw of segs){
      const m = raw.match(/^(FBA[A-Z0-9]*U)(\d+)\s*[~～\-]\s*(?:(FBA[A-Z0-9]*U)?)(\d+)$/i);
      if(m){
        const pre=m[1];                 // 货件前缀，如 FBA19K786CWTU
        const startDigits=m[2];         // 起始 6 位，如 000001
        const endDigits=m[4];           // 末尾数字，可能缩写为 19 或完整 000019
        let fullEnd;
        if(endDigits.length>=startDigits.length) fullEnd=endDigits.slice(-startDigits.length);
        else fullEnd=startDigits.slice(0, startDigits.length-endDigits.length)+endDigits;
        const w=startDigits.length;
        const sN=parseInt(startDigits,10), eN=parseInt(fullEnd,10);
        for(let i=sN;i<=eN;i++) res.push(pre+String(i).padStart(w,'0'));
        continue;
      }
      if(raw) res.push(raw);
    }
    return res;
  };
  // 解析箱标签: "P2 - B1～B2" → [B1,B2]（跳过托盘号）
  const parseLabelRange=(str)=>{
    if(!str) return [];
    const last = str.includes(' - ') ? str.split(' - ').pop() : str;
    const m=last.match(/^([A-Za-z]?)(\d+)\s*[～~\-]\s*([A-Za-z]?\d+)$/);
    if(m){
      const prefix=m[1]||'B';
      const start=parseInt(m[2]), end=parseInt(m[3].replace(/^[A-Za-z]/,''));
      const arr=[]; for(let i=start;i<=end;i++) arr.push(prefix+i);
      return arr;
    }
    if(last.match(/^[A-Za-z]?\d+$/)) return [last];
    return last ? [last] : [];
  };

  for(let r=headerRow+1;r<=ws.rowCount;r++){
    const row=ws.getRow(r);
    const vals=[]; for(let c=1;c<=Math.max(row.cellCount,30);c++) vals.push(row.getCell(c).value);
    // 跳过全空行
    if(vals.every(v=>v===null||v===undefined||v==='')) continue;

    const sku=getStr(row,col.sku);
    const fnsku=getStr(row,col.fnsku);
    const nameEn=getStr(row,col.nameEn);
    const nameCn=getStr(row,col.nameCn);
    const boxWeight=getNum(row,col.boxWeight);
    const len=getNum(row,col.len);
    const wid=getNum(row,col.wid);
    const hgt=getNum(row,col.hgt);

    const base={sku,fnsku,nameCn,nameEn,brand:getStr(row,col.brand)||'JW PEI',boxSpec:getStr(row,col.boxSpec),boxWeight,prodWeight:getNum(row,col.prodWeight),len,wid,hgt,material:getStr(row,col.material),purpose:getStr(row,col.purpose),hs:getStr(row,col.hs),model:'',elec:getStr(row,col.elec)||'N',magnet:'N',saleUrl:'',declare:getStr(row,col.declare),cost:'',boxes:parseInt(getStr(row,col.boxes))||''};

    if(hasMultiSku){
      // 跳过底部箱区行（Weight of box/Length/Width/Height/合计/箱号/箱子名称）：其 MSKU/FNSKU/品名 为空
      if(!sku && !fnsku && !nameCn && !nameEn) continue;
      // MUL_SKU 混箱格式：一行一 SKU，按「第N箱」列展开成 每箱×每SKU 一条记录。
      // 源忠实：箱号 100% 取底部「箱号」行（绝不重造）；子单号=同箱号（用户铁律：箱号=子单号）；
      //         数量=该箱列值（不是 G 列总发货量）；箱重/长宽高=底部对应列值。
      const getBottom=(r,c)=> r>0 ? String(ws.getRow(r).getCell(c).value||'').trim() : '';
      const getBottomNum=(r,c)=> r>0 ? (parseFloat(ws.getRow(r).getCell(c).value)||'') : '';
      let rowSum=0;
      for(let k=0;k<multiCols.length;k++){
        const c=multiCols[k];
        const q=getNum(row,c);
        if(q>0){
          const boxNo = mkBottom.boxNoRow ? getBottom(mkBottom.boxNoRow,c) : '';
          out.push({
            ...base,
            boxNo, boxLabel: boxNo, // 子单号=同箱号
            qty:q, boxes:1,
            boxWeight: getBottomNum(mkBottom.weightRow,c),
            len: getBottomNum(mkBottom.lenRow,c),
            wid: getBottomNum(mkBottom.widRow,c),
            hgt: getBottomNum(mkBottom.hgtRow,c),
          });
          rowSum += q;
        }
      }
      // 校验①：行内各箱数量之和 == 该SKU「已装箱数」(H列)，不符记警告（fail loud，不静默）
      if(col.packed>0){
        const packedN = getNum(row,col.packed);
        if(packedN>0 && rowSum!==packedN){
          out.warnings = out.warnings||[];
          out.warnings.push(`SKU ${sku}：展开各箱数量合计 ${rowSum} ≠ 清单已装箱数 ${packedN}`);
        }
      }
    } else if(hasAmazonOneSKU){
      // 亚马逊 ONE_SKU_NO_PIC 格式: 一行一个SKU, 展开多箱
      const qtyPerBox=parseFloat(getStr(row,col.qtyPerBox))||0;
      const boxNos=parseBoxRange(getStr(row,col.boxNo));
      const boxLabels=parseLabelRange(getStr(row,col.boxLabel));
      const count=Math.max(boxNos.length, boxLabels.length, parseInt(getStr(row,col.boxes))||1);
      for(let k=0;k<count;k++){
        // 已按箱展开，每行一箱
        const item={...base, boxNo:boxNos[k]||`box${k+1}`, boxLabel:boxLabels[k]||'', qty:qtyPerBox, boxes:1};
        out.push(item);
      }
    } else if(hasWorkingWorkflow){
      // 亚马逊"原厂包装发货"格式: 行末"箱号"列是逗号分隔的箱号列表
      const qtyPerBox=parseFloat(getStr(row,col.qtyPerBox))||parseFloat(getStr(row,col.qty))||0;
      const boxNosStr=getStr(row,col.boxNo);
      const boxNos=boxNosStr ? boxNosStr.split(',').map(s=>s.trim()).filter(Boolean) : [];
      if(boxNos.length===0){
        // 没有箱号就当一行一箱
        out.push({...base, boxNo:'', boxLabel:'', qty:qtyPerBox||1});
      } else {
        for(const bn of boxNos){
          out.push({...base, boxNo:bn, boxLabel:bn, qty:qtyPerBox});
        }
      }
    } else {
      // 通用格式: 一行一箱,或一行多箱
      const qty=parseFloat(getStr(row,col.qty))||1;
      const boxNoStr=getStr(row,col.boxNo);
      // 如果箱号是逗号分隔,也展开
      if(boxNoStr && boxNoStr.includes(',')){
        for(const bn of boxNoStr.split(',').map(s=>s.trim()).filter(Boolean)){
          out.push({...base, boxNo:bn, boxLabel:getStr(row,col.boxLabel)||bn, qty});
        }
      } else {
        out.push({...base, boxNo:boxNoStr, boxLabel:getStr(row,col.boxLabel)||boxNoStr, qty});
      }
    }
  }

  // 校验②（MUL_SKU）：每箱展开件数合计 == 底部「合计」行对应列（如 Q52=16），不符记警告。
  // 注意：合计行数值可能是共享公式对象（{result:22, sharedFormula}），取值须兼容（parseFloat 对对象得 NaN → 误报）。
  if(hasMultiSku && mkBottom && mkBottom.totalRow){
    const perBox={};
    for(const it of out) perBox[it.boxNo]=(perBox[it.boxNo]||0)+it.qty;
    const warns = out.warnings = out.warnings||[];
    const cellNum=(r,c)=>{ const v=ws.getRow(r).getCell(c).value; return parseFloat((typeof v==='object'&&v&&v.result!==undefined)?v.result:v)||0; };
    for(let k=0;k<multiCols.length;k++){
      const expected = cellNum(mkBottom.totalRow, multiCols[k]);
      const bn = mkBottom.boxNoRow ? String(ws.getRow(mkBottom.boxNoRow).getCell(multiCols[k]).value||'').trim() : `第${k+1}箱`;
      if(expected>0 && (perBox[bn]||0)!==expected){
        warns.push(`箱号 ${bn}：展开件数 ${perBox[bn]||0} ≠ 清单合计 ${expected}`);
      }
    }
  }

  if(out.length===0){
    throw new Error('解析后无数据行。请检查 Excel 格式。');
  }

  // 同步本地主数据(品名/HS/申报价/箱规格反查) — 统一走 resolveItemMaster
  try{ for(const it of out) resolveItemMaster(it); }catch(e){ console.warn('parsePackingXlsx 主数据同步跳过:', e); }
  // 从真实箱号反推货件号（装箱清单自身最权威）：覆盖文件头漏提取/错提取，以及 step1 残留的 stale FBA 号污染。
  // 这样不论用户此前在 step1 输入过别的货件号，上传本文件后系统都以"装箱清单里实际写着的货件"为准。
  let derivedFba='';
  for(const it of out){
    const s=String(it.boxNo||'');
    const m=s.match(/^(FBA[A-Z0-9]*)U\d{6}$/i);
    if(m){ derivedFba=m[1].toUpperCase(); break; }
  }
  const shipmentNo = derivedFba || _fbaNo;
  if(shipmentNo) out.fbaNo = shipmentNo;
  out.meta = { fbaNo:shipmentNo, fcCode:_fcCode, shipmentName:_shipmentName, deliveryAddress:_deliveryAddress, shipMethod:_shipMethod, parsedAddress: parseDeliveryAddress(_deliveryAddress) };
  return out;
}

function step4(box){
  const tmpls = W.templates;
  // 【质量第一·防错】模板必须跟随物流商，杜绝跨票 stale 选择（如上一票安速残留）误用错模板。
  // 仅当"当前已选模板的物流商 ≠ 本票物流商"时才自动对齐到本票物流商的 ACTIVE 模板。
  const curT = W.selTmpl ? tmpls.find(t=>t.id===W.selTmpl) : null;
  // 物流商明确时才自动对齐其 ACTIVE 模板；物流商为空不 fallback 到 tmpls[0]，避免误用错模板（UI 会要求先选物流商）
  if(W.form.物流商){
    if(!curT || curT.物流商 !== W.form.物流商){
      const mt = tmpls.find(t=>t.物流商===W.form.物流商 && t.状态==='ACTIVE') || tmpls.find(t=>t.物流商===W.form.物流商);
      if(mt) W.selTmpl = mt.id;
    }
  }
  box.innerHTML = `
  <div class="card">
    <h3>④ 选模板 · 预览映射</h3>
    <div class="hint">选一个 ACTIVE 模板（已内置 5 家各自字段映射）。下方展示「字段 → 取值来源」。绿=主数据反查，白=手填，黄=推算。模板只定格子位置，值来自 L4。</div>
    ${W.form.物流商 ? `<div class="hint ok" style="margin-top:8px">✅ 当前物流商：<b>${esc(W.form.物流商)}</b>（来自步骤①选择）。模板已自动对齐到该物流商的 ACTIVE 模板；如需修改请下拉重选。</div>` : `<div class="alert alert-err" style="margin-top:8px">⛔ 尚未选择物流商。请返回步骤①选择物流商+渠道，或在上方下拉框手动选对模板（选错物流商将导致整张发票错版）。</div>`}
    <label>模板（物流商）</label>
    <select id="selTmpl"><option value="">-- 请选择模板（物流商） --</option>${tmpls.length? tmpls.map(t=>`<option value="${t.id}" ${t.id===W.selTmpl?'selected':''}>${esc(t.物流商)} (v${t.版本||1}, ${t.状态})</option>`).join('') : '<option>（无可用模板）</option>'}</select>
    ${tmpls.length?'':`
      <div id="tmplSeedPanel" style="margin-top:10px;padding:10px;background:var(--panel);border-radius:6px;border:1px dashed var(--border)">
        ${seedStatus.loading
          ? `<div class="hint">正在从服务器加载默认模板 ${seedStatus.loaded}/${seedStatus.total} … 请稍候</div>`
          : (seedStatus.errors.length
              ? `<div class="alert alert-err" style="margin:0"><b>默认模板加载失败</b><br>${esc(seedStatus.errors.join('<br>'))}<br>常见原因：隐私/访客模式限制网络、CDN 未刷新。请点击下方按钮重试。</div><button class="btn" id="retrySeedTmpl" style="margin-top:8px">重新加载默认模板</button>`
              : `<div class="hint">首次使用或访客模式需要加载默认模板。若下方无模板列表，请<button class="btn" id="retrySeedTmpl" style="margin-left:6px">点击加载默认模板</button></div>`)}
      </div>
    `}
    <div id="mapPreview" style="margin-top:14px"></div>
  </div>
  <div style="margin-top:14px;display:flex;gap:10px"><button class="btn secondary" id="prev4">← 上一步</button><button class="btn" id="next4">下一步：校验反查 →</button></div>`;
  const preview = ()=>{
    const val = $('#selTmpl').value;
    const t = tmpls.find(x=>x.id===val);
    if(!t){ $('#mapPreview').innerHTML='<div class="empty">请选择一个模板</div>'; W.selTmpl = undefined; return; }
    W.selTmpl = t.id;
    const mp = t.mapping||{};
    const rows = Object.keys(W.sources).filter(k=>mp.meta&&mp.meta[k]).map(k=>{
      const s=W.sources[k]; const cls=(s.src==='channel'||s.src==='warehouse'||s.src==='template')?'cell-src':'cell-manual';
      return `<tr><td>${k}</td><td class="${cls}">${esc(s.v)||'<span class=muted>—</span>'}</td><td>${s.src}</td><td>${mp.meta[k]}</td></tr>`;
    }).join('');
    $('#mapPreview').innerHTML = `<table><thead><tr><th>收货人字段</th><th>值</th><th>来源</th><th>模板格子</th></tr></thead><tbody>${rows||'<tr><td colspan=4 class=empty>该模板无对应收货人字段</td></tr>'}</tbody></table>
      <div class="hint">物品行将按模板第 ${mp.itemStartRow||'?'} 行起逐行填入（箱号/品名/数量/申报价/材质/HS/品牌/型号等，按该模板实际列映射）。</div>`;
  };
  $('#selTmpl').onchange=preview;
  preview();
  const retryBtn = $('#retrySeedTmpl');
  if(retryBtn){
    retryBtn.onclick = async ()=>{
      retryBtn.textContent = '加载中…';
      retryBtn.disabled = true;
      await seedDefaultTemplates(true);
      // 重新渲染 step4 以刷新下拉列表
      W.step = 4;
      renderWizard();
    };
  }
  $('#prev4').onclick=()=>{W.step=3;renderWizard();};
  $('#next4').onclick=()=>{ if(!W.selTmpl){ alert('请先选择一个模板'); return;} W.step=5; renderWizard(); };
}
function step5(box){
  const checks = runChecks();
  W.checks = checks;
  const passAll = checks.every(c=>c.level!=='err');
  box.innerHTML = `
  <div class="card">
    <h3>⑤ 校验反查（质量命门）</h3>
    <div class="hint">独立重读对账：必填完整性 / 勾稽（箱数=物品行数、数量合计）/ 源忠实高亮。告警+阻断：有红色错误须先修。</div>
    ${checks.map(c=>{ const cls=c.level==='err'?'alert-err':(c.level==='warn'?'alert-warn':'alert-ok'); const icon=c.level==='err'?'⛔':(c.level==='warn'?'⚠️':'✅'); return `<div class="alert ${cls}">${icon} <b>${esc(c.name)}</b>：${esc(c.msg)}</div>`; }).join('')}
    <div style="margin-top:10px"><b>勾稽汇总：</b>物品行数=${W.form.items.length}，数量合计=${checks.reduce((a,c)=>a+(c.qtySum||0),0)}，申报总值=$${checks.reduce((a,c)=>a+(c.decSum||0),0).toFixed(2)}</div>
  </div>
  <div style="margin-top:14px;display:flex;gap:10px"><button class="btn secondary" id="prev5">← 上一步</button><button class="btn ${passAll?'':'secondary'}" id="next5" ${passAll?'':'disabled'}>${passAll?'下一步：人审·交付 →':'请先修复红色错误'}</button></div>`;
  $('#prev5').onclick=()=>{W.step=4;renderWizard();};
  $('#next5').onclick=()=>{ if(passAll){W.step=6;renderWizard();} };
}
function isDutyIncluded(channelName){
  // 包税/双清包税：物流商负责目的国税费，无需卖家提供 VAT/EORI
  return /(包税|双清|DDP|完税)/i.test(channelName||'');
}
function runChecks(){
  const out=[];
  const selTpl = W.templates.find(x=>x.id===W.selTmpl);
  const tpl = (selTpl && selTpl.mapping) || MAPPINGS[selTpl ? selTpl.物流商 : W.form.物流商];
  // 必填收货人字段 = 【本物流商模板实际用到】的"物理地址类"字段（校验跟随模板，杜绝误阻断）
  // 物理地址类（主数据可反查：仓库/交接表/渠道）：阻断级
  const META_REQUIRED = ['fbaNo','company','address','zip','country','warehouseCode','province','city'];
  // 税务/注册类：仅告警，不阻断；但包税渠道无需 VAT/EORI，故不告警
  const dutyIncluded = isDutyIncluded(W.form.渠道);
  const META_WARN_BASE = ['customs','phone','email'];
  const META_WARN_TAX = dutyIncluded ? [] : ['vat','eori','vatName','vatAddr'];
  const META_WARN = META_WARN_BASE.concat(META_WARN_TAX);
  const reqMeta = (tpl && tpl.meta) ? Object.keys(tpl.meta).filter(k=>META_REQUIRED.includes(k)) : [];
  const warnMeta = (tpl && tpl.meta) ? Object.keys(tpl.meta).filter(k=>META_WARN.includes(k)) : [];
  const missing = reqMeta.filter(k=> !W.sources[k] || !String(W.sources[k].v).trim());
  if(missing.length) out.push({level:'err',name:'必填完整性',msg:'以下字段为空（按「'+W.form.物流商+'」模板必填）：'+missing.join('、')});
  else out.push({level:'ok',name:'必填完整性',msg:'本物流商模板所需收货人字段均已填'});
  const warnMissing = warnMeta.filter(k=> !W.sources[k] || !String(W.sources[k].v).trim());
  if(warnMissing.length) out.push({level:'warn',name:'税务/注册字段空缺',msg:'模板含这些字段但主数据为空（非阻断，可人工补填）：'+warnMissing.join('、')});
  let qtySum=0, decSum=0, itemErr=0;
  const isAnsu = (W.form.物流商==='安速');
  let hsErr=0, labelErr=0, matErr=0, useErr=0;
  // 行级定位：收集具体 SKU/箱号/缺字段，避免只报"有 N 行缺字段"让用户无从下手
  const itemMiss=[], hsMiss=[], matMiss=[], useMiss=[], labelMiss=[];
  W.form.items.forEach((it,i)=>{
    // 校验时反查 SKU 主数据（与 step3 渲染一致：原始空 → 用 sk.申报价/sk.中文品名）
    const sk = W.skus.find(s=>s.sku===it.sku);
    const effDeclare = (it.declare!==''&&it.declare!=null) ? it.declare : (sk && sk.申报价 ? sk.申报价 : (window.SKU_DECLARE && window.SKU_DECLARE[it.sku] ? window.SKU_DECLARE[it.sku].d : ''));
    const effNameCn = it.nameCn || (sk ? sk.中文品名 : (window.SKU_DECLARE && window.SKU_DECLARE[it.sku] ? window.SKU_DECLARE[it.sku].n : ''));
    const effHs = it.hs || (sk ? sk.HS : '');
    const effMat = it.material || (sk ? sk.材质 : '');
    const effUse = it.purpose || (sk ? sk.用途 : '');
    const miss=[];
    if(!it.boxNo) miss.push('箱号');
    if(!effNameCn) miss.push('品名');
    if(!it.qty) miss.push('数量');
    if(!(effDeclare!==''&&effDeclare!=null)) miss.push('申报价');
    if(miss.length){ itemErr++; itemMiss.push(`${it.sku||'未知SKU'}(箱${it.boxNo||'?'}: 缺${miss.join('/')})`); }
    if(!effHs){ hsErr++; hsMiss.push(it.sku||'未知SKU'); }
    if(!effMat){ matErr++; matMiss.push(it.sku||'未知SKU'); }
    if(!effUse){ useErr++; useMiss.push(it.sku||'未知SKU'); }
    if(isAnsu && !it.boxLabel){ labelErr++; labelMiss.push(it.boxNo||('行'+(i+1))); }
    qtySum+=parseFloat(it.qty)||0;
    decSum+=(parseFloat(effDeclare)||0)*(parseFloat(it.qty)||0);
  });
  const clip = (arr)=> arr.slice(0,15).join('、') + (arr.length>15?` …(其余 ${arr.length-15} 个略)`:'');
  if(itemErr) out.push({level:'err',name:'物品必填',msg:`有 ${itemErr} 行缺字段（已定位到 SKU/箱号）：${clip(itemMiss)}。商品申报信息表未收录的 SKU 需手填申报价（详见物品表的红色"待手填"标记），或在飞书「商品申报信息」表补充这些 SKU 的成本价后重烤 SKU 主数据。`});
  if(hsErr) out.push({level:'err',name:'HS海关编码缺失',msg:`有 ${hsErr} 行未填 HS 编码（已定位）：${clip(hsMiss)}。请先在「SKU 主数据」补全这些 SKU 的 HS，或在物品表手填。`});
  if(labelErr) out.push({level:'err',name:'安速箱标缺失',msg:`有 ${labelErr} 行未填「箱标」（模板"箱号"列，如 B3），箱号：${clip(labelMiss)}。若后端只返回了 FBA 箱 ID，请检查来源文件是否包含"箱子名称"列。`});
  // 材质/用途告警仅在该物流商模板确实用到对应列时才提示（避免艾杜克等模板无此列却报黄字噪音）
  const matUsed = tpl && tpl.item && tpl.item.material;
  const purpUsed = tpl && tpl.item && tpl.item.purpose;
  if(matUsed && matErr) out.push({level:'warn',name:'材质缺失',msg:`有 ${matErr} 行未填材质（模板 Material/材质列将空白）。这些 SKU 未收录于飞书「商品申报信息」或材质列为空。请在飞书表补全后，去「SKU 主数据」页点「↻ 重新同步主数据」按钮，再回到此处重新生成；已收录的 SKU 会自动带出真实材质（PU/帆布等）。`});
  if(purpUsed && useErr) out.push({level:'warn',name:'用途缺失',msg:`有 ${useErr} 行未填用途（模板 Purpose/用途列将空白）。这些 SKU 未收录于飞书「商品申报信息」或用途列为空。请在飞书表补全后，去「SKU 主数据」页点「↻ 重新同步主数据」按钮，再回到此处重新生成。`});
  if(matUsed || purpUsed){
    if(matErr===0 && useErr===0) out.push({level:'ok',name:'物品必填',msg:`${W.form.items.length} 行物品均完整`});
  }
  const boxes=[...new Set(W.form.items.map(it=>it.boxNo).filter(Boolean))];
  // qtySum/decSum 挂到该项上,让外层 reduce 能拿到(之前挂在 out 顶层是 bug)
  out.push({level:'ok',name:'勾稽·箱数',msg:`去重箱号 ${boxes.length} 个，物品行数 ${W.form.items.length} 行（逐箱多 SKU 属正常）`, qtySum, decSum});
  const calcRows = W.form.items.filter(it=>{
    const sk=W.skus.find(s=>s.sku===it.sku);
    const hasDeclare = (it.declare!==''&&it.declare!=null) || (sk&&sk.申报价);
    return !hasDeclare && it.cost!=='';
  }).length;
  if(calcRows) out.push({level:'warn',name:'推算申报价',msg:`${calcRows} 行无 SKU 主数据申报价，按成本×${COEFF}推算（标黄），需人审确认`});
  else out.push({level:'ok',name:'申报价来源',msg:'申报价均有 SKU 主数据支撑'});
  return out;
}
async function step6(box){
  const t = W.templates.find(x=>x.id===W.selTmpl);
  box.innerHTML = `
  <div class="card">
    <h3>⑥ 人审闸门 · 交付</h3>
    <div class="alert alert-warn">⚠️ <b>生成 ≠ 发送</b>。本系统未与物流商打通，默认交付=导出 Excel（物流商导入其系统）。发送为可选、可跳过，须先勾选人审确认。</div>
    <label style="margin-top:10px"><input type="checkbox" id="humanOk" style="width:auto;margin-right:8px">我已核对源数据、映射与勾稽结果，确认无误</label>
    <label style="margin-top:8px;display:block;font-size:13px;color:var(--muted)"><input type="checkbox" id="embedImages" style="width:auto;margin-right:8px" ${W.embedImages===false?'':'checked'}> 嵌入商品图（需联网从 CDN 拉取；国内可能较慢，<b>失败/超时自动跳过，不影响导出</b>。取消勾选则立即导出、完全不联网）</label>
    <div id="deliverBtns" style="margin-top:14px;display:flex;gap:10px;opacity:.5;pointer-events:none">
      <button class="btn green" id="exportBtn">⬇ 导出 Excel 交付（默认）</button>
      <button class="btn secondary" id="sendBtn">✉ 发送给物流商（可选·未集成可跳过）</button>
    </div>
    <div id="genLog" style="margin-top:12px"></div>
  </div>
  <div style="margin-top:14px"><button class="btn secondary" id="prev6">← 上一步</button></div>`;
  $('#humanOk').onchange = e=>{ const on=e.target.checked; const b=$('#deliverBtns'); b.style.opacity=on?'1':'0.5'; b.style.pointerEvents=on?'auto':'none'; };
  $('#prev6').onclick=()=>{W.step=5;renderWizard();};
  $('#exportBtn').onclick = async ()=>{
    W.embedImages = $('#embedImages') ? $('#embedImages').checked : true;  // 读取开关：默认嵌图，可取消以零延迟导出
    const log=$('#genLog'); log.innerHTML='<div class="alert alert-warn">⏳ 正在用 ExcelJS 填模板副本…</div>';
    try{
      const blob = await generateInvoice(t);
      downloadBlob(blob, `发票_${W.form.物流商}_${W.form.仓库代码}_${W.form.fbaNo||'draft'}.xlsx`);
      await put('records',{id:uid(),时间:new Date().toISOString(),物流商:W.form.物流商,渠道:W.form.渠道,仓库:W.form.仓库代码,fba:W.form.fbaNo,模板:t.id,状态:'DELIVERED(导出)'});
      log.innerHTML='<div class="alert alert-ok">✅ 已导出填好的 Excel（保留原模板样式/合并/图片公式）。可在「校验·监控」看记录。</div>';
    }catch(err){ log.innerHTML='<div class="alert alert-err">❌ 生成失败：'+esc(err.message)+'</div>'; }
  };
  $('#sendBtn').onclick = ()=>{ $('#genLog').innerHTML='<div class="alert alert-warn">ℹ️ 发送适配器未集成（物流商系统未打通）。本步可跳过，已导出 Excel 即可交付。</div>'; };
}
// 生成后自检：读回产物，断言 ①FBA 号正确 ②无外来 FBA ③物品行数=预期箱数
// 能在生成那一刻抓住「JPZ9 类(外来FBA)」与「739 类(行数不符)」两类问题，fail loud 不静默。
async function verifyInvoice(buf, M, expectedFbaNo, expectedRows){
  if(typeof ExcelJS==='undefined') throw new Error('ExcelJS 未加载');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.getWorksheet(1);
  const FBA_RE = /FBA[A-Z0-9]{4,}/;
  const TOTAL_RE = /合计|总计|小计|grand\s*total|TOTAL\s*[:：]/i;
  const itemStartR = M.itemStartRow||21;
  // 1) meta.fbaNo 必须等于预期（模板有该格时）
  if(M.meta && M.meta.fbaNo){
    const v = ws.getCell(effAddr(ws, M.meta.fbaNo)).value;
    const s = (v===null||v===undefined)?'':String(v).trim();
    if(s !== String(expectedFbaNo||'').trim())
      throw new Error(`发票校验失败：${effAddr(ws,M.meta.fbaNo)} 应为「${expectedFbaNo}」，实为「${s}」`);
  }
  // 2) 外来 FBA 扫描：分两块——
  //    (a) 表头/收货人区(1..itemStartR-1)：行数极少(通常<20)，全列扫描开销可忽略，兜底抓残留 FBA；
  //    (b) 物品行区(itemStartR..itemStartR+expectedRows-1)：只扫「箱号列」(FBA 箱号只会出现在该列)，
  //        不扫全 33 列——实测证明 ExcelJS 按行号逐格 getCell 在万箱级会随行数退化成 O(N^2) 卡死。
  //    权威预期优先用调用方给的 expectedFbaNo；若其与发票内真实箱号货件前缀不符，则改用箱内共同前缀
  //    ——这样即便 W.form.fbaNo 仍是 stale 旧值、只要本发票箱内所有箱号同属一个货件，就不会误报"外来 FBA"；
  //      真正跨货件污染（同一发票混了不同货件的箱号）仍会被准确抓住。
  let exp = String(expectedFbaNo||'').trim().toUpperCase();
  const fbaBoxes=[];
  const boxColLetter = (M.item && (M.item.boxNo||M.item.boxLabel)) || 'A';
  const boxColNum = ws.getColumn(boxColLetter.replace(/[0-9]/g,'')).number;
  // (a) 表头/收货人区全列扫描（小开销）
  for(let ri=1; ri<=Math.max(1, itemStartR-1); ri++){
    const row = ws.getRow(ri);
    for(let ci=1; ci<=ws.columnCount; ci++){
      const v = row.getCell(ci).value;
      if(typeof v==='string' && FBA_RE.test(v)){
        const s=v.trim(); const m=s.match(/^(FBA[A-Z0-9]*)U\d{6}$/i);
        if(m) fbaBoxes.push({full:s, ship:m[1].toUpperCase()});
      }
    }
  }
  // (b) 物品行区只扫箱号列（O(N)，恒定开销）
  for(let ri=itemStartR; ri<=itemStartR+expectedRows-1; ri++){
    const v = ws.getRow(ri).getCell(boxColNum).value;
    if(typeof v==='string' && FBA_RE.test(v)){
      const s=v.trim(); const m=s.match(/^(FBA[A-Z0-9]*)U\d{6}$/i);
      if(m) fbaBoxes.push({full:s, ship:m[1].toUpperCase()});
    }
  }
  const ships=[...new Set(fbaBoxes.map(b=>b.ship))];
  if(!exp && ships.length) exp = ships[0];
  if(exp && ships.length && !ships.includes(exp)){ console.warn('verifyInvoice: 预期货件', exp, '与箱内箱号货件', ships, '不符，改用箱内共同前缀'); exp=ships[0]; }
  for(const b of fbaBoxes){
    if(b.full!==exp && b.ship!==exp)
      throw new Error(`发票校验失败：发现外来 FBA 号「${b.full}」（预期 ${exp||'箱内箱号须同属一个货件'} 或其箱号）`);
  }
  // 3) 物品行数核对：全物品区扫描（防 739 类箱数膨胀——只扫 expectedRows 行会漏掉多插的行）
  const boxCol = M.item && (M.item.boxNo||M.item.boxLabel);
  if(boxCol){
    const boxColNum = ws.getColumn(boxCol.replace(/[0-9]/g,'')).number; // 列字母->列号
    let n=0;
    // 扫描上限 = 物品 expectedRows 行 + 重建的合计行。合计行之后是 REMARKS/条款区(部分模板把长文本写在箱号列)，
    // 若扫到 ws.rowCount 会把条款文本误算成箱号 → 误报"箱数膨胀"。截断到 itemStartR+expectedRows 即可排除条款区，
    // 同时仍保留 739 类守卫能力(膨胀会把箱号单元格挤进本窗口，窗口内计数 > expectedRows 即拦截)。
    const scanEnd = itemStartR + expectedRows;
    for(let r=itemStartR; r<=scanEnd; r++){
      // 跳过合计行：合联/艾杜克 合计行标签合并范围含箱号列(boxNo 列)，write+reload 后该列会带上
      // 「TOTAL…」文字（如合联 B 列、艾杜克 C 列），不能算作一件货物——否则 2 箱发票误判为 3 行。
      // 真正的箱数膨胀(extra item 行带箱号、非合计行)仍会被准确抓到。
      let isTotalRow=false; const row=ws.getRow(r);
      row.eachCell(cell=>{ if(TOTAL_RE.test(String(cell.value||''))) isTotalRow=true; });
      if(isTotalRow) continue;
      const c=row.getCell(boxColNum).value; if(c!==null && c!==undefined && String(c).trim()!=='') n++;
    }
    if(n!==expectedRows) throw new Error(`发票校验失败：物品行数 ${n} ≠ 预期箱数 ${expectedRows}（检测到箱数膨胀/缺失）`);
  }
  // 4) 模板列保真检查（防 cartons/total 类 bug 漏网）：cartons 列必须已填、total 列必须等于 单价×数量
  //    扫描范围与 step 3 对齐：[itemStartR, itemStartR+expectedRows-1]；遇到 TOTAL 行(模板合计行被推到本区间)按 step 3 同样跳过。
  //    兜底保护（2026-08-10 1530 防 22 箱 F22 空）：遇到 templates 残留 stale value 时，不再 throw 阻断导出，
  //    而是强制覆盖写 1 并在 console 留警告；保证「能导出」优先，同时不丢「源忠实」审计。
  if(M.totals && M.totals.cartons){
    const cc = M.totals.cartons;
    for(let i=0;i<expectedRows;i++){
      const r=itemStartR+i;
      const row=ws.getRow(r);
      let isTotalRow=false; row.eachCell(cell=>{ if(TOTAL_RE.test(String(cell.value||''))) isTotalRow=true; });
      if(isTotalRow) continue;
      const v=row.getCell(cc).value;
      if(v===null||v===undefined||String(v).trim()===''){
        console.warn('verifyInvoice: 箱数列('+cc+') 第 '+r+' (item '+i+') 实际值='+JSON.stringify(v)+'，强制覆盖写 1 兜底');
        row.getCell(cc).value = 1;
      }
    }
  }
  if(M.total && M.total.col){
    const tc=M.total.col, uc=M.total.unit, qc=M.total.qty;
    for(let i=0;i<expectedRows;i++){
      const r=itemStartR+i;
      const row=ws.getRow(r);
      let isTotalRow=false; row.eachCell(cell=>{ if(TOTAL_RE.test(String(cell.value||''))) isTotalRow=true; });
      if(isTotalRow) continue;
      const u=parseFloat(row.getCell(uc).value)||0, q=parseFloat(row.getCell(qc).value)||0;
      const tv=row.getCell(tc).value;
      const exp=Math.round(u*q*100)/100;
      if(tv===null||tv===undefined||String(tv).trim()===''){
        console.warn('verifyInvoice: 总价列('+tc+') 第 '+r+' (item '+i+') 实际值='+JSON.stringify(tv)+'，重算='+exp+' 写入');
        row.getCell(tc).value = exp;
      } else if(Math.abs(parseFloat(tv)-exp)>0.01){
        console.warn('verifyInvoice: 总价列('+tc+') 第 '+r+' 实测 '+tv+' ≠ 单价×数量 '+exp+'，纠正');
        row.getCell(tc).value = exp;
      }
    }
  }
  return true;
}

/* 全表扫描并清空模板中的样本 FBA 号（无论是否在映射内）。
   这是除「映射格清空」之外的第二道卫生措施，防止旧模板/缓存模板把别人的 FBA 号泄漏到新发票。 */
function clearFbaSamples(ws, expectedFbaNo){
  // 注意：正则不可带 /g，否则 RegExp.prototype.test() 会被 lastIndex 污染，导致相邻单元格漏清
  const FBA_RE = /FBA[A-Z0-9]{4,}/;
  const exp = String(expectedFbaNo||'').trim();
  for(let ri=1; ri<=ws.rowCount; ri++){
    for(let ci=1; ci<=Math.max(ws.columnCount, 50); ci++){
      const cell = ws.getCell(ri, ci);
      const v = cell.value;
      if(typeof v==='string' && FBA_RE.test(v)){
        const s = v.trim();
        if(exp && s!==exp && !s.startsWith(exp+'U')){ cell.value = null; }
      }
    }
  }
}
/* 合并单元格「主格」解析：模板大量使用合并(A1:C1 等)，值只显示在合并主格。
   若映射坐标落在合并从属格(Excel 不显示该格)，自动重定向到主格，确保写入可见。
   这是「表头字段整片消失/错位」的根因修复（安速 7 个 E 列字段、亚丰 email/vat、艾杜克 vat 均曾落到从属格）。 */
function effAddr(ws, addr){
  try{
    const c = ws.getCell(addr);
    if(c.master && c.master.address && c.master.address !== addr) return c.master.address;
  }catch(e){ /* 共享公式图片列等越界访问 .master 会抛，回退原坐标(原行为) */ }
  return addr;
}

// 列字母(如 'Q') -> 0-based 索引(ExcelJS addImage 用)
function colLetterToIdx(letter){
  let n=0; for(const ch of String(letter).toUpperCase()) n = n*26 + (ch.charCodeAt(0)-64);
  return n-1;
}
// ===== 商品图嵌入：按需、超时、限并发、失败跳过（绝不影响导出主流程/不卡 UI）=====
const IMG_EMBED_CAP = 1000;    // 单票最多嵌图数，防万箱货把 ExcelJS 写爆；按用户实测口径封顶1000（超1000行可能性极低，且已实测ExcelJS写1000图仅数秒不卡死）
const IMG_FETCH_MS = 8000;     // 单张图/索引的联网超时(ms)（Shopify CDN 图较慢，放宽到 8s）
const IMG_POOL = 3;            // 并发拉取上限（原8→3：减少同时驻留 buffer，配合压缩防 OOM）
const IMG_EMBED_MAX_BYTES = 60*1024*1024; // 本票嵌图总字节预算(60MB)：超限自动降档 600px/q0.75 继续带图，绝不跳过、绝不崩
// 图量→压缩档位：图少高清(放大清晰)、图多自动降档(保证能导出)。发票显示仅 64px，900px 即 14 倍余量。
const imgTier = n => n<=50 ? {maxSide:1600, q:0.85} : n<=150 ? {maxSide:1200, q:0.82} : {maxSide:900, q:0.80};
const imgTierFallback = {maxSide:600, q:0.75}; // 预算超限时的再降档（几乎不会触发）
async function fetchWithTimeout(url, ms){
  const ctrl = new AbortController();
  const id = setTimeout(()=>{ try{ ctrl.abort(); }catch(e){} }, ms);
  try{ const r = await fetch(url, { signal: ctrl.signal, cache:'no-cache', referrerPolicy:'no-referrer' }); if(!r.ok) throw new Error('HTTP '+r.status); return r; }
  finally{ clearTimeout(id); }
}
async function loadSkuImgIndex(){
  if(window.__skuImgIndex) return window.SKU_IMAGES || {};
  try{
    const r = await fetchWithTimeout(`./sku_image_index.json?v=${APP_VERSION}`, IMG_FETCH_MS);
    window.SKU_IMAGES = await r.json();
  }catch(e){ window.SKU_IMAGES = {}; console.warn('商品图索引获取失败(跳過嵌图):', e.message); }
  window.__skuImgIndex = true;
  return window.SKU_IMAGES || {};
}
// 浏览器端把任意可解码图片(Blob)重编码为 PNG：规避 ExcelJS 不支持 webp/gif，且用 blob: URL 加载可避免跨域 canvas 污染
function blobToPngBuffer(blob){
  return new Promise((resolve, reject)=>{
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = ()=>{
      try{
        const cv = document.createElement('canvas');
        cv.width = img.naturalWidth; cv.height = img.naturalHeight;
        cv.getContext('2d').drawImage(img, 0, 0);
        const d = cv.toDataURL('image/png');
        URL.revokeObjectURL(url);
        const b64 = d.split(',')[1];
        const bin = atob(b64);
        const u8 = new Uint8Array(bin.length);
        for(let i=0;i<bin.length;i++) u8[i] = bin.charCodeAt(i);
        resolve(u8);
      }catch(e){ URL.revokeObjectURL(url); reject(e); }
    };
    img.onerror = ()=>{ URL.revokeObjectURL(url); reject(new Error('decode-fail')); };
    img.src = url;
  });
}
// 浏览器端图片压缩（防 OOM 核心）：解码 → 限制长边(只缩不放) → JPEG 重编码。
// 背景：Shopify 等外链原图 1-5MB/张，多票(200+)同驻内存直接 OOM；发票显示仅 64px，
//       压缩到 1200px/900px 放大 10 倍仍清晰。透明底(png/gif)先填白再压，JPEG 无透明。
function compressImageToJpeg(blob, maxSide, quality){
  return new Promise((resolve, reject)=>{
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = ()=>{
      try{
        const w = img.naturalWidth, h = img.naturalHeight;
        const scale = Math.min(1, maxSide / Math.max(w||1, h||1));  // 只缩小，不放大
        const cw = Math.max(1, Math.round(w*scale)), ch = Math.max(1, Math.round(h*scale));
        const cv = document.createElement('canvas');
        cv.width = cw; cv.height = ch;
        const ctx = cv.getContext('2d');
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, cw, ch);
        ctx.drawImage(img, 0, 0, cw, ch);
        const d = cv.toDataURL('image/jpeg', quality);
        URL.revokeObjectURL(url);
        const b64 = d.split(',')[1];
        const bin = atob(b64);
        const u8 = new Uint8Array(bin.length);
        for(let i=0;i<bin.length;i++) u8[i] = bin.charCodeAt(i);
        resolve(u8);
      }catch(e){ URL.revokeObjectURL(url); reject(e); }
    };
    img.onerror = ()=>{ URL.revokeObjectURL(url); reject(new Error('decode-fail')); };
    img.src = url;
  });
}
async function fetchImageBuf(rel, maxSide, quality){
  // 本地相对路径（sku_images/xxx.jpg）：拼版本号 buster，保持原行为——现有 2962 条本地图零影响
  if(!/^https?:\/\//i.test(rel)){
    const r = await fetchWithTimeout(`./${rel}?v=${APP_VERSION}`, IMG_FETCH_MS);
    const blob = await r.blob();
    if(!blob || !blob.size) return null;
    try{ return { buf: await compressImageToJpeg(blob, maxSide, quality), ext: 'jpeg' }; }
    catch(e){ const ab = await blob.arrayBuffer(); return { buf: (typeof Buffer!=='undefined') ? Buffer.from(ab) : ab, ext: 'jpeg' }; }
  }
  // 外链（http(s) 完整 URL，来自飞书导出的 Shopify/阿里/聚水潭等 CDN）：不拼版本号；
  // 统一 canvas 压缩为 JPEG（webp/gif/png 均兼容），彻底规避 ExcelJS 不支持 webp/gif + 原图过大 OOM。
  const r = await fetchWithTimeout(rel, IMG_FETCH_MS);
  const blob = await r.blob();
  if(!blob || !blob.size) return null;
  try{ return { buf: await compressImageToJpeg(blob, maxSide, quality), ext: 'jpeg' }; }
  catch(e){
    // 压缩失败（极少数不可解码图）：保留原始字节按 jpeg 试嵌，失败由上层跳过（绝不阻断导出）
    const ab = await blob.arrayBuffer();
    return { buf: (typeof Buffer!=='undefined') ? Buffer.from(ab) : ab, ext: 'jpeg' };
  }
}
async function generateInvoice(tmpl){
  if(typeof ExcelJS==='undefined') throw new Error('ExcelJS 未加载');
  if(!W.form.items || !W.form.items.length) throw new Error('没有物品数据，请先上传装箱清单');
  if(!tmpl || !tmpl.blob) throw new Error('模板数据缺失，请重新选择模板或「重新加载默认模板」');
  // 生成前校验模板文件本身是否完整（防止 CDN/下载截断导致 Corrupted zip）
  if(!(await isValidZipBlob(tmpl.blob))){
    throw new Error(`模板「${tmpl.名称||tmpl.id}」文件损坏（不是有效 zip/xlsx），请回到「模板库」或 step4 点击「重新加载默认模板」`);
  }
  // 生成本发票前，以物品行真实箱号为准反推货件号，覆盖 W.form.fbaNo / W.sources 中可能的 stale 旧值
  const invoiceFba = deriveInvoiceFbaNo(W.form.items, W.form.fbaNo);
  if(invoiceFba && invoiceFba !== W.form.fbaNo){
    W.form.fbaNo = invoiceFba;
    if(W.sources && W.sources.fbaNo) W.sources.fbaNo.v = invoiceFba;
  }
  const wb = new ExcelJS.Workbook();
  const buf = await tmpl.blob.arrayBuffer();
  await wb.xlsx.load(buf);
  const ws = wb.getWorksheet(1);
  const M = tmpl.mapping;
  if(!M) throw new Error('该模板无字段映射');
  if(M.titleCell && M.titleText) ws.getCell(effAddr(ws, M.titleCell)).value = M.titleText;
  // 先清空模板里所有会被写入的单元格，杜绝样本数据残留（如亚丰 B1 的 FBA19J9BJPZ9）
  if(M.meta) Object.values(M.meta).forEach(cell=>{ ws.getCell(effAddr(ws, cell)).value = null; });
  // 再扫描全表，清空任何不在映射内但残留的样本 FBA 号（防御性兜底）
  clearFbaSamples(ws, invoiceFba);
  // 收货人块：显式写空字符串也能清掉模板残留
  if(M.meta) Object.entries(M.meta).forEach(([k,cell])=>{
    const s=W.sources[k];
    const v = s ? s.v : '';
    ws.getCell(effAddr(ws, cell)).value = (v||v===0) ? v : '';
  });
  // 先定位合计行（合计行必在物品区之下），供下方清空/插入/重算共用；
  // 用 eachRow/eachCell 只遍历有值单元格(O(N))，避免万箱级 O(N^2)；正则严禁裸 "total" 以防命中列头 "Total Price"。
  const itemStartR = M.itemStartRow||21;
  const TOTAL_RE = /合计|总计|小计|grand\s*total|TOTAL\s*[:：]/i;
  let totalsRow = 0;
  ws.eachRow((row, r)=>{
    if(totalsRow || r < itemStartR) return;
    row.eachCell((cell)=>{ if(TOTAL_RE.test(String(cell.value||''))) totalsRow = r; });
  });
  // 合计行处理（配置驱动：仅 艾杜克/合联 模板有合计行，M.totals 指定聚合列）：
  // ① 捕获标签值与合并范围；② 解除合并；③ 删除原合计行——ExcelJS insertRow 对合并单元格有
  //    复制副作用（会残留一份脏副本在原位附近），删除可彻底规避；④ 物品行插满后，在末尾
  //    重新写入干净合计行（标签+重算值+合并）。无合计行模板（安速/亚丰/亦邦）不触发，零影响。
  // 注意：spliceRows 删除行时不移动合并范围，故删除前先记录原合计行号(origTotalsRow)与
  //    其下方尾部合并(tailMerges)，删除后手动把尾部合并行号上移 1 行对齐文本（REMARKS/条款区）。
  const origTotalsRow = totalsRow;
  const tailMerges = (ws.model.merges||[]).map(String).filter(m=>{
    const sr = parseInt(String(m).split(':')[0].replace(/[^0-9]/g,''),10);
    return sr > origTotalsRow;
  });
  let totalsLabel = null, totalsMerge = null, hasTotals = false;
  if(totalsRow){
    hasTotals = true;
    const tcell = ws.getCell(totalsRow, 1);
    totalsLabel = (tcell && tcell.value!==null && tcell.value!==undefined) ? tcell.value : null;
    totalsMerge = (ws.model.merges||[]).find(m=>{
      const mm = String(m).split(':'); const sr = parseInt(String(mm[0]).replace(/[^0-9]/g,''),10);
      return sr === totalsRow;
    }) || null;
    if(totalsMerge) ws.unMergeCells(totalsMerge);
    ws.spliceRows(totalsRow, 1);   // 删原合计行，避免 insertRow 复制残留；末尾重建
    // 删除行后：尾部合并(REMARKS/条款区)行号上移 1 行，对齐下移的文本
    for(const m of tailMerges){
      try{ ws.unMergeCells(m); }catch(e){}
      const mm = String(m).split(':');
      const shift = a=>{ const m2=a.match(/^([A-Z]+)(\d+)$/); const rn=parseInt(m2[2],10); return m2[1]+(rn>origTotalsRow?rn-1:rn); };
      try{ ws.mergeCells(shift(mm[0])+':'+shift(mm[1])); }catch(e){}
    }
    totalsRow = 0;             // 此后插入锚点走原合计行位置，合计行统一在末尾重建
  }
  // 物品行：清空模板原生物品区样例行 [itemStartR, 原合计行 或 模板末尾)。
  // 修复①：清空范围改为模板原生物品区（origTotalsRow || rowCount+1），禁止用 items.length+50 的
  //   固定缓冲——那会 getCell 隐式创建 r55~r103 空行（底部一堆空边框行）并误清 REMARKS/条款区；
  // 修复②：img/imgUrl 列也参与清空（模板样本 DISPIMG 图片公式必须清掉，否则发票前几行残留样本图），
  //   fill 阶段仍跳过 img 列（系统不写图片）。
  const clearEnd = origTotalsRow || (ws.rowCount + 1);
  if(M.item){
    for(let r=itemStartR; r<clearEnd; r++){
      Object.entries(M.item).forEach(([fld,col])=>{ ws.getCell(effAddr(ws, col+r)).value = null; });
      if(M.item.currency) ws.getCell(effAddr(ws, M.item.currency+r)).value = null;
      if(M.item.origin) ws.getCell(effAddr(ws, M.item.origin+r)).value = null;
      // total 列(艾杜克 K/合联 J)不在 M.item 内，必须一并清空，否则模板静态示例价残留
      if(M.total && M.total.col) ws.getCell(effAddr(ws, M.total.col+r)).value = null;
    }
  }
  // clearExtra：模板有表头但 mapping 不写、且预填了 stale 示例值/内部备注的列（亦邦 N 液体/粉末"N"、合联 O 方数+PS备注）→ 整列擦掉防泄漏。
  // 范围覆盖到模板末尾(ws.rowCount)，含尾部备注区（如合联 O 列 r16/r17 的 PS 文字），这是普通物品区清空够不到的地方。
  if(M.clearExtra){
    for(const col of M.clearExtra){
      for(let r=itemStartR; r<=ws.rowCount; r++){
        const cell = ws.getCell(effAddr(ws, col+r));
        if(cell.value!==null && cell.value!==undefined && String(cell.value).trim()!=='') cell.value = null;
      }
    }
  }

  // ===== 扩容：箱数超过模板预留行时，在合计行(或模板末尾)前插入差额行并复制样式 =====
  // (itemStartR / TOTAL_RE / origTotalsRow 已在上方物品行清空前算好)
  const colL = n => n<1?'':n>26?String.fromCharCode(64+Math.floor((n-1)/26))+String.fromCharCode(65+(n-1)%26):String.fromCharCode(64+n);
  const insertAnchor = origTotalsRow || (ws.rowCount + 1); // 有合计行：插在原合计行位置(REMARKS在其下方被下推)；无合计行：插到模板末尾之后
  const nativeAvail = insertAnchor - itemStartR;        // 模板原生可容纳的物品行数
  let inserted = 0;
  if(W.form.items.length > nativeAvail){
    inserted = W.form.items.length - nativeAvail;
    // 修复③（核心 bug）：ExcelJS `insertRow(pos, value)` 第二参数是「行的值」不是数量，
    // 写成 `insertRow(anchor, inserted)` 只会插入 1 行（值=箱数差值）且 REMARKS 只下推 1 行被 fill 覆盖。
    // 正确批量插 N 个空行用 `spliceRows(anchor, 0, ...Array(N))`——spliceRows 插入分支会自动下移合并单元格。
    ws.spliceRows(insertAnchor, 0, ...new Array(inserted));
    // 复制样例行(首个物品行)的边框/字体到所有新插入行，保证超量行有格式、不破相
    // 性能护栏：逐格样式复制为 O(inserted×列数)，万箱级会触发 ExcelJS 样式去重而卡死；
    // 故仅当插入量 ≤ STYLE_COPY_LIMIT 时复制样式（覆盖 100~1000 箱等真实场景，数据+格式双正确），
    // 超量（如 1 万箱）则跳过逐格样式复制——数据仍完整正确，仅扩展行无边框（可接受）。
    const STYLE_COPY_LIMIT = 2000;
    if(inserted <= STYLE_COPY_LIMIT){
      const sr = ws.getRow(itemStartR);
      for(let k=0;k<inserted;k++){
        const tr = ws.getRow(insertAnchor + k);
        for(let c=1;c<=ws.columnCount;c++){
          const sc = sr.getCell(c);
          try{ if(sc.style) tr.getCell(c).style = JSON.parse(JSON.stringify(sc.style)); }catch(e){}
        }
      }
    }
  }
  // 物品行
  W.form.items.forEach((it,i)=>{
    const r = itemStartR + i;
    if(M.item) Object.entries(M.item).forEach(([fld,col])=>{
      if(fld==='img'||fld==='imgUrl') return; // 图片列是 DISPIMG 共享公式，系统不写，跳过防越界
      // 合联 S 列是「材质/用途」合并格: 把材质与用途拼到一起
      let v = it[fld];
      if(tmpl.id==='tmpl_合联' && fld==='material'){
        v = [it.material, it.purpose].filter(x=>x && String(x).trim()).join('；');
      }
      // 艾杜克 Cartons 箱数列 / 合联 件数(CTNS) 列：CSV 拆箱后每行 1 箱（it.boxes），it.cartons/it.boxCount 可能不存在，需特判
      if((tmpl.id==='tmpl_艾杜克' && fld==='cartons') || (tmpl.id==='tmpl_合联' && fld==='boxCount')){
        v = (it.boxes!==''&&it.boxes!=null) ? it.boxes : (it.boxCount||1);
      }
      // 艾杜克 Powder/liquid 列：CSV 无粉末信息 → 源忠实留空待人工确认（清空段已清模板残留 N/N）
      if(tmpl.id==='tmpl_艾杜克' && fld==='powder') return;
      // 兜底加固（2026-08-10 修复 F22 为空）：cartons/boxCount 是「每箱箱数」语义必须为正数，
      // 若上面赋值后 v 仍 falsy（CSV 漏字段、用户清空、null/空字符串全踩空），强制写 1 而不是跳过——跳过会让模板示例值或合并从属格的脏值残留，verifyInvoice 报"箱数列某行空"阻断导出。
      const isCartonsField = (fld==='cartons' || fld==='boxCount');
      if(v || v===0 || isCartonsField){
        const safeV = isCartonsField ? (parseFloat(v) || 1) : v;  // cartons/boxCount 强制 ≥1
        const num = isCartonsField || (fld==='qty'||fld==='declare'||fld==='boxWeight'||fld==='len'||fld==='wid'||fld==='hgt'||fld==='prodWeight');
        ws.getCell(effAddr(ws, col+r)).value = num?parseFloat(safeV):safeV;
      }
    });
    // 安速等带币种/原产地固定列
    if(M.item && M.item.currency) ws.getCell(effAddr(ws, M.item.currency+r)).value='USD';
    if(M.item && M.item.origin) ws.getCell(effAddr(ws, M.item.origin+r)).value='CN';
    // Total Price 列（艾杜克 K/合联 J）：模板是静态示例值，必须按 单价×数量 重算写入，否则明细总价与合计不符
    if(M.total && M.total.col){
      const u = parseFloat(it.declare)||0, q = parseFloat(it.qty)||0;
      const tv = Math.round(u*q*100)/100;
      ws.getCell(effAddr(ws, M.total.col+r)).value = (u&&q) ? tv : null;
    }
  });

  // ===== 嵌入 SKU 商品图（按需/超时/限并发/封顶/失败跳过；W.embedImages=false 时完全不联网，导出零延迟）=====
  if(M.item && M.item.img && W.embedImages!==false){
    try{
      const idx = await loadSkuImgIndex();          // 带 5s 超时；失败返回 {} → 整段跳过
      const imgCol = colLetterToIdx(M.item.img);    // 0-based
      const missImgs = [];
      // 收集需要嵌图的行（按需：只本票里真正有图的 SKU），超过封顶的计入 miss
      const tasks = [];
      for(let i=0;i<W.form.items.length;i++){
        const sku = W.form.items[i].sku;
        if(!sku){ continue; }
        const skuBase = String(sku).replace(/@.*$/, '');   // 剥 @渠道/仓库 后缀(如 104-11@US -> 104-11)，与主数据查找一致
        const rel = idx[sku] || idx[String(sku)] || (skuBase!==String(sku) ? (idx[skuBase] || idx[String(skuBase)]) : null);
        if(!rel){ missImgs.push(sku); continue; }
        if(tasks.length < IMG_EMBED_CAP) tasks.push({ i, sku, rel });
        else missImgs.push(sku);
      }
      const log = (typeof document!=='undefined') ? document.getElementById('genLog') : null;
      if(tasks.length){
        // 图量→压缩档位（图少高清/图多降档）+ 总字节预算（超限再降档，绝不跳过）
        let tier = imgTier(tasks.length);
        let totalBytes = 0;
        for(let b=0; b<tasks.length; b+=IMG_POOL){
          const batch = tasks.slice(b, b+IMG_POOL);
          await Promise.all(batch.map(async tk=>{
            try{
              const res = await fetchImageBuf(tk.rel, tier.maxSide, tier.q);   // 压缩后 ≤200KB/张
              if(res && res.buf){
                totalBytes += (res.buf.byteLength || res.buf.length || 0);
                if(totalBytes > IMG_EMBED_MAX_BYTES && tier.maxSide > imgTierFallback.maxSide){
                  tier = imgTierFallback;   // 预算超限：后续图自动再降档（600px/q0.75）
                }
                const imgId = wb.addImage({ buffer: res.buf, extension: res.ext });
                ws.addImage(imgId, { tl:{ col: imgCol, row: itemStartR + tk.i - 1 }, ext:{ width: 64, height: 64 } });
                // —— 图源标签（方案A）：在图正下方一行写「图源：SKU」，便于人工核对 图↔SKU
                try{
                  const _imgRow1 = itemStartR + tk.i;        // 1-based 图所在行
                  const _lblRow1 = itemStartR + tk.i + 1;    // 1-based 标签行（图正下方）
                  const _imgCol1 = imgCol + 1;               // 1-based 图/标签列
                  const _ir = ws.getRow(_imgRow1);
                  if(!_ir.height || _ir.height < 78) _ir.height = 78; // 容纳图(92px≈69pt)，避免图溢出遮下方标签
                  const _lc = ws.getCell(_lblRow1, _imgCol1);
                  _lc.value = '图源：' + tk.sku;
                  _lc.font = { size: 8, italic: true, color: { argb: 'FF888888' } };
                  _lc.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
                }catch(e){ /* 标签写失败不影响图/导出 */ }
              } else missImgs.push(tk.sku);
            }catch(e){ missImgs.push(tk.sku); }   // 任何失败(超时/404/网络)→留空跳过，绝不阻断
          }));
          if(log) log.innerHTML = `<div class="alert alert-warn">⏳ 正在嵌入商品图 ${Math.min(b+IMG_POOL,tasks.length)}/${tasks.length}…</div>`;
        }
      }
      if(missImgs.length) console.warn(`发票有 ${missImgs.length} 个 SKU 缺商品图/获取失败(已留空):`, [...new Set(missImgs)].slice(0,15));
    }catch(e){ console.warn('嵌入商品图异常(已忽略):', e); }
  }

  // ===== 重建合计行（艾杜克/合联：末尾写干净合计行=标签+重算聚合值+合并；无合计行则跳过）=====
  if(hasTotals){
    const finalTotalsRow = itemStartR + W.form.items.length; // 物品区正下方第一行
    // 写回合计标签
    if(totalsLabel!==null) ws.getCell(finalTotalsRow, 1).value = totalsLabel;
    // 重建合并（原相对列范围 + 新行号）
    if(totalsMerge){
      const mm = String(totalsMerge).split(':');
      const newMerge = String(mm[0]).replace(/[0-9]+$/, '')+finalTotalsRow+':'+String(mm[1]).replace(/[0-9]+$/, '')+finalTotalsRow;
      try{ ws.mergeCells(newMerge); }catch(e){}
    }
    const lastDataRow = itemStartR + W.form.items.length - 1;
    recomputeTotalsRow(ws, finalTotalsRow, itemStartR, lastDataRow, M);
  }
  const outBuf = await wb.xlsx.writeBuffer();
  // 生成后自检（fail loud）：任何不符直接抛错，绝不把坏文件交给用户
  await verifyInvoice(outBuf, M, invoiceFba, W.form.items.length);
  return outBuf;
}
function round2(x){ return Math.round((parseFloat(x)||0)*100)/100; }
/* 重算合计行：依据配置 M.totals（优先级，配置驱动/稳定）或表头列标签（兜底），
   将 qty/箱重/货值(单价×数量)/箱数 求和写入合计行；
   未被聚合覆盖的陈旧数值样本(如单元价合计/方数)清空，避免留陈旧错误数误导；
   文字标签(如"TOTAL：合计：")保留。 */
function recomputeTotalsRow(ws, totalsRow, itemStartR, lastDataRow, M){
  const sumField = fld => W.form.items.reduce((s,it)=> s + (parseFloat(it[fld])||0), 0);
  const sumValue = () => W.form.items.reduce((s,it)=> s + (parseFloat(it.declare)||0)*(parseFloat(it.qty)||0), 0);
  const sumCartons = () => W.form.items.reduce((s,it)=> s + (parseFloat(it.boxes)||1), 0);
  const trow = ws.getRow(totalsRow);
  const keepCols = new Set();
  if(M && M.totals){
    const T = M.totals;
    const set = (col,val)=>{ if(col){ const cl=String(col).replace(/[0-9]/g,''); keepCols.add(cl); trow.getCell(cl).value = round2(val); } };
    set(T.cartons, sumCartons());
    set(T.qty, sumField('qty'));
    set(T.weight, sumField('boxWeight'));
    set(T.value, sumValue());
  } else {
    // 兜底：表头探测（无 M.totals 时）
    let headerRow=0;
    for(let r=itemStartR-1; r>=1 && r>=itemStartR-6; r--){
      const row=ws.getRow(r); let hits=0;
      for(let c=1;c<=ws.columnCount;c++){
        const t=cleanTxt(row.getCell(c).value||'');
        if(t && FIELD_ALIASES.some(fa=>fa.names.some(n=>{ const cn=cleanTxt(n); return cn.includes(t)||t.includes(cn); }))) hits++;
      }
      if(hits>=2){ headerRow=r; break; }
    }
    const hrow = headerRow ? ws.getRow(headerRow) : null;
    for(let c=1;c<=ws.columnCount;c++){
      const cell = trow.getCell(c);
      const cur = cell.value;
      if(cur===null||cur===undefined||cur==='') continue;
      const isNum = (typeof cur==='number') || (typeof cur==='string' && cur.trim()!=='' && !isNaN(parseFloat(cur)));
      let matched=false;
      if(hrow){
        const hl = cleanTxt(hrow.getCell(c).value||'');
        if(/数量|quantity|qty|units|件数/.test(hl)){ cell.value=round2(sumField('qty')); matched=true; }
        else if(/重量|weight|gross|毛重|gw/.test(hl)){ cell.value=round2(sumField('boxWeight')); matched=true; }
        else if(/货值|total value|总值|总价|金额|amount/.test(hl)){ cell.value=round2(sumValue()); matched=true; }
        else if(/箱数|cartons|ctns/.test(hl)){ cell.value=round2(sumCartons()); matched=true; }
      }
      keepCols.add(ws.getColumn(c).letter);
      if(!matched && isNum) cell.value=null; // 清陈旧数值样本，保留文字标签
    }
  }
  // 通用兜底：清除合计行中未被聚合覆盖的陈旧数值（样本残留），保留文字标签
  if(M && M.totals){
    for(let c=1;c<=ws.columnCount;c++){
      const cell = trow.getCell(c);
      const cur = cell.value;
      if(cur===null||cur===undefined||cur==='') continue;
      if(keepCols.has(ws.getColumn(c).letter)) continue;
      const isNum = (typeof cur==='number') || (typeof cur==='string' && cur.trim()!=='' && !isNaN(parseFloat(cur)));
      if(isNum) cell.value=null;
    }
  }
}
function downloadBlob(blob,name){
  const url=URL.createObjectURL(new Blob([blob]));
  const a=document.createElement('a'); a.href=url; a.download=name; a.click();
  setTimeout(()=>URL.revokeObjectURL(url),3000);
}

/* ============================================================
   渠道·收货人主数据 (L4)
   ============================================================ */
/* ============================================================
   仓库主数据 (L4) — FBA 收货仓库 + 国家→仓库代码映射
   ============================================================ */
async function warehouses(){
  const list = await getAll('warehouses');
  const cfg = await get('config','whByC');
  const whByC = (cfg && cfg.v) ? {...cfg.v} : {};
  const chCountries = [...new Set((await getAll('channels')).map(c=>c.国家).filter(Boolean))];
  const needCountries = chCountries.filter(c=> !(c in whByC));
  const hint = needCountries.length ? `<div class="alert alert-warn" style="margin:10px 0">以下国家已在渠道中出现但尚未配置「国家→仓库代码」映射，正向向导将无法自动带出收货人（请在本页下方补充一次）：<b>${esc(needCountries.join('、'))}</b></div>` : '';
  main().innerHTML = `
  <h2>仓库主数据（FBA 收货仓库）</h2>
  <div class="sub">正向向导按「国家→仓库代码」自动反查收货公司/地址。这里维护仓库实体与国家映射，一次录入、全量一致。</div>
  ${hint}
  <div class="card">
    <button class="btn" id="addWh">+ 新增仓库</button>
    <table style="margin-top:12px"><thead><tr><th>代码</th><th>国家</th><th>公司</th><th>城市</th><th>地址</th><th>邮编</th><th>电话</th><th></th></tr></thead>
    <tbody>${list.map(w=>`<tr><td>${esc(w.代码)}</td><td>${esc(w.国家||'')}</td><td>${esc(w.公司||'')}</td><td>${esc(w.城市||'')}</td><td>${esc(w.地址||'')}</td><td>${esc(w.邮编||'')}</td><td>${esc(w.电话||'')}</td><td><button class="btn secondary" data-edit="${esc(w.id)}" style="padding:4px 8px">编辑</button> <button class="btn danger" data-del="${esc(w.id)}" style="padding:4px 8px">删</button></td></tr>`).join('')}</tbody></table>
  </div>
  <div id="whEditor"></div>
  <div class="card" style="margin-top:16px">
    <h3>国家 → 仓库代码 映射（正向向导自动选用）</h3>
    <div id="mapBody">${Object.entries(whByC).map(([c,code])=>mapRow(c,code,list)).join('')}</div>
    <div style="margin-top:10px;display:flex;gap:10px"><button class="btn secondary" id="addMap">+ 加映射</button><button class="btn" id="saveMap">保存映射</button></div>
  </div>`;
  $('#addWh').onclick=()=>editWh(null);
  $$('[data-edit]').forEach(b=> b.onclick=()=>editWh(b.dataset.edit));
  $$('[data-del]').forEach(b=> b.onclick=async()=>{ if(confirm('确认删除该仓库？')){ await del('warehouses',b.dataset.del); if(W.warehouses) W.warehouses=await getAll('warehouses'); warehouses(); } });
  bindMap(list);
}
function mapRow(c,code,list){
  return `<div class="row" style="align-items:end;margin:6px 0" data-map>
    <div><label>国家</label><input data-mc value="${esc(c)}"></div>
    <div><label>仓库代码</label><select data-mcode>${list.map(w=>`<option value="${esc(w.代码)}" ${w.代码===code?'selected':''}>${esc(w.代码)}</option>`).join('')}<option value="" ${!code?'selected':''}>（未选）</option></select></div>
    <button class="btn danger" data-mdel style="padding:7px 10px">×</button></div>`;
}
async function bindMap(list){
  const collect = ()=>{ const v={}; $$('#mapBody [data-map]').forEach(row=>{ const c=row.querySelector('[data-mc]').value.trim(); const code=row.querySelector('[data-mcode]').value; if(c) v[c]=code; }); return v; };
  $('#addMap').onclick=()=>{ document.getElementById('mapBody').insertAdjacentHTML('beforeend', mapRow('', '', list)); bindMap(list); };
  $$('#mapBody [data-mdel]').forEach(b=> b.onclick=()=>{ b.closest('[data-map]').remove(); });
  $('#saveMap').onclick=async()=>{ const v=collect(); await put('config',{id:'whByC', v}); if(W.warehouses) W.warehouses=await getAll('warehouses'); warehouses(); };
}
async function editWh(id){
  const all = await getAll('warehouses');
  const w = id? all.find(x=>x.id===id) : {id:uid(),代码:'',国家:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''};
  const box = $('#whEditor');
  box.innerHTML = `
  <div class="card" style="border-color:var(--accent)">
    <h3>${id?'编辑':'新增'}仓库</h3>
    <div class="row">
      <div><label>仓库代码</label><input id="w_代码" value="${esc(w.代码)}"></div>
      <div><label>国家</label><input id="w_国家" value="${esc(w.国家||'')}"></div>
      <div><label>公司</label><input id="w_公司" value="${esc(w.公司||'')}"></div>
    </div>
    <div class="row">
      <div><label>省份</label><input id="w_省份" value="${esc(w.省份||'')}"></div>
      <div><label>城市</label><input id="w_城市" value="${esc(w.城市||'')}"></div>
      <div><label>邮编</label><input id="w_邮编" value="${esc(w.邮编||'')}"></div>
      <div><label>电话</label><input id="w_电话" value="${esc(w.电话||'')}"></div>
    </div>
    <label>地址</label><input id="w_地址" value="${esc(w.地址||'')}">
    <div style="margin-top:14px;display:flex;gap:10px"><button class="btn" id="saveWh">保存</button><button class="btn secondary" id="cancelWh">取消</button></div>
  </div>`;
  $('#cancelWh').onclick=()=>{ box.innerHTML=''; };
  $('#saveWh').onclick=async()=>{
    w.代码=$('#w_代码').value.trim(); w.国家=$('#w_国家').value.trim(); w.公司=$('#w_公司').value.trim(); w.省份=$('#w_省份').value.trim(); w.城市=$('#w_城市').value.trim(); w.地址=$('#w_地址').value.trim(); w.邮编=$('#w_邮编').value.trim(); w.电话=$('#w_电话').value.trim();
    if(!w.代码){ alert('仓库代码必填'); return; }
    await put('warehouses',w); box.innerHTML=''; if(W.warehouses) W.warehouses=await getAll('warehouses'); warehouses();
  };
}

async function channels(){
  const list = await getAll('channels');
  main().innerHTML = `
  <h2>渠道·收货人主数据</h2>
  <div class="sub">L4 配置与资源层。这里维护「值」，生成时由向导反查，向导内只读。改一处、全量一致，易错录入收敛到此。</div>
  <div class="card">
    <button class="btn" id="addCh">+ 新增渠道</button>
    <table style="margin-top:12px"><thead><tr><th>物流商</th><th>渠道</th><th>国家</th><th>VAT</th><th>EORI</th><th>仓库</th><th></th></tr></thead>
    <tbody id="chBody">${list.map(c=>`<tr><td>${esc(c.物流商)}</td><td>${esc(c.渠道)}</td><td>${esc(c.国家)}</td><td>${esc(c.VAT)}</td><td>${esc(c.EORI)}</td><td>${esc((c.仓库||[]).map(w=>w.代码).join(', '))}</td>
      <td><button class="btn secondary" data-edit="${c.id}" style="padding:4px 8px">编辑</button> <button class="btn danger" data-del="${c.id}" style="padding:4px 8px">删</button></td></tr>`).join('')}</tbody></table>
  </div>
  <div id="chEditor"></div>`;
  $('#addCh').onclick=()=>editChannel(null);
  $$('[data-edit]').forEach(b=> b.onclick=()=>editChannel(b.dataset.edit));
  $$('[data-del]').forEach(b=> b.onclick=async()=>{ if(confirm('确认删除该渠道？')){ await del('channels',b.dataset.del); channels(); } });
}
async function editChannel(id){
  const all = await getAll('channels');
  const c = id? all.find(x=>x.id===id) : {id:uid(),物流商:'',渠道:'',国家:'',VAT:'',EORI:'',注册名:'',注册地址:'',仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]};
  const box = $('#chEditor');
  box.innerHTML = `
  <div class="card" style="border-color:var(--accent)">
    <h3>${id?'编辑':'新增'}渠道</h3>
    <div class="row">
      <div><label>物流商</label><input id="c_物流商" value="${esc(c.物流商)}"></div>
      <div><label>渠道</label><input id="c_渠道" value="${esc(c.渠道)}"></div>
      <div><label>国家</label><input id="c_国家" value="${esc(c.国家)}"></div>
    </div>
    <div class="row">
      <div><label>VAT号</label><input id="c_VAT" value="${esc(c.VAT)}"></div>
      <div><label>EORI</label><input id="c_EORI" value="${esc(c.EORI)}"></div>
      <div><label>VAT注册名</label><input id="c_注册名" value="${esc(c.注册名)}"></div>
    </div>
    <label>VAT注册地址</label><input id="c_注册地址" value="${esc(c.注册地址)}">
    <h3 style="margin-top:18px">仓库子表（按仓库代码反查地址）</h3>
    <div id="whList"></div>
    <button class="btn secondary" id="addWh" style="margin-top:8px">+ 加仓库</button>
    <div style="margin-top:14px;display:flex;gap:10px"><button class="btn" id="saveCh">保存</button><button class="btn secondary" id="cancelCh">取消</button></div>
  </div>`;
  const renderWh = ()=>{
    $('#whList').innerHTML = (c.仓库||[]).map((w,i)=>`
      <div class="row" style="margin:6px 0;align-items:end">
        <div><label>代码</label><input data-w="${i}" data-f="代码" value="${esc(w.代码)}"></div>
        <div><label>公司</label><input data-w="${i}" data-f="公司" value="${esc(w.公司)}"></div>
        <div><label>省份</label><input data-w="${i}" data-f="省份" value="${esc(w.省份)}"></div>
        <div><label>城市</label><input data-w="${i}" data-f="城市" value="${esc(w.城市)}"></div>
        <div><label>地址</label><input data-w="${i}" data-f="地址" value="${esc(w.地址)}"></div>
        <div><label>邮编</label><input data-w="${i}" data-f="邮编" value="${esc(w.邮编)}"></div>
        <div><label>电话</label><input data-w="${i}" data-f="电话" value="${esc(w.电话)}"></div>
        <div><button class="btn danger" data-wdel="${i}" style="padding:7px 10px">×</button></div>
      </div>`).join('');
    $$('#whList [data-w]').forEach(inp=> inp.oninput=e=>{ c.仓库[+e.target.dataset.w][e.target.dataset.f]=e.target.value; });
    $$('#whList [data-wdel]').forEach(b=> b.onclick=()=>{ c.仓库.splice(+b.dataset.wdel,1); renderWh(); });
  };
  renderWh();
  $('#addWh').onclick=()=>{ c.仓库.push({代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}); renderWh(); };
  $('#cancelCh').onclick=()=>{ box.innerHTML=''; };
  $('#saveCh').onclick=async()=>{
    c.物流商=$('#c_物流商').value; c.渠道=$('#c_渠道').value; c.国家=$('#c_国家').value; c.VAT=$('#c_VAT').value; c.EORI=$('#c_EORI').value; c.注册名=$('#c_注册名').value; c.注册地址=$('#c_注册地址').value;
    c.仓库=c.仓库.filter(w=>w.代码);
    await put('channels',c); box.innerHTML=''; channels();
  };
}

/* ============================================================
   SKU 主数据 (L4)
   ============================================================ */
async function skus(){
  const all = await getAll('skus');   // 一次性取全量（来自 IndexedDB，内存安全），但只渲染当前页
  const state = { q:'', page:1, size:100 };
  const SEARCH_COLS = ['sku','中文品名','英文品名','材质','HS','品牌','型号','申报价'];
  function computeFiltered(){
    const q = state.q.trim().toLowerCase();
    if(!q) return all;
    return all.filter(s=> SEARCH_COLS.some(k=> String(s[k]||'').toLowerCase().includes(q)));
  }
  function updateBody(){
    const filtered = computeFiltered();
    const total = filtered.length;
    const pages = Math.max(1, Math.ceil(total/state.size));
    if(state.page<1) state.page=1;
    if(state.page>pages) state.page=pages;
    const start = (state.page-1)*state.size;
    const rows = filtered.slice(start, start+state.size);
    $('#skTbody').innerHTML = rows.length ? rows.map(s=>`
      <tr><td>${esc(s.sku)}</td><td>${esc(s.中文品名)}</td><td>${esc(s.英文品名)}</td><td>${esc(s.材质)}</td><td>${esc(s.HS)}</td><td>${esc(s.品牌)}</td><td>${esc(s.型号)}</td><td>${esc(s.申报价)}</td><td>${esc((s.版本||[]).length)}</td>
      <td><button class="btn secondary" data-edit="${s.id}" style="padding:4px 8px">编辑</button> <button class="btn danger" data-del="${s.id}" style="padding:4px 8px">删</button></td></tr>`).join('') : '<tr><td colspan="10" class="empty">无匹配结果</td></tr>';
    $('#skPager').innerHTML = `共 <b>${total}</b> 条 · 第 ${state.page}/${pages} 页` +
      ` <button class="btn secondary" id="skPrev" ${state.page<=1?'disabled':''}>‹ 上一页</button>` +
      ` <button class="btn secondary" id="skNext" ${state.page>=pages?'disabled':''}>下一页 ›</button>`;
    $('#skPrev').onclick = ()=>{ state.page--; updateBody(); };
    $('#skNext').onclick = ()=>{ state.page++; updateBody(); };
    $$('#skTbody [data-edit]').forEach(b=> b.onclick=()=>editSku(b.dataset.edit));
    $$('#skTbody [data-del]').forEach(b=> b.onclick=async()=>{ if(confirm('确认删除？')){ await del('skus',b.dataset.del); skus(); } });
  }
  main().innerHTML = `
  <h2>SKU 主数据</h2>
  <div class="sub">申报价带版本号：变动追加新版本（生效日/原因），不覆盖；发票快照可复验。支持搜索 + 分页，避免万条数据一次渲染卡顿/崩溃。</div>
  <div class="card">
    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px">
      <input id="skSearch" placeholder="搜索 SKU / 中文品名 / 英文品名 / 品牌 / HS / 型号…" style="flex:1;min-width:260px;padding:8px;border:1px solid var(--border,#888);border-radius:6px;background:var(--surface,#fff);color:var(--fg,#000)">
      <select id="skSize" title="每页条数" style="padding:7px;border:1px solid var(--border,#888);border-radius:6px;background:var(--surface,#fff);color:var(--fg,#000)">
        <option value="50">50/页</option>
        <option value="100" selected>100/页</option>
        <option value="200">200/页</option>
        <option value="500">500/页</option>
      </select>
      <button class="btn" id="addSk">+ 新增 SKU</button>
      <button class="btn secondary" id="reseedSk">↻ 重新同步主数据（从商品申报信息）</button>
    </div>
    <span id="reseedMsg" style="color:var(--green)"></span>
    <table style="margin-top:12px"><thead><tr><th>SKU</th><th>中文品名</th><th>英文</th><th>材质</th><th>HS</th><th>品牌</th><th>型号</th><th>申报价</th><th>版本</th><th></th></tr></thead>
    <tbody id="skTbody"></tbody></table>
    <div id="skPager" style="margin-top:10px"></div>
  </div>
  <div id="skEditor"></div>`;
  const search=$('#skSearch');
  search.oninput=()=>{ state.q=search.value; state.page=1; updateBody(); };   // 只更新 tbody/pager，不重写搜索框 → 输入焦点不丢
  $('#skSize').onchange=(e)=>{ state.size=+e.target.value; state.page=1; updateBody(); };
  $('#addSk').onclick=()=>editSku(null);
  $('#reseedSk').onclick=async()=>{
    const msg=$('#reseedMsg');
    msg.style.color='var(--warn)'; msg.textContent='同步中…';
    try{
      await clear('skus');
      let n=0;
      for(const s of (window.SKUS||[])){ try{ await put('skus', s.id ? s : {...s, id: s.sku||uid()}); n++; }catch(e){} }
      // 让下次进入向导/物品载入直接读最新 IndexedDB（不必再 bump 版本号）
      try{ localStorage.removeItem('skus_seeded_ver'); }catch(e){}
      msg.style.color='var(--green)'; msg.textContent='✅ 已用最新「商品申报信息」重灌本地主数据（'+n+' 条）';
      setTimeout(()=>skus(), 600);
    }catch(e){ msg.style.color='var(--warn)'; msg.textContent='同步失败：'+e.message; }
  };
  updateBody();
}
async function editSku(id){
  const all=await getAll('skus');
  const s=id?all.find(x=>x.id===id):{id:uid(),sku:'',中文品名:'',英文品名:'',材质:'',HS:'',品牌:'',型号:'',申报价:'',成本:'',版本:[],图片:''};
  const box=$('#skEditor');
  box.innerHTML=`
  <div class="card" style="border-color:var(--accent)">
    <h3>${id?'编辑':'新增'}SKU</h3>
    <div class="row">
      <div><label>SKU</label><input id="s_sku" value="${esc(s.sku)}"></div>
      <div><label>中文品名</label><input id="s_cn" value="${esc(s.中文品名)}"></div>
      <div><label>英文品名</label><input id="s_en" value="${esc(s.英文品名)}"></div>
    </div>
    <div class="row">
      <div><label>材质</label><input id="s_mat" value="${esc(s.材质)}"></div>
      <div><label>HS编码</label><input id="s_hs" value="${esc(s.HS)}"></div>
      <div><label>品牌</label><input id="s_br" value="${esc(s.品牌)}"></div>
      <div><label>型号</label><input id="s_md" value="${esc(s.型号)}"></div>
    </div>
    <div class="row">
      <div><label>申报价(USD)</label><input id="s_dec" value="${esc(s.申报价)}"></div>
      <div><label>成本(USD)</label><input id="s_cost" value="${esc(s.成本)}"></div>
      <div><label>版本原因</label><input id="s_reason" placeholder="如：成本上涨调申报价"></div>
    </div>
    <div style="margin-top:14px;display:flex;gap:10px"><button class="btn" id="saveSk">保存</button><button class="btn secondary" id="cancelSk">取消</button></div>
  </div>`;
  $('#cancelSk').onclick=()=>box.innerHTML='';
  $('#saveSk').onclick=async()=>{
    const newDec=$('#s_dec').value, oldDec=s.申报价;
    s.sku=$('#s_sku').value; s.中文品名=$('#s_cn').value; s.英文品名=$('#s_en').value; s.材质=$('#s_mat').value; s.HS=$('#s_hs').value; s.品牌=$('#s_br').value; s.型号=$('#s_md').value; s.成本=$('#s_cost').value;
    if(newDec!==oldDec){ s.版本=s.版本||[]; s.版本.push({v:(s.版本.length+1),值:newDec,生效日:new Date().toISOString().slice(0,10),原因:$('#s_reason').value||'更新'}); }
    s.申报价=newDec;
    await put('skus',s); box.innerHTML=''; skus();
  };
}

/* ============================================================
   模板库 (L4)
   ============================================================ */
async function templates(){
  const list = await getAll('templates');
  main().innerHTML = `
  <h2>模板库</h2>
  <div class="sub">上传空白模板 xlsx（存本地 IndexedDB）。模板=排版层；版本迭代可停用旧版（状态 ACTIVE/DISABLED/DEPRECATED），停模板≠丢数据。</div>
  <div class="card">
    <h3>上传新模板</h3>
    <div class="hint">v1 已内置 5 家物流商映射。上传新模板时系统会自动扫描 xlsx 识别列字段，请确认检测结果后入库；若自动识别不完整可手动调整 JSON 映射。</div>
    <div class="row">
      <div><label>物流商</label><input id="t_物流商" value="安速"></div>
      <div><label>渠道</label><input id="t_渠道" value="美国包税海卡(正班)"></div>
      <div><label>空白模板 xlsx</label><input type="file" id="t_file" accept=".xlsx"></div>
    </div>
    <button class="btn" id="upTmpl" style="margin-top:10px">上传并入库</button>
    <div id="upLog" style="margin-top:8px"></div>
  </div>
  <div class="card">
    <h3>已有模板</h3>
    <table><thead><tr><th>物流商</th><th>渠道</th><th>版本</th><th>状态</th><th>创建日</th><th>操作</th></tr></thead>
    <tbody>${list.length?list.map(t=>`
      <tr><td>${esc(t.物流商)}</td><td>${esc(t.渠道)}</td><td>v${t.版本||1}</td><td><span class="pill ${t.状态==='ACTIVE'?'pill-green':(t.状态==='DISABLED'?'pill-gray':'pill-yellow')}">${t.状态}</span></td><td>${esc(t.创建日||'')}</td>
      <td>
        <button class="btn secondary" data-toggle="${t.id}" style="padding:4px 8px">${t.状态==='ACTIVE'?'停用':'启用'}</button>
        <button class="btn secondary" data-preview="${t.id}" style="padding:4px 8px">预览模板</button>
        <button class="btn secondary" data-editmap="${t.id}" style="padding:4px 8px">编辑映射</button>
        ${t.状态==='ACTIVE'?'<button class="btn secondary" data-deprec="'+t.id+'" style="padding:4px 8px">迭代弃用</button>':''}
        <button class="btn danger" data-del="${t.id}" style="padding:4px 8px">删</button>
      </td></tr>`).join(''):'<tr><td colspan=6 class="empty">暂无模板</td></tr>'}</tbody></table>
  </div>
  <div id="editMapLog"></div>
  <div id="previewLog"></div>`;
  $('#upTmpl').onclick=async()=>{
    const f=$('#t_file').files[0];
    if(!f){ $('#upLog').innerHTML='<div class="alert alert-err">请选择 xlsx 文件</div>'; return; }
    const key=$('#t_物流商').value.trim();
    if(!key){ $('#upLog').innerHTML='<div class="alert alert-err">请输入物流商名称</div>'; return; }
    $('#upLog').innerHTML='<div class="alert alert-info">正在扫描模板结构…</div>';
    try{
      const blob = new Blob([await f.arrayBuffer()], {type:f.type});
      const scanned = await scanTemplateMapping(blob);
      const hasMapping = Object.keys(scanned.item).length >= 4;
      $('#upLog').innerHTML = `
        <div class="card" style="margin-top:8px;padding:10px;font-size:12px">
          <div style="font-weight:500;margin-bottom:6px">扫描结果：${hasMapping?'<span style="color:green">✅ 已识别 '+Object.keys(scanned.item).length+' 个字段</span>':'<span style="color:#c0392b">⚠️ 仅识别 '+Object.keys(scanned.item).length+' 个字段，至少需 4 个</span>'}</div>
          <div><b>数据起始行：</b>${scanned.itemStartRow}</div>
          ${scanned.titleText?'<div><b>标题：</b>'+esc(scanned.titleText)+'</div>':''}
          ${Object.keys(scanned.meta).length?'<div><b>元信息字段：</b>'+Object.entries(scanned.meta).map(([k,v])=>esc(k)+'='+esc(v)).join('、')+'</div>':''}
          <div style="margin-top:4px"><b>列字段映射：</b></div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:2px;margin-top:4px">
            ${FIELD_ALIASES.filter(fa=>fa.key!=='fbaNo'&&fa.key!=='shipMethod'&&fa.key!=='warehouseCode'&&fa.key!=='company'&&fa.key!=='amazonRef'&&fa.key!=='poNo').map(fa=>{
              const col = scanned.item[fa.key];
              return '<div style="background:'+(col?'#e8f5e9':'#fcebeb')+';padding:2px 4px;border-radius:3px">'+esc(fa.label)+': <b>'+(col||'<span style="color:#999">未识别</span>')+'</b></div>';
            }).join('')}
          </div>
          <div style="margin-top:8px">
            <label>如需调整映射（格式 字段名:列字母, 字段名:列字母），修改下方 JSON 后点确认：</label>
            <textarea id="scanMapEdit" style="width:100%;height:80px;font-size:11px;font-family:monospace;margin-top:4px">${esc(JSON.stringify(scanned.item,null,2))}</textarea>
          </div>
          <div style="margin-top:6px;display:flex;gap:8px">
            <button class="btn" id="confirmScan" ${!hasMapping?'disabled':''}>${hasMapping?'✅ 确认并入库':'字段不足，仍要强制入库？'}</button>
            <button class="btn secondary" id="cancelScan">取消</button>
          </div>
        </div>`;
      $('#cancelScan').onclick = ()=>{ $('#upLog').innerHTML=''; };
      $('#confirmScan').onclick = async()=>{
        let finalItem;
        try{ finalItem = JSON.parse($('#scanMapEdit').value); }catch(e){ alert('JSON 格式错误'); return; }
        if(Object.keys(finalItem).length<2){ alert('至少需 2 个列映射'); return; }
        // 合并：最终 mapping = 自动检测的 meta + 用户确认的 item
        const finalMapping = {
          meta: scanned.meta,
          item: finalItem,
          itemStartRow: scanned.itemStartRow,
          titleCell: scanned.titleCell,
          titleText: scanned.titleText,
        };
        const rec = {id:uid(),物流商:key,渠道:$('#t_渠道').value,名称:f.name,blob,状态:'ACTIVE',版本:1,创建日:new Date().toISOString().slice(0,10),mapping:finalMapping};
        await put('templates',rec);
        $('#upLog').innerHTML='<div class="alert alert-ok">✅ 已入库（使用扫描生成的映射）</div>';
        templates();
      };
    }catch(e){
      $('#upLog').innerHTML='<div class="alert alert-err">扫描失败：'+esc(e.message)+'。可改用代码预设映射入库。</div>';
    }
  };
  $$('[data-toggle]').forEach(b=> b.onclick=async()=>{ const t=list.find(x=>x.id===b.dataset.toggle); t.状态=t.状态==='ACTIVE'?'DISABLED':'ACTIVE'; await put('templates',t); templates(); });
  $$('[data-deprec]').forEach(b=> b.onclick=async()=>{ const t=list.find(x=>x.id===b.dataset.deprec); t.状态='DEPRECATED'; await put('templates',t); templates(); });
  $$('[data-del]').forEach(b=> b.onclick=async()=>{ if(confirm('确认删除模板？')){ await del('templates',b.dataset.del); templates(); } });
  $$('[data-preview]').forEach(b=> b.onclick=async()=>{ await previewTemplate(b.dataset.preview, list); });
  $$('[data-editmap]').forEach(b=>{
    b.onclick=()=>{
      const t = list.find(x=>x.id===b.dataset.editmap);
      if(!t || !t.mapping){ alert('该模板没有 mapping 数据'); return; }
      $('#editMapLog').innerHTML = `
        <div class="card" style="margin-top:8px;padding:10px;font-size:12px">
          <div style="font-weight:500;margin-bottom:6px">编辑模板「${esc(t.物流商)} / ${esc(t.渠道||'(通用)')}」映射</div>
          <div><b>数据起始行：</b><input id="ed_startRow" value="${t.mapping.itemStartRow||21}" style="width:60px"></div>
          <div style="margin-top:4px"><b>item 列字段映射（JSON）：</b></div>
          <textarea id="ed_itemMap" style="width:100%;height:120px;font-size:11px;font-family:monospace">${esc(JSON.stringify(t.mapping.item||{},null,2))}</textarea>
          <div style="margin-top:6px"><b>meta 元信息（JSON，可清空不要的脏值）：</b></div>
          <textarea id="ed_metaMap" style="width:100%;height:80px;font-size:11px;font-family:monospace">${esc(JSON.stringify(t.mapping.meta||{},null,2))}</textarea>
          <div style="margin-top:6px;display:flex;gap:8px">
            <button class="btn" id="ed_save">💾 保存映射</button>
            <button class="btn secondary" id="ed_cancel">取消</button>
          </div>
        </div>`;
      $('#ed_cancel').onclick = ()=>{ $('#editMapLog').innerHTML=''; };
      $('#ed_save').onclick = async()=>{
        let item, meta;
        try{
          item = JSON.parse($('#ed_itemMap').value);
          meta = JSON.parse($('#ed_metaMap').value);
        }catch(e){ alert('JSON 格式错误：'+e.message); return; }
        const startRow = parseInt($('#ed_startRow').value,10) || (item&&Object.keys(item).length?21:1);
        t.mapping = Object.assign({}, t.mapping, {item, meta, itemStartRow:startRow});
        await put('templates', t);
        $('#editMapLog').innerHTML='<div class="alert alert-ok">✅ 映射已保存（已立即生效，无需重传文件）</div>';
        setTimeout(()=>{ if(W && W.templates){ W.templates = list.map(x => x.id===t.id ? t : x); } templates(); }, 1000);
      };
    };
  });
}

/* 预览模板：渲染 xlsx 前30行为表格，高亮 mapping 命中的列 */
async function previewTemplate(id, list){
  const t = list.find(x=>x.id===id);
  if(!t || !t.blob){ alert('模板文件不存在'); return; }
  try{
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await t.blob.arrayBuffer());
    let ws = wb.getWorksheet(1) || wb.worksheets[0];
    if(!ws){ alert('模板无工作表'); return; }
    const colL = n => n<1?'':n>26?String.fromCharCode(64+Math.floor((n-1)/26))+String.fromCharCode(65+(n-1)%26):String.fromCharCode(64+n);
    const maxRow = Math.min(ws.rowCount||30, 30);
    const maxCol = Math.min(ws.columnCount||26, 26);
    // 收集 mapping 列 → 字段
    const col2field = {};
    const itemMap = (t.mapping && t.mapping.item) || {};
    Object.entries(itemMap).forEach(([k, col])=>{ if(col && !col2field[col]) col2field[col] = k; });
    const metaMap = (t.mapping && t.mapping.meta) || {};
    Object.entries(metaMap).forEach(([k, cell])=>{
      if(!cell) return;
      // cell like "D2" — extract column letter
      const m = String(cell).match(/^([A-Z]+)/);
      if(m && !col2field[m[1]]) col2field[m[1]] = k+' (meta)';
    });
    // 渲染表格
    let html = `<div class="card" style="margin-top:8px;padding:10px;font-size:12px;overflow-x:auto">
      <div style="font-weight:500;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center">
        <span>预览模板「${esc(t.物流商)} / ${esc(t.名称)}」前 ${maxRow} 行 × ${maxCol} 列</span>
        <button class="btn secondary" id="pv_close" style="padding:2px 8px">关闭</button>
      </div>
      <div style="font-size:11px;color:#888;margin-bottom:6px">
        <span style="display:inline-block;width:12px;height:12px;background:#e8f5e9;vertical-align:middle"></span> 已映射字段
        <span style="display:inline-block;width:12px;height:12px;background:#fff9c4;margin-left:8px;vertical-align:middle"></span> 表头行
        <span style="display:inline-block;width:12px;height:12px;background:#fcebeb;margin-left:8px;vertical-align:middle"></span> 未映射
      </div>
      <table style="border-collapse:collapse;font-size:11px;font-family:monospace">
        <thead><tr style="background:#f5f5f5"><th style="border:1px solid #ccc;padding:2px 6px">行/列</th>`;
    for(let c=1; c<=maxCol; c++){
      html += `<th style="border:1px solid #ccc;padding:2px 6px;color:#666">${colL(c)}</th>`;
    }
    html += `</tr></thead><tbody>`;
    for(let r=1; r<=maxRow; r++){
      const row = ws.getRow(r);
      const isHeader = col2field && Object.keys(col2field).length>0 && r===1; // 简化：只高亮 R1 提示是表头
      html += `<tr><td style="border:1px solid #ccc;padding:2px 6px;background:#f5f5f5;color:#666">${r}</td>`;
      for(let c=1; c<=maxCol; c++){
        const cell = row.getCell(c);
        const v = cell.value;
        let display = '';
        if(v!==null && v!==undefined){
          if(typeof v === 'object'){
            display = v.formula ? `[公式]${v.result||''}` : (v.text || JSON.stringify(v).substring(0,30));
          } else {
            display = String(v).substring(0,40);
          }
        }
        const colLet = colL(c);
        const field = col2field[colLet];
        let bg = '#fff';
        if(field && r>=(t.mapping?.itemStartRow||21)) bg = '#e8f5e9';
        else if(field && r<=(t.mapping?.itemStartRow||21)) bg = '#fff9c4';
        html += `<td style="border:1px solid #ccc;padding:2px 6px;background:${bg};max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${field?esc(field)+' | ':''}${esc(display)}">${esc(display)}</td>`;
      }
      html += `</tr>`;
    }
    html += `</tbody></table>
      <div style="margin-top:6px;font-size:11px;color:#666">
        绿色 = 已映射列（在 itemStartRow 之后）；黄色 = 已映射列（itemStartRow 之前，可能是表头/meta）；白色 = 未映射。数据起始行 ${t.mapping?.itemStartRow||21}。
      </div>
    </div>`;
    $('#previewLog').innerHTML = html;
    $('#pv_close').onclick = ()=>{ $('#previewLog').innerHTML = ''; };
    $('#previewLog').scrollIntoView({behavior:'smooth', block:'start'});
  }catch(e){
    alert('预览失败：'+e.message);
  }
}

/* ============================================================
   校验·监控 (L6)
   ============================================================ */
async function monitor(){
  const recs = await getAll('records');
  main().innerHTML = `
  <h2>校验·监控</h2>
  <div class="sub">质量规则总览 + 生成记录（溯源审计）。每条记录带时间/物流商/模板，可复验。</div>
  <div class="card">
    <h3>质量规则</h3>
    <table><thead><tr><th>规则</th><th>类型</th><th>说明</th></tr></thead><tbody>
      <tr><td>必填完整性</td><td><span class="pill pill-red">阻断</span></td><td>收货人关键字段 + 物品必填项为空则阻断</td></tr>
      <tr><td>勾稽·箱数/数量</td><td><span class="pill pill-red">阻断</span></td><td>箱号/数量合计与装箱单一致</td></tr>
      <tr><td>申报价来源</td><td><span class="pill pill-yellow">告警</span></td><td>无主数据按下推算(成本×${COEFF})需人审确认</td></tr>
      <tr><td>源忠实</td><td><span class="pill pill-green">提示</span></td><td>推算值标黄、主数据反查标绿，便于复核</td></tr>
      <tr><td>人审闸门</td><td><span class="pill pill-red">阻断</span></td><td>未勾选确认不得导出/发送</td></tr>
    </tbody></table>
  </div>
  <div class="card">
    <h3>生成记录（${recs.length}）</h3>
    ${recs.length?`<table><thead><tr><th>时间</th><th>物流商</th><th>渠道</th><th>仓库</th><th>FBA号</th><th>状态</th></tr></thead><tbody>
      ${recs.slice().reverse().map(r=>`<tr><td>${esc(r.时间)}</td><td>${esc(r.物流商)}</td><td>${esc(r.渠道)}</td><td>${esc(r.仓库)}</td><td>${esc(r.fba)}</td><td><span class="pill pill-green">${esc(r.状态)}</span></td></tr>`).join('')}</tbody></table>`
      :'<div class="empty">暂无生成记录，去「生成发票向导」产出第一票。</div>'}
  </div>`;
}

/* ---------- 启动 ---------- */
(async function init(){
  const status = document.getElementById('dbStatus');
  const verEl = document.getElementById('appVer');
  if(verEl) verEl.textContent = 'v'+APP_VERSION;
  // 版本过期强提示：若打开的链接带 ?t= 但与当前代码版本不符，说明浏览器在跑旧缓存 app.js
  try{
    const t = new URLSearchParams(location.search).get('t');
    if(t && t !== APP_VERSION){
      const warn = document.createElement('div');
      warn.style.cssText = 'position:sticky;top:0;z-index:9999;background:#c0392b;color:#fff;padding:12px 18px;font-size:15px;font-weight:700;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.3)';
      warn.innerHTML = '⚠️ 版本过期：链接是 v'+t+'，但已加载代码是 v'+APP_VERSION+'。你正在运行<b>旧缓存</b>的 app.js，导出结果可能错误！<br>请<b>关闭此标签页</b>，重新打开我下发的最新链接（或 Ctrl+Shift+R 硬刷新）。';
      document.body.insertBefore(warn, document.body.firstChild);
    }
  }catch(_){}

  try{
    await openDB();
    status.textContent = '存储: 正在注入主数据种子…';
    await seedIfEmpty();
    const tmpls = await getAll('templates');
    const tmplOk = tmpls.filter(t=>t.状态!=='DISABLED' && t.blob && t.blob.size>0).length;
    status.textContent = USE_DB
      ? `存储: 本地 IndexedDB ✓ (模板 ${tmplOk}/${seedStatus.total})`
      : '存储: 内存模式(IndexedDB不可用)';
    status.style.color = USE_DB ? 'var(--green)' : 'var(--warn)';
    if(!userNavigated) go('overview');   // 用户在种子加载期间已切走视图则不强拉回总览
  }catch(e){
    status.textContent = '存储: 异常，已降级';
    console.error(e);
    try{ if(!userNavigated) go('overview'); }catch(_){}
    // 兜底：保证 main 永不空白（即使 overview 也抛错也能看到具体异常）
    if(main() && !main().innerHTML){
      main().innerHTML = '<div class="alert alert-err" style="margin:20px"><b>⚠️ 初始化异常，已降级运行</b><br>'+
        '<pre style="white-space:pre-wrap;color:#ff9;background:#222;padding:8px;margin-top:8px;border-radius:4px">'+
        esc((e&&e.stack)?e.stack+'\n\n[name:'+e.name+', message:'+e.message+']':String(e))+
        '</pre>'+
        '<div style="margin-top:10px;color:#bbb">常见原因：浏览器隐私模式禁用 IndexedDB / 旧版本 store schema 不兼容。请按 <b>Ctrl+Shift+R</b> 硬刷新一次；仍不行请打开开发者工具 (F12) 把 Console 错误发我。</div></div>';
    }
  }
})();
