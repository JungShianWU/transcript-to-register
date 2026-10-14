/* 謄本 → 地號清冊  核心程式（解析 + 產生 xlsx）
   純前端、不連網。瀏覽器與 Node 共用。 */
(function (root) {
  'use strict';

  // ───────────── 共用小工具 ─────────────
  const toHalf = s => String(s || '')
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[Ａ-Ｚａ-ｚ]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/－/g, '-').replace(/＊/g, '*').replace(/　/g, ' ');
  const num = s => { const m = String(s || '').replace(/[,，\s]/g, '').match(/\d+(?:\.\d+)?/); return m ? parseFloat(m[0]) : null; };
  const CN = { 零: 0, 〇: 0, 一: 1, 二: 2, 兩: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 };
  function cnNum(s) {
    s = toHalf(s);
    if (/^\d+$/.test(s)) return parseInt(s, 10);
    let total = 0, cur = 0, ok = false;
    for (const ch of s) {
      if (ch in CN) { cur = CN[ch]; ok = true; }
      else if (ch === '十') { total += (cur || 1) * 10; cur = 0; ok = true; }
      else if (ch === '百') { total += (cur || 1) * 100; cur = 0; ok = true; }
      else return null;
    }
    return ok ? total + cur : null;
  }
  const frac = s => { const m = String(s || '').match(/(\d+)\s*分之\s*(\d+)/); return m ? { den: +m[1], num: +m[2] } : null; };

  // ───────────── 1. 文字前處理 ─────────────
  function cleanText(t) {
    t = String(t || '').replace(/\r\n?/g, '\n');
    // 跨頁：「（續次頁）」到下一頁「頁次：N」之間是浮水印、驗證碼、頁首，整段移除
    t = t.replace(/（續次頁）[\s\S]*?頁次：\s*\d+[^\n]*\n/g, '\n');
    return t;
  }

  // ───────────── 2. 解析 ─────────────
  function parseHeader(seg) {
    const h = seg.match(/([^\s：:]+?(?:區|鄉|鎮|市))([^\s：:]+?段)([^\s：:]+?小段)?\s*(\d{4,5})-(\d{3,4})\s*(地號|建號)/);
    if (!h) return null;
    const city = (seg.match(/資料管轄機關：\s*(\S+?[市縣])/) || [])[1] || '';
    return { city, township: h[1], section: h[2], subsection: h[3] || '', main: h[4], sub: h[5], kind: h[6] };
  }

  function between(seg, startRe, endRe) {
    const s = seg.search(startRe); if (s < 0) return '';
    const rest = seg.slice(s);
    const e = rest.slice(1).search(endRe);
    return e < 0 ? rest : rest.slice(0, e + 1);
  }

  function parseOwners(sec) {
    if (!sec) return [];
    const parts = sec.split(/（\d{4}）登記次序/).slice(1);
    return parts.map(b => {
      const name = toHalf((b.match(/所有權人：\s*(\S+)/) || [])[1] || '');
      const id = toHalf((b.match(/統一編號：\s*([A-Za-z0-9*＊]+)/) || [])[1] || '');
      const rLine = (b.match(/(?<!取得|設定)權利範圍：([^\n]*)/) || [])[1] || '';
      let f = frac(rLine);
      if (!f && /全部/.test(rLine)) f = { den: 1, num: 1 };
      const order = +((b.match(/^[：:]\s*(\d+)/) || [])[1] || 0);
      const note = toHalf((b.match(/其他登記事項：([\s\S]*)$/) || [])[1] || '').replace(/\s+/g, '');
      return { name, id, share: f, shareText: rLine.replace(/[*＊]/g, '').trim(), order, note, joint: /公同共有/.test(rLine) };
    }).filter(o => o.name);
  }

  // 他項權利部：每筆（登記次序）一個區塊
  const lbl = t => [...t].map(c => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s*');
  // 同一行可能有兩個欄位（例：登記日期：…  登記原因：設定），取到下一個「標籤：」之前為止
  const fld = (b, label) => { const m = b.match(new RegExp(lbl(label) + '\\s*[：:]\\s*([^\\n]*?)(?=\\s+[^\\s：:()（）]{2,10}\\s*(?:\\([^)]*\\))?\\s*[：:]|\\n|$)')); return m ? toHalf(m[1]).replace(/[*＊]{2,}/g, ' ').replace(/\s+/g, ' ').trim() : ''; };
  const multi = (b, label, stops) => { const m = b.match(new RegExp(lbl(label) + '\\s*[：:]\\s*([\\s\\S]*?)(?=' + stops.map(lbl).join('|') + '|$)')); return m ? toHalf(m[1]).replace(/\s+/g, ' ').trim() : ''; };
  function parseOtherRights(sec) {
    if (!sec) return [];
    const parts = sec.split(/（\d{4}）登記次序/).slice(1);
    return parts.map(b => {
      const order = toHalf((b.match(/^[：:]\s*([\d-]+)/) || [])[1] || '');
      const STOP = ['共同擔保建號', '共同擔保地號', '其他登記事項', '證明書字號', '流抵約定', '權利標的'];
      const amountTxt = fld(b, '擔保債權總金額');
      return {
        order,
        type: (b.match(/權利種類：\s*(\S+)/) || [])[1] || '',
        year: fld(b, '收件年期'), no: (b.match(/字號：\s*(\S+)/) || [])[1] || '',
        regDate: fld(b, '登記日期'), reason: fld(b, '登記原因'),
        holder: fld(b, '權利人'), holderId: fld(b, '統一編號'),
        ratio: fld(b, '債權額比例'), amount: num(amountTxt), amountText: amountTxt,
        term: fld(b, '存續期間'), detDate: fld(b, '擔保債權確定期日').replace(/[。.]$/, ''),
        repay: fld(b, '清償日期').replace(/[。.]$/, ''),
        target: fld(b, '權利標的'), targetOrder: toHalf((b.match(/標的登記次序：\s*([\d\s-]+)/) || [])[1] || '').trim().split(/\s+/).filter(Boolean),
        scope: fld(b, '設定權利範圍'), cert: fld(b, '證明書字號'),
        debtor: multi(b, '債務人及債務額比例', ['權利標的', '標的登記次序', '設定權利範圍', '證明書字號', '擔保債權總金額']),
        obligor: fld(b, '設定義務人'),
        coLand: multi(b, '共同擔保地號', STOP.filter(x => x !== '共同擔保地號')),
        coBldg: multi(b, '共同擔保建號', STOP.filter(x => x !== '共同擔保建號')),
      };
    }).filter(x => x.type || x.holder);
  }

  function parseLand(seg, hd) {
    const mark = between(seg, /土地標示部/, /土地所有權部/);
    const own = between(seg, /土地所有權部/, /土地他項權利部|$/);
    const oth = between(seg, /土地他項權利部/, /$^/);
    const bl = between(mark, /地上建物建號：/, /其他登記事項/);
    return {
      type: 'land', ...hd,
      parcel: hd.main + '-' + hd.sub,
      area: num((mark.match(/面\s*積：[*＊\s]*([\d,.]+)\s*平方公尺/) || [])[1]),
      landValue: num((mark.match(/公告土地現值：[*＊\s]*([\d,.]+)\s*元/) || [])[1]),
      landValueYear: +((mark.match(/民國\s*(\d+)\s*年\s*\d+\s*月\s*公告土地現值/) || [])[1] || 0),
      zoningRaw: toHalf((mark.match(/使用分區：\s*(\S+)/) || [])[1] || ''),
      buildingsOnLand: (bl.match(/\d{5}-\d{3}/g) || []),
      owners: parseOwners(own),
      rights: parseOtherRights(oth),
    };
  }

  function parseBuilding(seg, hd) {
    const mark = between(seg, /建物標示部/, /建物所有權部|（無所有權資料）|$/);
    const own = between(seg, /建物所有權部/, /建物他項權利部|$/);
    const sitText = between(mark, /建物坐落地號：/, /主要用途/);
    const sitting = [];
    sitText.replace(/([^\s：:]+?段)([^\s：:]+?小段)?\s*(\d{4})-(\d{4})/g, (_, s, ss, m, n) => { sitting.push({ section: s, subsection: ss || '', parcel: m + '-' + n }); });
    const annex = [];
    mark.replace(/附屬建物用途：\s*(\S+)\s*面\s*積：[*＊\s]*([\d,.]+)/g, (_, u, a) => { annex.push({ usage: u, area: num(a) }); });
    const common = [];
    mark.replace(/共有部分：\s*\S*?(\d{5})-(\d{3})建號[*＊\s]*([\d,.]+)\s*平方公尺[\s\S]*?權利範圍：([^\n]*)/g,
      (_, m, n, a, r) => { common.push({ no: m + '-' + n, area: num(a), share: frac(r) }); });
    const mains = [];
    mark.replace(/主建物資料：\s*\S*?(\d{5})-(\d{3})建號[\s\S]*?權利範圍：([^\n]*)/g,
      (_, m, n, r) => { mains.push({ no: m + '-' + n, share: frac(r) }); });
    const usage = ((mark.match(/主要用途：\s*(\S+)/) || [])[1] || '');
    const owners = parseOwners(own);
    const rights = parseOtherRights(between(seg, /建物他項權利部/, /$^/));
    return {
      type: 'building', ...hd, rights,
      no: hd.main + '-' + hd.sub,
      door: toHalf((mark.match(/建物門牌：\s*([^\n]+)/) || [])[1] || '').trim(),
      sitting, usage,
      floors: num((mark.match(/層\s*數：\s*(\d+)\s*層/) || [])[1]),
      level: ((mark.match(/層\s*次：\s*(\S+)/) || [])[1] || ''),
      area: num((mark.match(/總面積：[*＊\s]*([\d,.]+)/) || [])[1]),
      annex, common, mains, owners,
      isCommon: /共有部分/.test(usage) || (/無所有權資料/.test(seg) && owners.length === 0),
    };
  }

  function parseText(rawText, fileName, meta) {
    const text = cleanText(rawText);
    const out = [], warn = [];
    const segs = text.split(/〈\s*本謄本列印完畢\s*〉/);
    for (let seg of segs) {
      const st = seg.search(/(?:土地|建物)登記第[一二三12]類謄本/);
      if (st < 0) continue;
      seg = seg.slice(st);
      const hd = parseHeader(seg);
      if (!hd) { warn.push(`${fileName}：有一份謄本讀不到地號／建號標題，已略過`); continue; }
      out.push(hd.kind === '地號' ? parseLand(seg, hd) : parseBuilding(seg, hd));
    }
    if (!out.length) {
      // 區分「沒有文字層（掃描檔）」與「有文字但不是謄本格式」
      const chars = (String(rawText || '').match(/[\u4e00-\u9fff0-9A-Za-z]/g) || []).length;
      const printed = meta && /print\s*to\s*pdf|microsoft:\s*print|pdfcreator|cutepdf|bullzip|ghostscript|quartz pdfcontext/i.test((meta.producer || '') + ' ' + (meta.creator || ''));
      if (chars < 50 && printed) warn.push(`${fileName}：這份 PDF 曾經過「列印成 PDF」轉存（檔案屬性：${meta.producer || meta.creator}），文字已變成圖形，無法辨識。轉存可能在取得檔案之前就已發生；請向來源索取從地政網路申領系統下載的原始檔，或重新申請電子謄本`);
      else if (chars < 50) warn.push(`${fileName}：這份 PDF 沒有文字（掃描檔或影像檔），無法辨識。請改用地政事務所網路申領的電子謄本`);
      else if (!/登記第[一二三12]類謄本|本謄本列印完畢/.test(text)) warn.push(`${fileName}：讀得到文字，但不是土地／建物登記謄本的格式，已略過`);
      else warn.push(`${fileName}：是電子謄本，但標題或格式與目前支援的不同，無法辨識。請回報這份檔案的格式`);
    }
    return { items: out, warn };
  }

  // ───────────── 3. 對應成清冊列 ─────────────
  const ownerKey = o => o.name + '|' + o.id;
  const landKey = (sec, sub, parcel) => sec + (sub || '') + parcel;
  const parcelDisp = p => { const [m, s] = p.split('-'); const mi = parseInt(m, 10), si = parseInt(s, 10); return si ? `${mi}-${si}` : mi; };
  const bldDisp = n => { const m = parseInt(n.split('-')[0], 10), s = parseInt(n.split('-')[1], 10); return s ? `${m}-${s}` : m; };

  function floorLabel(b) {
    let lv = b.level || '';
    let m;
    if ((m = lv.match(/^地下(\S+?)層/))) { const n = cnNum(m[1]); if (n) return 'B' + n; }
    if (/屋頂突出|屋突/.test(lv)) return 'R';
    if ((m = lv.match(/^(\S+?)層/))) { const n = cnNum(m[1]); if (n) return String(n); }
    const d = b.door;
    if ((m = d.match(/地下(\S+?)層/))) { const n = cnNum(m[1]); if (n) return 'B' + n; }
    if (/地下室/.test(d)) return 'B1';
    if ((m = d.match(/號(\S+?)樓/))) { const n = cnNum(m[1]); if (n) return String(n); }
    return '';
  }

  const ANNEX_COL = [
    [/騎樓/, 'W'], [/平台|平臺/, 'AG'], [/夾層/, 'AI'], [/屋頂突出|屋突/, 'AK'], [/防空|避難/, 'AM'],
    [/./, 'AB'], // 陽臺、露臺、雨遮、花台等其他附屬建物
  ];

  function buildModel(items) {
    const warn = [];
    const lands = items.filter(i => i.type === 'land');
    const blds = items.filter(i => i.type === 'building');
    // 去重（同一份謄本重複上傳）
    const uniq = (arr, k) => { const s = new Map(); arr.forEach(a => { const kk = k(a); if (!s.has(kk)) s.set(kk, a); }); return [...s.values()]; };
    const L = uniq(lands, l => landKey(l.section, l.subsection, l.parcel)).sort((a, b) => (a.section + a.subsection).localeCompare(b.section + b.subsection) || a.parcel.localeCompare(b.parcel));
    const B = uniq(blds, b => b.section + b.no);
    const mainB = B.filter(b => !b.isCommon);
    const commonB = B.filter(b => b.isCommon);

    // 建物 × 所有權人 → 歸到哪一筆土地
    const landByKey = new Map(L.map(l => [landKey(l.section, l.subsection, l.parcel), l]));
    const pairsByLand = new Map(L.map(l => [l, []]));
    for (const b of mainB) {
      const sitLands = b.sitting.map(s => landByKey.get(landKey(s.section, s.subsection, s.parcel))).filter(Boolean);
      if (!sitLands.length) { warn.push(`建號 ${b.no}（${b.door}）坐落的土地謄本沒有上傳，未列入清冊`); continue; }
      if (!b.owners.length) warn.push(`建號 ${b.no} 沒有所有權人資料`);
      for (const o of b.owners) {
        const target = sitLands.find(l => l.owners.some(lo => ownerKey(lo) === ownerKey(o))) || sitLands[0];
        pairsByLand.get(target).push({ b, o });
        if (sitLands.length > 1) warn.push(`建號 ${b.no} 坐落多筆土地，已列在 ${target.parcel} 地號下`);
      }
    }

    const blocks = [], info = [];
    // 公同共有：同一組共有人共同持有同一個權利範圍，只能算一次
    function jointGroups(owners) {
      const groups = new Map(); // key -> [owners]
      const keyOf = new Map();
      const ranges = owners.filter(o => o.joint).map(o => { const m = o.note.match(/登記次序(\d+)至(\d+)公同共有/); return m ? [+m[1], +m[2]] : null; }).filter(Boolean);
      for (const o of owners) {
        if (!o.joint) continue;
        const rg = ranges.find(([a, b]) => o.order >= a && o.order <= b);
        const k = rg ? `R${rg[0]}-${rg[1]}` : `F${o.share ? o.share.num + '/' + o.share.den : '?'}`;
        keyOf.set(o, k);
        if (!groups.has(k)) groups.set(k, []);
        groups.get(k).push(o);
      }
      return { groups, keyOf };
    }
    for (const l of L) {
      const rows = [];
      const jg = jointGroups(l.owners);
      const jointRanges = [];
      // 依謄本順序排列，同一組公同共有的人排在一起（放在該組第一位出現的位置）
      const ordered = [], seen = new Set();
      for (const o of l.owners) {
        if (seen.has(o)) continue;
        const k = jg.keyOf.get(o);
        const members = k ? jg.groups.get(k) : [o];
        members.forEach(m => seen.add(m));
        ordered.push(members);
      }
      const pairs = pairsByLand.get(l).sort((x, y) => x.b.no.localeCompare(y.b.no));
      const used = new Set();
      for (const members of ordered) {
        const isJoint = members.length > 1 || (members[0] && members[0].joint);
        const start = rows.length;
        members.forEach((o, mi) => {
          const landShare = isJoint && mi > 0 ? null : o.share;
          const mine = pairs.filter(p => ownerKey(p.o) === ownerKey(o));
          if (!mine.length) rows.push({ owner: o, land: landShare, b: null, joint: isJoint, jointCont: isJoint && mi > 0 });
          mine.forEach((p, i) => { used.add(p); rows.push({ owner: o, land: i === 0 ? landShare : null, b: p.b, bshare: p.o.share, joint: isJoint, jointCont: isJoint && (mi > 0 || i > 0) }); });
        });
        if (isJoint) {
          rows[start].jointSize = members.length;
          if (rows.length - start > 1) jointRanges.push([start, rows.length - 1]);
          const sh = members[0].share;
          info.push(`${l.parcel} 地號：${members.map(m => m.name).join('、')} 共 ${members.length} 人公同共有 ${sh ? sh.num + '/' + sh.den : ''}，持分只計一次，H～N 欄跨列合併`);
          const diff = members.filter(m => !m.share || !sh || m.share.num * sh.den !== sh.num * m.share.den);
          if (diff.length) warn.push(`${l.parcel} 地號：公同共有的 ${members.map(m => m.name).join('、')} 權利範圍不一致，請人工確認`);
        }
      }
      for (const p of pairs) if (!used.has(p)) {
        rows.push({ owner: p.o, land: null, b: p.b, bshare: p.o.share });
        warn.push(`${l.parcel} 地號：建號 ${p.b.no} 的所有權人 ${p.o.name} 不是土地所有權人，已列在區塊最後（土地持分空白）`);
      }
      if (!rows.length) rows.push({ owner: null, land: null, b: null });
      // 檢核：土地持分合計
      let s = 0; ordered.forEach(ms => { const o = ms[0]; if (o.share) s += o.share.num / o.share.den; });
      l.shareSum = s;
      if (Math.abs(s - 1) > 1e-9) warn.push(`${l.parcel} 地號：土地持分合計 = ${s.toFixed(6)}，不等於 1`);
      // 檢核：地上建物建號是否都有謄本
      const have = new Set(B.map(b => b.no));
      const miss = l.buildingsOnLand.filter(n => !have.has(n));
      if (miss.length) warn.push(`${l.parcel} 地號：地上建物 ${miss.join('、')} 沒有建物謄本`);
      blocks.push({ land: l, rows, jointRanges });
    }
    // 檢核：共有部分持分合計（以共有部分謄本的主建物資料為準）
    for (const c of commonB) {
      let s = 0; c.mains.forEach(m => { if (m.share) s += m.share.num / m.share.den; });
      if (c.mains.length && Math.abs(s - 1) > 1e-9) warn.push(`共有部分 ${c.no}：各主建物持分合計 = ${s.toFixed(6)}，不等於 1`);
      const onMain = new Set(mainB.filter(b => b.common.some(x => x.no === c.no)).map(b => b.no));
      const miss = c.mains.filter(m => !onMain.has(m.no)).map(m => m.no);
      if (miss.length) warn.push(`共有部分 ${c.no}：主建物 ${miss.join('、')} 沒有建物謄本，這部分持分沒有分配`);
    }
    mainB.forEach(b => { if (b.common.length > 2) info.push(`建號 ${b.no} 有 ${b.common.length} 個共有部分，第 3 個起放在工作表最右邊的隱藏欄（點欄位上方的「+」展開），已計入建物總面積`); });
    mainB.forEach(b => { if (!floorLabel(b)) warn.push(`建號 ${b.no}：無法判斷樓層（層次：${b.level}），樓次欄留空`); });
    // 他項權利：同一筆設定（收件年期＋字號）常同時登記在土地與建物上，合併成一筆
    const rightsMap = new Map();
    const props = [...L.map(l => ({ p: l, label: `${parseInt(l.parcel, 10)}地號`, kind: 'land', key: landKey(l.section, l.subsection, l.parcel) })), ...B.filter(b => !b.isCommon).map(b => ({ p: b, label: `${bldDisp(b.no)}建號`, kind: 'bldg', key: b.section + b.no }))];
    for (const { p, label, kind, key: pkey } of props) {
      const ownerByOrder = new Map(p.owners.map(o => [String(o.order).padStart(4, '0'), o.name]));
      // 順位：同一標的（同一所有權人持分）上，依登記次序先後
      const seen = new Map();
      for (const r of (p.rights || [])) {
        const tkey = r.targetOrder.join(',') || '*';
        const rank = (seen.get(tkey) || 0) + 1; seen.set(tkey, rank);
        const key = (r.year + '|' + r.no) || (p.parcel || p.no) + r.order;
        const owners = r.targetOrder.map(o => ownerByOrder.get(o.split('-')[0].padStart(4, '0'))).filter(Boolean);
        const where = { label, kind, key: pkey, owners, scope: r.scope, target: r.target };
        if (!rightsMap.has(key)) rightsMap.set(key, { ...r, rank, where: [where], firstProp: label });
        else { const x = rightsMap.get(key); x.where.push(where); x.rank = Math.min(x.rank, rank); }
      }
    }
    const rights = [...rightsMap.values()];
    return { blocks, lands: L, buildings: B, warn: [...new Set(warn)], info, rights };
  }

  // ───────────── 4. 寫入 xlsx（直接改 XML，保留範本格式）─────────────
  const colNum = c => { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const xesc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const LASTCOL = colNum('CE');
  const FIRST_ROW = 5;

  function parseSheet(xml) {
    const sdS = xml.indexOf('<sheetData'), sdE = xml.indexOf('</sheetData>');
    const sdOpenEnd = xml.indexOf('>', sdS) + 1;
    const body = xml.slice(sdOpenEnd, sdE);
    const rows = new Map();
    const rowRe = /<row r="(\d+)"([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    let m;
    while ((m = rowRe.exec(body))) {
      const cells = [];
      const cRe = /<c r="([A-Z]+)(\d+)"((?: [^>\/]*)?)(\/>|>([\s\S]*?)<\/c>)/g;
      let c;
      while ((c = cRe.exec(m[4] || ''))) cells.push({ col: c[1], attrs: c[3], inner: c[5], xml: c[0] });
      rows.set(+m[1], { attrs: m[2], cells });
    }
    return { head: xml.slice(0, sdOpenEnd), tail: xml.slice(sdE), rows };
  }

  function cellStyle(cell) { const m = cell && cell.attrs.match(/ s="(\d+)"/); return m ? m[1] : null; }

  function readSharedStrings(xml) {
    const out = [];
    (xml.match(/<si>[\s\S]*?<\/si>/g) || []).forEach(si => {
      out.push((si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join(''));
    });
    return out;
  }

  // 從範本第 5 列取公式，轉成任意列可用的樣板
  function formulaTemplate(f) {
    return f.replace(/\$E\$5(?!\d)/g, '$E${B}')
      .replace(/(^|[^A-Z$\]!\w])(\$?)([A-Z]{1,3})5(?!\d)/g, (all, pre, d, col) => d ? all : `${pre}${col}{R}`);
  }

  // ───────────── 他項權利表（另一個分頁，所有範本共用）─────────────
  const rocToSerial = t => { const m = String(t || '').match(/民國\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日/); if (!m) return null; return (Date.UTC(+m[1] + 1911, +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000; };
  const fracShort = t => { const f = frac(t); return f ? `${f.num}/${f.den}` : String(t || '').replace(/\s+/g, ''); };
  function coText(txt, kind) {
    if (!txt || /空白/.test(txt)) return '';
    const groups = []; let cur = null;
    String(txt).replace(/([^\s\d、，,]+?段(?:[^\s\d、，,]+?小段)?)|(\d{4,5})-(\d{3,4})/g, (m0, sec, a, b) => {
      if (sec) { cur = { sec, nos: [] }; groups.push(cur); }
      else if (cur) { const main = parseInt(a, 10), sub = parseInt(b, 10); cur.nos.push(sub ? `${main}-${sub}` : String(main)); }
    });
    return groups.filter(g => g.nos.length).map(g => `${kind === 'land' ? '土地' : '建物'}：${g.sec}${g.nos.join('、')}${kind === 'land' ? '地號' : '建號'}`).join('\n');
  }
  // 他項權利：跟著各列的地號／建號往右延伸。每筆權利一組欄位，同一列有多筆就再往右加一組。
  async function appendRightsColumns(zip, sheetPath, rowMeta, rights) {
    if (!rights.length || !rowMeta.length) return;
    // 1. 每一筆權利放到哪些列：有建物標的 → 該建號的列；只有土地 → 該所有權人在該地號的第一列
    const perRow = new Map();
    const push = (r, x) => { if (!perRow.has(r)) perRow.set(r, []); const arr = perRow.get(r); if (!arr.some(y => y.right === x.right)) arr.push(x); };
    for (const right of rights) {
      const bw = right.where.filter(w => w.kind === 'bldg'), lw = right.where.filter(w => w.kind === 'land');
      let placed = false;
      for (const w of bw) for (const m of rowMeta) if (m.bldKey === w.key) { push(m.r, { right, scope: [...lw.filter(l => l.owners.some(o => o === m.owner)).map(l => `${l.label} ${fracShort(l.scope)}`), `${w.label} ${fracShort(w.scope)}`].join('；') }); placed = true; }
      if (!placed) for (const w of lw) for (const m of rowMeta) if (m.landKey === w.key && m.firstOfOwner && w.owners.includes(m.owner)) { push(m.r, { right, scope: `${w.label} ${fracShort(w.scope)}` }); placed = true; }
      if (!placed) for (const w of lw) { const m = rowMeta.find(x => x.landKey === w.key); if (m) push(m.r, { right, scope: `${w.label} ${w.owners.join('、')} ${fracShort(w.scope)}` }); }
    }
    for (const arr of perRow.values()) arr.sort((x, y) => (x.right.regDate || '').localeCompare(y.right.regDate || ''));
    const groups = Math.max(0, ...[...perRow.values()].map(a => a.length));
    if (!groups) return;

    // 2. 樣式（附加在既有樣式表之後）
    let st = await zip.file('xl/styles.xml').async('string');
    const addTo = (xml, tag, item) => {
      const re = new RegExp(`<${tag}( count="(\\d+)")?([^>]*)>([\\s\\S]*?)</${tag}>`);
      const m = xml.match(re); if (!m) return { xml, idx: -1 };
      const item1 = { fonts: 'font', borders: 'border', cellXfs: 'xf', numFmts: 'numFmt', fills: 'fill' }[tag];
      const n = (m[4].match(new RegExp(`<${item1}[\\s>/]`, 'g')) || []).length;
      return { xml: xml.replace(re, `<${tag} count="${n + 1}"${m[3]}>${m[4]}${item}</${tag}>`), idx: n };
    };
    const FONT = '<name val="標楷體"/><family val="3"/><charset val="136"/>';
    let r1 = addTo(st, 'fonts', `<font><b/><sz val="11"/>${FONT}</font>`); st = r1.xml; const fHead = r1.idx;
    r1 = addTo(st, 'fonts', `<font><sz val="11"/>${FONT}</font>`); st = r1.xml; const fBody = r1.idx;
    r1 = addTo(st, 'fills', '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>'); st = r1.xml; const fillHead = r1.idx;
    r1 = addTo(st, 'borders', '<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>'); st = r1.xml; const bThin = r1.idx;
    st = st.replace(/<numFmts\b[^>]*\/>/, '');
    const ids = (st.match(/numFmtId="(\d+)"/g) || []).map(x => +x.match(/\d+/)[0]);
    const nMoney = Math.max(199, ...ids) + 1, nDate = nMoney + 1;
    const f1 = `<numFmt numFmtId="${nMoney}" formatCode="&quot;新台幣&quot;#,##0&quot;元&quot;"/>`, f2 = `<numFmt numFmtId="${nDate}" formatCode="[$-404]gge&quot;年&quot;m&quot;月&quot;d&quot;日&quot;;@"/>`;
    if (/<numFmts/.test(st)) st = addTo(addTo(st, 'numFmts', f1).xml, 'numFmts', f2).xml;
    else st = st.replace(/(<styleSheet[^>]*>)/, `$1<numFmts count="2">${f1}${f2}</numFmts>`);
    // 表頭可換行；資料列單行顯示（主表列高固定），欄寬放得下完整內容
    const AL = (h, wrap) => `<alignment horizontal="${h}" vertical="center"${wrap ? ' wrapText="1"' : ''}/>`;
    const xf = (numFmt, font, fill, align) => `<xf numFmtId="${numFmt}" fontId="${font}" fillId="${fill}" borderId="${bThin}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${align}</xf>`;
    r1 = addTo(st, 'cellXfs', xf(0, fHead, fillHead, AL('center', true))); st = r1.xml; const sHead = r1.idx;
    r1 = addTo(st, 'cellXfs', xf(0, fBody, 0, AL('left'))); st = r1.xml; const sText = r1.idx;
    r1 = addTo(st, 'cellXfs', xf(nMoney, fBody, 0, AL('right'))); st = r1.xml; const sMoney = r1.idx;
    r1 = addTo(st, 'cellXfs', xf(nDate, fBody, 0, AL('center'))); st = r1.xml; const sDate = r1.idx;
    zip.file('xl/styles.xml', st);

    // 3. 欄位：每組 10 欄
    const RANK = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
    const FIELDS = [
      ['權利種類', 16, x => ({ v: x.right.type })],
      ['登記次序', 9, x => ({ v: `第${RANK[x.right.rank - 1] || x.right.rank}順位` })],
      ['權利人', 32, x => ({ v: x.right.holder })],
      ['權利價值', 23, x => (x.right.amount != null ? { n: x.right.amount, s: sMoney } : { v: x.right.amountText })],
      ['擔保債權確定期日', 34, x => { const d = rocToSerial(x.right.detDate); if (d != null) return { n: d, s: sDate }; const t = (x.right.term || '').match(/自\s*(\d+)年(\d+)月(\d+)日至\s*(\d+)年(\d+)月(\d+)日/); return { v: t ? `${t[1]}.${t[2]}.${t[3]}～${t[4]}.${t[5]}.${t[6]}（存續期間）` : (x.right.term ? `存續期間${x.right.term}` : '') }; }],
      ['清償日期', 34, x => ({ v: x.right.repay })],
      ['登記日期', 17, x => { const d = rocToSerial(x.right.regDate); return d != null ? { n: d, s: sDate } : { v: x.right.regDate }; }],
      ['登記字號', 20, x => ({ v: x.right.no })],
      ['登記原因', 9, x => ({ v: x.right.reason })],
      ['設定權利範圍', 34, x => ({ v: x.scope })],
    ];
    let xml = await zip.file(sheetPath).async('string');
    let maxCol = 0;
    xml.replace(/<c r="([A-Z]+)\d+"/g, (m0, c) => { maxCol = Math.max(maxCol, colNum(c)); return m0; });
    const START = maxCol + 2, GW = FIELDS.length;
    const cx = (ref, style, it) => {
      if (it && it.n != null) return `<c r="${ref}" s="${it.s ?? style}"><v>${it.n}</v></c>`;
      if (!it || it.v == null || it.v === '') return `<c r="${ref}" s="${style}"/>`;
      return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xesc(it.v)}</t></is></c>`;
    };
    const addCells = (rn, cellsXml) => {
      const re = new RegExp(`<row r="${rn}"([^>]*?)(/>|>([\\s\\S]*?)</row>)`);
      if (re.test(xml)) xml = xml.replace(re, (m0, at, end, inner) => `<row r="${rn}"${at}>${inner || ''}${cellsXml}</row>`);
      else {
        // 該列不存在：依列號插入正確位置
        const rowsAfter = [...xml.matchAll(/<row r="(\d+)"/g)].find(m => +m[1] > rn);
        const tag = `<row r="${rn}">${cellsXml}</row>`;
        xml = rowsAfter ? xml.slice(0, rowsAfter.index) + tag + xml.slice(rowsAfter.index) : xml.replace('</sheetData>', tag + '</sheetData>');
      }
    };
    const merges = [];
    // 表頭：第 2 列組名、第 3 列欄名（第 3～4 列合併）
    let h2 = '', h3 = '', h4 = '';
    for (let g = 0; g < groups; g++) {
      const a = START + g * GW, z = a + GW - 1;
      h2 += cx(colName(a) + 2, sHead, { v: `他項權利（${RANK[g] || g + 1}）` }); for (let ci = a + 1; ci <= z; ci++) h2 += `<c r="${colName(ci)}2" s="${sHead}"/>`;
      merges.push(`${colName(a)}2:${colName(z)}2`);
      FIELDS.forEach(([lab], k) => { const ci = a + k; h3 += cx(colName(ci) + 3, sHead, { v: lab }); h4 += `<c r="${colName(ci)}4" s="${sHead}"/>`; merges.push(`${colName(ci)}3:${colName(ci)}4`); });
    }
    addCells(2, h2); addCells(3, h3); addCells(4, h4);
    // 資料列
    for (const m of rowMeta) {
      const list = perRow.get(m.r) || [];
      let out = '';
      for (let g = 0; g < groups; g++) {
        const x = list[g];
        FIELDS.forEach(([, , fn], k) => { out += cx(colName(START + g * GW + k) + m.r, sText, x ? fn(x) : null); });
      }
      addCells(m.r, out);
    }
    // 合併儲存格、欄寬
    if (/<mergeCells/.test(xml)) xml = xml.replace(/<mergeCells count="(\d+)">([\s\S]*?)<\/mergeCells>/, (m0, n, inner) => `<mergeCells count="${+n + merges.length}">${inner}${merges.map(x => `<mergeCell ref="${x}"/>`).join('')}</mergeCells>`);
    else xml = xml.replace('</sheetData>', `</sheetData><mergeCells count="${merges.length}">${merges.map(x => `<mergeCell ref="${x}"/>`).join('')}</mergeCells>`);
    const a0 = START, z0 = START + groups * GW - 1;
    const newCols = []; for (let ci = a0; ci <= z0; ci++) newCols.push(`<col min="${ci}" max="${ci}" width="${FIELDS[(ci - a0) % GW][1]}" customWidth="1"/>`);
    if (/<cols>/.test(xml)) {
      xml = xml.replace(/<cols>([\s\S]*?)<\/cols>/, (all, inner) => {
        const out = [];
        (inner.match(/<col [^>]*\/>/g) || []).forEach(c => {
          const mn = +c.match(/ min="(\d+)"/)[1], mx = +c.match(/ max="(\d+)"/)[1];
          if (mx < a0 || mn > z0) { out.push([mn, c]); return; }
          if (mn < a0) out.push([mn, c.replace(/ max="\d+"/, ` max="${a0 - 1}"`)]);
          if (mx > z0) out.push([z0 + 1, c.replace(/ min="\d+"/, ` min="${z0 + 1}"`)]);
        });
        newCols.forEach((c, i) => out.push([a0 + i, c]));
        out.sort((x, y) => x[0] - y[0]);
        return `<cols>${out.map(x => x[1]).join('')}</cols>`;
      });
    } else xml = xml.replace('<sheetData', `<cols>${newCols.join('')}</cols><sheetData`);
    xml = xml.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/, (m0, c1, r1x, c2, r2) => `<dimension ref="${c1}${r1x}:${colName(Math.max(colNum(c2), z0))}${r2}"/>`);
    zip.file(sheetPath, xml);
  }

  // 檔案作者：只換標籤內文字，保留原標籤（含命名空間宣告）
  async function setAuthor(zip, author) {
    if (!author || !zip.file('docProps/core.xml')) return;
    let cx = await zip.file('docProps/core.xml').async('string');
    const a = xesc(author), now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const setTag = (xml, tag, ns) => {
      const open = new RegExp(`(<${tag}\\b[^>]*?)(\\/>|>[\\s\\S]*?<\\/${tag}>)`);
      if (open.test(xml)) return xml.replace(open, (m, head) => `${head}>${a}</${tag}>`);
      return xml.replace('</cp:coreProperties>', `<${tag}${ns}>${a}</${tag}></cp:coreProperties>`);
    };
    cx = setTag(cx, 'dc:creator', /xmlns:dc=/.test(cx.slice(0, cx.indexOf('>', cx.indexOf('<cp:coreProperties')))) ? '' : ' xmlns:dc="http://purl.org/dc/elements/1.1/"');
    cx = setTag(cx, 'cp:lastModifiedBy', '');
    cx = cx.replace(/(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${now}$2`);
    zip.file('docProps/core.xml', cx);
  }
  // 公式位置變動後移除 calcChain，Excel 開檔會自動重建
  async function dropCalcChain(zip) {
    if (!zip.file('xl/calcChain.xml')) return;
    zip.remove('xl/calcChain.xml');
    const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');  // 讀最新版本（可能已加入他項權利表）
    zip.file('xl/_rels/workbook.xml.rels', wbRels.replace(/<Relationship [^>]*calcChain[^>]*\/>/, ''));
    const ct = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', ct.replace(/<Override [^>]*calcChain[^>]*\/>/, ''));
  }

  async function fillWorkbook(JSZip, templateBuf, model, opts) {
    opts = opts || {};
    const zip = await JSZip.loadAsync(templateBuf);
    const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const wb = await zip.file('xl/workbook.xml').async('string');
    // 找第一個工作表
    const firstSheetRid = (wb.match(/<sheet [^>]*r:id="([^"]+)"/) || [])[1];
    const target = (wbRels.match(new RegExp(`<Relationship [^>]*Id="${firstSheetRid}"[^>]*Target="([^"]+)"`)) || wbRels.match(new RegExp(`<Relationship [^>]*Target="([^"]+)"[^>]*Id="${firstSheetRid}"`)) || [])[1];
    const sheetPath = 'xl/' + target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const xml = await zip.file(sheetPath).async('string');
    const ssFile = zip.file('xl/sharedStrings.xml');
    const sst = ssFile ? readSharedStrings(await ssFile.async('string')) : [];
    const sh = parseSheet(xml);

    // 範本樣式：首列 = 第5列、中間列 = 第6列、末列 = 第11列
    const byCol = r => { const m = {}; (sh.rows.get(r) || { cells: [] }).cells.forEach(c => { m[c.col] = c; }); return m; };
    const T = { first: byCol(5), mid: byCol(6), last: byCol(11) };
    const row5Attrs = sh.rows.get(5).attrs;
    // 第 5 列公式與常數
    const F = {}, CONST = {};
    for (const [col, c] of Object.entries(T.first)) {
      if (colNum(col) > LASTCOL) continue;
      const fm = (c.inner || '').match(/<f[^>]*>([\s\S]*?)<\/f>/);
      if (fm) F[col] = formulaTemplate(fm[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'));
      else if (/<v>/.test(c.inner || '')) {
        const v = c.inner.match(/<v>([\s\S]*?)<\/v>/)[1];
        CONST[col] = / t="s"/.test(c.attrs) ? sst[+v] : +v;
      }
    }
    const K = CONST.K ?? 3, M = CONST.M ?? 0.8;

    // 舊資料區範圍
    let oldLast = FIRST_ROW;
    for (const [r, row] of sh.rows) if (r >= FIRST_ROW && row.cells.some(c => colNum(c.col) <= LASTCOL && c.inner)) oldLast = Math.max(oldLast, r);

    // 範本第 5 列是否有 CD:CE 合併（公司範本有，內建範本沒有）
    const origMerges = (sh.tail.match(/<mergeCell ref="([^"]+)"\/>/g) || []).map(m => m.match(/ref="([^"]+)"/)[1]);
    const cdMerge = origMerges.includes('CD5:CE5');
    // 案名寫在第 1 列包含 D 欄的合併儲存格左上角（沒有就寫 D1）
    let titleRef = 'D1';
    for (const mr of origMerges) { const m = mr.match(/^([A-Z]+)1:([A-Z]+)(\d+)$/); if (m && colNum(m[1]) <= 4 && colNum(m[2]) >= 4) { titleRef = m[1] + '1'; break; } }
    // 共有部分欄位標題
    const slotNos = [new Set(), new Set()];
    // 第 3 個以後的共有部分：放在工作表目前用到的最後一欄之後（不碰右側試算區），每組 7 欄
    let maxUsedCol = LASTCOL;
    for (const [, row] of sh.rows) for (const c of row.cells) maxUsedCol = Math.max(maxUsedCol, colNum(c.col));
    const EXTRA_START = maxUsedCol + 2, EXTRA_W = 7;
    const extraCols = g => { const b = EXTRA_START + g * EXTRA_W; return ['no', 'area', 'num', 'sl', 'den', 'm2', 'ping'].reduce((o, k, i) => (o[k] = colName(b + i), o), {}); };
    const EXTRA_STYLE_FROM = { no: 'P', area: 'AU', num: 'AV', sl: 'AW', den: 'AX', m2: 'AY', ping: 'AZ' };
    let maxExtra = 0;

    // 產生每列資料
    const gen = new Map(); // row -> {col: {v, t:'n'|'s'|'f'}}
    const rowMeta = [];
    const merges = [];
    let r = FIRST_ROW;
    for (const blk of model.blocks) {
      const b0 = r, l = blk.land;
      blk.rows.forEach((row, i) => {
        const pos = blk.rows.length === 1 ? 'first' : i === 0 ? 'first' : i === blk.rows.length - 1 ? 'last' : 'mid';
        const d = {};
        const put = (col, v, t) => { if (v !== null && v !== undefined && v !== '') d[col] = { v, t: t || (typeof v === 'number' ? 'n' : 's') }; };
        const fx = col => { if (F[col]) d[col] = { v: F[col].replace(/\{R\}/g, r).replace(/\{B\}/g, b0), t: 'f' }; };
        if (i === 0) {
          put('A', l.township); put('B', l.section); put('C', l.subsection);
          put('D', parcelDisp(l.parcel)); put('E', l.area); put('O', l.landValue);
        }
        if (row.owner) { put('F', row.owner.name); put('G', row.owner.id); }
        if (!row.jointCont) put('I', '/');
        if (row.land) {
          put('H', row.land.num); put('J', row.land.den);
          if (F.K) fx('K'); else put('K', K);
          if (F.M) fx('M'); else put('M', M);
          d.L = { v: `$E$${b0}*0.3025*H${r}/J${r}`, t: 'f' };
          fx('N');
        }
        // 建物
        put('T', '/'); put('AD', '/'); put('AQ', '/'); put('AW', '/');
        const b = row.b;
        if (b) {
          const bs = row.bshare || { num: 1, den: 1 };
          put('P', bldDisp(b.no)); put('R', b.area); put('S', bs.num); put('U', bs.den);
          const part = bs.num !== bs.den;
          d.V = { v: part ? `R${r}*0.3025*S${r}/U${r}` : `R${r}*0.3025`, t: 'f' };
          const acc = {};
          b.annex.forEach(a => { const col = ANNEX_COL.find(x => x[0].test(a.usage))[1]; acc[col] = (acc[col] || 0) + (a.area || 0); });
          for (const [col, v] of Object.entries(acc)) put(col, Math.round(v * 100) / 100);
          if (acc.W) { put('X', bs.num); put('Y', '/'); put('Z', bs.den); }
          if (acc.AB) { put('AC', bs.num); put('AE', bs.den); }
          d.AF = { v: acc.AB && part ? `AB${r}*0.3025*AC${r}/AE${r}` : `AB${r}*0.3025`, t: 'f' };
          if (acc.AI) d.AJ = { v: `AI${r}*0.3025`, t: 'f' };
          if (acc.AK) d.AL = { v: `AK${r}*0.3025`, t: 'f' };
          if (acc.AM) d.AN = { v: `AM${r}*0.3025`, t: 'f' };
          const cs = b.common.slice().sort((x, y) => x.no.localeCompare(y.no));
          const slots = [['AO', 'AP', 'AR'], ['AU', 'AV', 'AX']];
          slots.forEach((s, k) => {
            const c = cs[k];
            if (c && c.share) { put(s[0], c.area); put(s[1], c.share.num * (part ? bs.num : 1)); put(s[2], c.share.den * (part ? bs.den : 1)); slotNos[k].add(c.no); }
            else { put(s[0], 0); put(s[1], 0); put(s[2], 1); }
          });
          put('BB', (floorLabel(b) ? floorLabel(b) + '/' + (b.floors || '') : ''));
          put('BC', b.door);
        } else {
          d.V = { v: `R${r}*0.3025`, t: 'f' }; d.AF = { v: `AB${r}*0.3025`, t: 'f' };
          put('AO', 0); put('AP', 0); put('AR', 1); put('AU', 0); put('AV', 0); put('AX', 1);
        }
        ['AA', 'AH', 'AS', 'AT', 'AY', 'AZ', 'BA'].forEach(fx);
        if (b) {
          const bs = row.bshare || { num: 1, den: 1 }, part = bs.num !== bs.den;
          const extras = b.common.slice().sort((x, y) => x.no.localeCompare(y.no)).slice(2);
          const pingCols = [];
          extras.forEach((c, g) => {
            const X = extraCols(g);
            maxExtra = Math.max(maxExtra, g + 1);
            put(X.no, c.no.replace(/-000$/, '').replace(/^0+/, ''));
            put(X.area, c.area); put(X.sl, '/');
            if (c.share) { put(X.num, c.share.num * (part ? bs.num : 1)); put(X.den, c.share.den * (part ? bs.den : 1)); }
            d[X.m2] = { v: `IF(${X.den}${r}=0,0,${X.area}${r}*${X.num}${r}/${X.den}${r})`, t: 'f' };
            d[X.ping] = { v: `${X.m2}${r}*0.3025`, t: 'f' };
            pingCols.push(X.ping + r);
          });
          if (pingCols.length && d.BA) d.BA = { v: d.BA.v + '+' + pingCols.join('+'), t: 'f' };
        }
        put('BD', 1);
        ['BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'CD'].forEach(fx);
        if (!row.land) { delete d.BX; delete d.BY; }
        gen.set(r, { d, pos });
        rowMeta.push({ r, landKey: landKey(l.section, l.subsection, l.parcel), owner: row.owner && row.owner.name, bldKey: row.b ? row.b.section + row.b.no : null, firstOfOwner: !!row.land || (row.joint && !row.jointCont) });
        if (cdMerge) merges.push(`CD${r}:CE${r}`);
        r++;
      });
      if (r - 1 > b0) { merges.push(`D${b0}:D${r - 1}`, `E${b0}:E${r - 1}`); }
      (blk.jointRanges || []).forEach(([a, z]) => { ['H', 'I', 'J', 'K', 'L', 'M', 'N'].forEach(c => merges.push(`${c}${b0 + a}:${c}${b0 + z}`)); });
    }
    const newLast = r - 1;

    // 組回 sheetData
    const cellXml = (ref, style, item) => {
      const s = style ? ` s="${style}"` : '';
      if (!item) return `<c r="${ref}"${s}/>`;
      if (item.t === 'f') return `<c r="${ref}"${s}><f>${xesc(item.v)}</f></c>`;
      if (item.t === 'n') return `<c r="${ref}"${s}><v>${item.v}</v></c>`;
      return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xesc(item.v)}</t></is></c>`;
    };
    const maxRow = Math.max(oldLast, newLast);
    for (let rr = FIRST_ROW; rr <= maxRow; rr++) {
      const existing = sh.rows.get(rr);
      const keep = existing ? existing.cells.filter(c => colNum(c.col) > LASTCOL) : [];
      const g = gen.get(rr);
      const cells = [];
      if (g) {
        const tpl = T[g.pos];
        for (let ci = 1; ci <= LASTCOL; ci++) {
          const col = colName(ci);
          const style = cellStyle(tpl[col]) || cellStyle(T.first[col]);
          cells.push({ col, xml: cellXml(col + rr, style, g.d[col]) });
        }
        for (let gi = 0; gi < maxExtra; gi++) {
          const X = extraCols(gi);
          for (const [k, col] of Object.entries(X)) {
            const src = EXTRA_STYLE_FROM[k];
            const style = cellStyle(tpl[src]) || cellStyle(T.first[src]);
            cells.push({ col, xml: cellXml(col + rr, style, g.d[col]) });
          }
        }
      }
      const all = cells.concat(keep.map(c => ({ col: c.col, xml: c.xml })));
      all.sort((a, b) => colNum(a.col) - colNum(b.col));
      const attrs = g ? row5Attrs : (existing ? existing.attrs : row5Attrs);
      if (!existing && !g) continue;
      sh.rows.set(rr, { attrs, cells: all.map(c => ({ col: c.col, xml: c.xml, attrs: '', inner: '' })), rebuilt: true });
    }
    const rowNums = [...sh.rows.keys()].sort((a, b) => a - b);
    let tail = sh.tail;
    // 合併儲存格：移除資料區（A..CE、第5列以下）舊合併，加入新的
    tail = tail.replace(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/, (all, inner) => {
      const keepM = (inner.match(/<mergeCell ref="[^"]+"\/>/g) || []).filter(mc => {
        const ref = mc.match(/ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/);
        return !(ref && +ref[2] >= FIRST_ROW && colNum(ref[1]) <= LASTCOL);
      });
      const allM = keepM.concat(merges.map(x => `<mergeCell ref="${x}"/>`));
      return `<mergeCells count="${allM.length}">${allM.join('')}</mergeCells>`;
    });
    let head = sh.head;
    // 標題列：共有部分欄名、D1 案名
    const setHeadCell = (ref, text) => {
      const rowNo = +ref.match(/\d+/)[0];
      const row = sh.rows.get(rowNo); if (!row) return;
      row.cells = row.cells.map(c => {
        if (c.col + rowNo !== ref) return c;
        const s = cellStyle(c); return { ...c, xml: `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${xesc(text)}</t></is></c>` };
      });
    };
    if (slotNos[0].size) setHeadCell('AO3', slotNos[0].size === 1 ? `${parseInt([...slotNos[0]][0], 10)}\n共有部分` : '共有部分\n(一)');
    if (slotNos[1].size) setHeadCell('AU3', slotNos[1].size === 1 ? `${parseInt([...slotNos[1]][0], 10)}\n共有部分` : '共有部分\n(二)');
    if (maxExtra) {
      const CN_NUM = ['三', '四', '五', '六', '七', '八', '九', '十'];
      const addHeadCell = (rowNo, col, text, styleFrom) => {
        const row = sh.rows.get(rowNo); if (!row) return;
        const srcCell = row.cells.find(c => c.col === styleFrom);
        const s = srcCell ? cellStyle(srcCell) : null;
        row.cells = row.cells.filter(c => c.col !== col);
        row.cells.push({ col, attrs: '', inner: '', xml: text === '' ? `<c r="${col}${rowNo}"${s ? ` s="${s}"` : ''}/>` : `<c r="${col}${rowNo}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${xesc(text)}</t></is></c>` });
        row.cells.sort((a, b) => colNum(a.col) - colNum(b.col));
      };
      for (let gi = 0; gi < maxExtra; gi++) {
        const X = extraCols(gi), n = CN_NUM[gi] || String(gi + 3);
        const h3 = { no: `共有部分\n(${n})建號`, area: `共有部分\n(${n})`, num: '持分', sl: '', den: '', m2: '持有面積', ping: '持有面積' };
        const h4 = { no: '', area: '(㎡)', num: '分子', sl: '/', den: '分母', m2: '(㎡)', ping: '(坪)' };
        for (const k of Object.keys(X)) { addHeadCell(3, X[k], h3[k], EXTRA_STYLE_FROM[k] === 'P' ? 'AU' : EXTRA_STYLE_FROM[k]); addHeadCell(4, X[k], h4[k], EXTRA_STYLE_FROM[k] === 'P' ? 'AU' : EXTRA_STYLE_FROM[k]); }
        merges.push(`${X.num}3:${X.den}3`);
      }
      // 欄位設成可展開的隱藏群組（Excel 欄位上方會出現「+」）
      const a = EXTRA_START, z = EXTRA_START + maxExtra * EXTRA_W - 1;
      const widths = { no: 8, area: 9, num: 6, sl: 2, den: 6, m2: 9, ping: 9 };
      const newCols = [];
      for (let ci = a; ci <= z; ci++) { const k = Object.keys(EXTRA_STYLE_FROM)[(ci - a) % EXTRA_W]; newCols.push(`<col min="${ci}" max="${ci}" width="${widths[k]}" customWidth="1" hidden="1" outlineLevel="1"/>`); }
      if (/<cols>/.test(head)) {
        head = head.replace(/<cols>([\s\S]*?)<\/cols>/, (all, inner) => {
          const out = [];
          (inner.match(/<col [^>]*\/>/g) || []).forEach(c => {
            const mn = +c.match(/ min="(\d+)"/)[1], mx = +c.match(/ max="(\d+)"/)[1];
            if (mx < a || mn > z) { out.push([mn, c]); return; }
            if (mn < a) out.push([mn, c.replace(/ max="\d+"/, ` max="${a - 1}"`)]);
            if (mx > z) out.push([z + 1, c.replace(/ min="\d+"/, ` min="${z + 1}"`)]);
          });
          newCols.forEach((c, i) => out.push([a + i, c]));
          out.sort((x, y) => x[0] - y[0]);
          return `<cols>${out.map(x => x[1]).join('')}</cols>`;
        });
      } else {
        head = head.replace(/(<sheetData)/, `<cols>${newCols.join('')}</cols>$1`);
      }
      head = /<sheetFormatPr[^>]*outlineLevelCol=/.test(head)
        ? head.replace(/(<sheetFormatPr[^>]*?)outlineLevelCol="\d+"/, '$1outlineLevelCol="1"')
        : head.replace(/<sheetFormatPr\b/, '<sheetFormatPr outlineLevelCol="1"');
      head = head.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/, (m0, c1, r1, c2, r2) => `<dimension ref="${c1}${r1}:${colName(Math.max(colNum(c2), z))}${r2}"/>`);
    }
    if (model.lands.length) {
      const l0 = model.lands[0];
      const city = (l0.city || '').replace(/^臺/, '台');
      setHeadCell(titleRef, `${city}${l0.township}${l0.section}${l0.subsection}${parcelDisp(l0.parcel)}地號` + (model.lands.length > 1 ? `等${model.lands.length}筆土地` : ''));
    }
    // 序列化
    let body = '';
    for (const rn of rowNums) {
      const row = sh.rows.get(rn);
      body += row.cells.length ? `<row r="${rn}"${row.attrs}>${row.cells.map(c => c.xml).join('')}</row>` : `<row r="${rn}"${row.attrs}/>`;
    }
    head = head.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/, (a, c1, r1, c2, r2) => `<dimension ref="${c1}${r1}:${c2}${Math.max(+r2, newLast)}"/>`);
    zip.file(sheetPath, head + body + tail);

    // 讓 Excel 開啟時重算；移除 calcChain（公式位置已變）
    // 列印範圍延伸到最後一列資料
    let wbx0 = wb.replace(/(<definedName name="_xlnm\.Print_Area"[^>]*>[^<]*?!\$[A-Z]+\$\d+:\$[A-Z]+\$)(\d+)(<\/definedName>)/, (a, pre, n, post) => pre + Math.max(newLast, FIRST_ROW) + post);
    let wbx = wbx0.replace(/<calcPr([^>]*?)\s*\/>/, (a, at) => `<calcPr${at.replace(/ fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`);
    zip.file('xl/workbook.xml', wbx);
    if (opts.rights !== false) await appendRightsColumns(zip, sheetPath, rowMeta, model.rights || []);
    await setAuthor(zip, opts.author);
    await dropCalcChain(zip, wbRels);
    return { buf: await zip.generateAsync({ type: root.document ? 'blob' : 'nodebuffer', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), rows: newLast - FIRST_ROW + 1, oldLast };
  }


  // ───────────── 5. 作者範本（權屬+分配表）─────────────
  // 版面：第 1～4 列表頭；範本第 5 列＝區塊首列、6＝中間列、14＝末列樣式；15＝合計列、16＝附加列。
  // 一筆地號一個區塊（地號層級欄位跨區塊合併），每列一位所有權人。
  async function fillAuthorWorkbook(JSZip, templateBuf, model, opts) {
    opts = opts || {};
    const zip = await JSZip.loadAsync(templateBuf);
    const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const wb = await zip.file('xl/workbook.xml').async('string');
    const rid = (wb.match(/<sheet [^>]*r:id="([^"]+)"/) || [])[1];
    const target = (wbRels.match(new RegExp(`<Relationship [^>]*Id="${rid}"[^>]*Target="([^"]+)"`)) || wbRels.match(new RegExp(`<Relationship [^>]*Target="([^"]+)"[^>]*Id="${rid}"`)) || [])[1];
    const sheetPath = 'xl/' + target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const sh = parseSheet(await zip.file(sheetPath).async('string'));
    const LAST = colNum('AX');
    const byCol = rr => { const m = {}; (sh.rows.get(rr) || { cells: [] }).cells.forEach(c => { m[c.col] = cellStyle(c); }); return m; };
    const ST = { first: byCol(5), mid: byCol(6), last: byCol(14), total: byCol(15), extra: byCol(16) };
    const RA = { data: (sh.rows.get(6) || sh.rows.get(5)).attrs, first: sh.rows.get(5).attrs, total: (sh.rows.get(15) || sh.rows.get(5)).attrs, extra: (sh.rows.get(16) || sh.rows.get(5)).attrs };

    const LAND_COLS = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR'];
    const fmtNum = x => String(Math.round(x * 1e6) / 1e6);
    const gen = new Map(); const merges = []; const info = []; const rowMeta = [];
    // 先算出每個區塊的列數，才能決定合計列位置
    let r = FIRST_ROW;
    const layout = model.blocks.map(blk => { const b0 = r; r += blk.rows.length; return { blk, b0, e: r - 1 }; });
    const lastData = r - 1, T = r, X = r + 1;

    for (const { blk, b0, e } of layout) {
      const l = blk.land;
      blk.rows.forEach((row, i) => {
        const rr = b0 + i;
        const pos = blk.rows.length === 1 ? 'first' : i === 0 ? 'first' : i === blk.rows.length - 1 ? 'last' : 'mid';
        const d = {};
        const put = (col, v) => { if (v !== null && v !== undefined && v !== '') d[col] = { v, t: typeof v === 'number' ? 'n' : 's' }; };
        const f = (col, v) => { d[col] = { v, t: 'f' }; };
        if (i === 0) {
          put('B', String(parcelDisp(l.parcel)));
          if (l.zoningRaw && !/空白/.test(l.zoningRaw)) put('C', l.zoningRaw);
          put('D', l.area); f('E', `D${rr}*0.3025`);
          put('F', l.landValue); f('G', `F${rr}/0.3025`); f('H', `D${rr}*F${rr}`);
          f('J', `I${rr}/0.3025`); f('K', `I${rr}*D${rr}`);
          put('O', 1); put('P', '/'); put('Q', 1);
          f('R', `D${rr}*O${rr}/Q${rr}`); f('S', `R${rr}*0.3025`); f('T', `S${rr}`); f('U', `T${rr}/$S$${T}`);
          f('AJ', `$AJ$4*E${rr}/$E$${T}`); f('AK', `$AK$4*E${rr}/$E$${T}`); f('AL', `$AL$4*E${rr}/$E$${T}`);
          f('AM', `$AM$4*E${rr}/$E$${T}`); f('AN', `IF(AL${rr}=0,0,AM${rr}/AL${rr})`);
          f('AO', `$AO$4*E${rr}/$E$${T}`); f('AP', `$AP$4*E${rr}/$E$${T}`);
          f('AQ', `$AQ$4*E${rr}/$E$${T}`); f('AR', `$AR$4*E${rr}/$E$${T}`);
        }
        if (row.owner) { put('L', row.owner.name); put('AI', row.owner.name); }
        if (row.land) f('V', `$E$${b0}*${row.land.num}/${row.land.den}`);
        const b = row.b;
        if (b) {
          const bs = row.bshare || { num: 1, den: 1 }, part = bs.num !== bs.den;
          const sh2 = part ? `*${bs.num}/${bs.den}` : '';
          put('W', parseInt(b.no.replace('-', ''), 10)); put('X', b.door);  // 範本格式 00000000：03627-000 → 03627000
          if (b.area) f('Z', `${fmtNum(b.area)}${sh2}`);
          const annex = b.annex.filter(a => a.area);
          if (annex.length) f('AB', annex.map(a => fmtNum(a.area) + sh2).join('+'));
          const cm = b.common.filter(c => c.area && c.share);
          if (cm.length) f('AD', cm.map(c => `${fmtNum(c.area)}*${c.share.num}/${c.share.den}${sh2}`).join('+'));
          f('AA', `Z${rr}*0.3025`); f('AC', `AB${rr}*0.3025`); f('AE', `AD${rr}*0.3025`);
          f('AF', `Z${rr}+AB${rr}+AD${rr}`);
          if (cm.length > 2) info.push(`建號 ${b.no} 有 ${cm.length} 個共有部分，已全部加總在「共有」欄（AD），算式可在儲存格中查看`);
        }
        f('AG', `AF${rr}*0.3025`); f('AH', `AG${rr}*(1-$AH$4)`);
        put('AS', 5); f('AU', `$AU$4*AG${rr}`); put('AV', 1); f('AW', `V${rr}`);
        gen.set(rr, { d, style: ST[pos], attrs: i === 0 ? RA.first : RA.data });
        rowMeta.push({ r: rr, landKey: landKey(l.section, l.subsection, l.parcel), owner: row.owner && row.owner.name, bldKey: row.b ? row.b.section + row.b.no : null, firstOfOwner: !!row.land || (row.joint && !row.jointCont) });
        merges.push(`L${rr}:N${rr}`);
      });
      if (e > b0) {
        LAND_COLS.forEach(c => merges.push(`${c}${b0}:${c}${e}`));
        (blk.jointRanges || []).forEach(([a, z]) => merges.push(`V${b0 + a}:V${b0 + z}`));
      }
    }
    // A 欄：同一地段（含小段）連續的區塊合併成一格
    let runStart = null, runName = null;
    const flush = endRow => { if (runStart !== null) { gen.get(runStart).d.A = { v: runName, t: 's' }; if (endRow > runStart) merges.push(`A${runStart}:A${endRow}`); } };
    layout.forEach(({ blk, b0, e }, k) => {
      const nm = blk.land.section + (blk.land.subsection || '');
      if (nm !== runName) { flush(b0 - 1); runStart = b0; runName = nm; }
      if (k === layout.length - 1) flush(e);
    });
    // 合計列與附加列
    const tot = {};
    const SUMC = ['D', 'E', 'H', 'K', 'R', 'S', 'T', 'U', 'V', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AJ', 'AK', 'AL', 'AM', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AV', 'AW'];
    SUMC.forEach(c => { tot[c] = { v: `SUM(${c}${FIRST_ROW}:${c}${lastData})`, t: 'f' }; });
    Object.assign(tot, {
      F: { v: `H${T}/D${T}`, t: 'f' }, G: { v: `H${T}/E${T}`, t: 'f' }, I: { v: `K${T}/D${T}`, t: 'f' }, J: { v: `K${T}/E${T}`, t: 'f' },
      W: { v: `COUNTA(W${FIRST_ROW}:W${lastData})`, t: 'f' }, AN: { v: `IF(AL${T}=0,0,AM${T}/AL${T})`, t: 'f' },
    });
    gen.set(T, { d: tot, style: ST.total, attrs: RA.total });
    gen.set(X, { d: { E: { v: `D${X}*0.3025`, t: 'f' }, AW: { v: `AW${T}/E${T}`, t: 'f' } }, style: ST.extra, attrs: RA.extra });
    merges.push(`B${T}:C${T}`, `W${T}:Y${T}`);

    // 重建第 5 列以後
    const cellXml = (ref, style, item) => {
      const s = style ? ` s="${style}"` : '';
      if (!item) return `<c r="${ref}"${s}/>`;
      if (item.t === 'f') return `<c r="${ref}"${s}><f>${xesc(item.v)}</f></c>`;
      if (item.t === 'n') return `<c r="${ref}"${s}><v>${item.v}</v></c>`;
      return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xesc(item.v)}</t></is></c>`;
    };
    for (const k of [...sh.rows.keys()]) if (k >= FIRST_ROW) sh.rows.delete(k);
    for (const [rr, g] of gen) {
      const cells = [];
      for (let c = 1; c <= LAST; c++) { const col = colName(c); cells.push({ col, xml: cellXml(col + rr, g.style[col], g.d[col]) }); }
      sh.rows.set(rr, { attrs: g.attrs, cells });
    }
    // 表頭：年度、案名
    const setCell = (ref, text) => {
      const rowNo = +ref.match(/\d+/)[0], col = ref.match(/[A-Z]+/)[0]; const row = sh.rows.get(rowNo); if (!row) return;
      row.cells = row.cells.map(c => c.col !== col ? c : { ...c, xml: `<c r="${ref}"${cellStyle(c) ? ` s="${cellStyle(c)}"` : ''} t="inlineStr"><is><t xml:space="preserve">${xesc(text)}</t></is></c>` });
    };
    const yr = (model.lands.find(x => x.landValueYear) || {}).landValueYear;
    if (yr) { setCell('F1', `${yr}年公告現值`); setCell('I1', `${yr}年公告地價`); }
    let body = '';
    for (const rn of [...sh.rows.keys()].sort((a, b) => a - b)) {
      const row = sh.rows.get(rn);
      body += row.cells.length ? `<row r="${rn}"${row.attrs}>${row.cells.map(c => c.xml).join('')}</row>` : `<row r="${rn}"${row.attrs}/>`;
    }
    let tail = sh.tail.replace(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/, (all, inner) => {
      const keep = (inner.match(/<mergeCell ref="[^"]+"\/>/g) || []).filter(mc => +mc.match(/ref="[A-Z]+(\d+):/)[1] < FIRST_ROW);
      const allM = keep.concat(merges.map(x => `<mergeCell ref="${x}"/>`));
      return `<mergeCells count="${allM.length}">${allM.join('')}</mergeCells>`;
    });
    // 頁首案名、頁尾製作日期（民國）
    const l0 = model.lands[0];
    const title = l0 ? `${l0.township}${l0.section}${l0.subsection}${parcelDisp(l0.parcel)}地號` + (model.lands.length > 1 ? `等共${model.lands.length}筆土地` : '') : '地號清冊';
    const now = new Date(), roc = `${now.getFullYear() - 1911}.${String(now.getMonth() + 1).padStart(2, '0')}.${String(now.getDate()).padStart(2, '0')}`;
    tail = tail.replace(/<oddHeader>[\s\S]*?<\/oddHeader>/, `<oddHeader>&amp;C&amp;"-,粗體"&amp;22${xesc(title)}</oddHeader>`)
      .replace(/<oddFooter>[\s\S]*?<\/oddFooter>/, `<oddFooter>&amp;R製作日期：${roc}</oddFooter>`);
    let head = sh.head.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/, (a, c1, r1, c2) => `<dimension ref="${c1}${r1}:${c2}${X}"/>`);
    zip.file(sheetPath, head + body + tail);
    let wbx = wb.replace(/(<definedName name="_xlnm\.Print_Area"[^>]*>[^<]*?!\$[A-Z]+\$\d+:\$[A-Z]+\$)(\d+)(<\/definedName>)/, (a, pre, n, post) => pre + T + post)
      .replace(/<calcPr([^>]*?)\s*\/>/, (a, at) => `<calcPr${at.replace(/ fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`);
    zip.file('xl/workbook.xml', wbx);
    if (opts.rights !== false) await appendRightsColumns(zip, sheetPath, rowMeta, model.rights || []);
    await setAuthor(zip, opts.author);
    await dropCalcChain(zip, wbRels);
    return { buf: await zip.generateAsync({ type: root.document ? 'blob' : 'nodebuffer', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), rows: lastData - FIRST_ROW + 1, info };
  }

  // ───────────── 6. PDF 文字（pdf.js）─────────────
  // meta（選填）：會填入 { producer, creator }，用來判斷檔案是否被「列印成 PDF」過
  async function pdfText(pdfjsLib, data, meta) {
    const pdf = await pdfjsLib.getDocument({ data, isEvalSupported: false }).promise;
    if (meta) { try { const md = await pdf.getMetadata(); meta.producer = (md.info && md.info.Producer) || ''; meta.creator = (md.info && md.info.Creator) || ''; } catch (e) { /* 讀不到屬性不影響辨識 */ } }
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const tc = await (await pdf.getPage(i)).getTextContent();
      let t = '';
      for (const it of tc.items) { t += it.str; if (it.hasEOL) t += '\n'; }
      pages.push(t);
    }
    await pdf.destroy();
    return pages.join('\n');
  }

  root.TranscriptCore = { parseText, buildModel, fillWorkbook, fillAuthorWorkbook, pdfText, floorLabel, _cleanText: cleanText };
})(typeof window !== 'undefined' ? window : globalThis);
