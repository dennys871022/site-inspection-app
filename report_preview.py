"""
Word 版面預覽＋直接在預覽上換照片

- 依 Word 樣板的樣子畫出每一頁（每頁 4 張，上面是工程資料表頭）
- 電腦：按住照片拖到另一張上 → 兩張交換
- 手機：長按照片約半秒再拖，或「點一張、再點另一張」→ 兩張交換
- 文字要不要跟著照片走，由 app 上的開關決定（這裡只回傳新的照片順序）
"""
import base64
import io

import streamlit as st
from PIL import Image, ImageOps

try:
    import streamlit.components.v2 as _cv2
    PREVIEW_OK = True
except Exception:
    _cv2 = None
    PREVIEW_OK = False

_HTML = """<div class="rp"><div class="rp-tip"></div><div class="rp-pages"></div></div>"""

_CSS = """
.rp { font-family: "Noto Sans TC", "Microsoft JhengHei", sans-serif; color: #111; }
.rp-tip { font-size: .85rem; margin: 2px 0 8px; color: var(--st-text-color, #333); min-height: 1.2em; }
.rp-tip b { color: var(--st-primary-color, #ff4b4b); }
.rp-pages { display: flex; flex-direction: column; gap: 14px; }
.rp-page { background: #fff; border: 1px solid #c8c8c8; box-shadow: 0 1px 4px rgba(0,0,0,.12);
           padding: 10px; max-width: 760px; margin: 0 auto; width: 100%; box-sizing: border-box; }
.rp-pno { text-align: right; font-size: .75rem; color: #888; margin-bottom: 4px; }
.rp-head { width: 100%; border-collapse: collapse; font-size: .78rem; margin-bottom: 6px; }
.rp-head td { border: 1px solid #555; padding: 3px 5px; }
.rp-head td.k { width: 18%; background: #f3f3f3; text-align: center; white-space: nowrap; }
.rp-title { text-align: center; letter-spacing: .5em; font-weight: 700; font-size: .9rem; margin: 4px 0 6px; }
.rp-grid { display: grid; grid-template-columns: 1fr 1fr; border-top: 1px solid #555; border-left: 1px solid #555; }
.rp-cell { border-right: 1px solid #555; border-bottom: 1px solid #555; position: relative;
           -webkit-user-select: none; user-select: none; -webkit-touch-callout: none; }
.rp-img { aspect-ratio: 4 / 3; background: #eee; display: flex; align-items: center; justify-content: center;
          overflow: hidden; cursor: grab; }
.rp-img img { width: 100%; height: 100%; object-fit: contain; pointer-events: none; -webkit-user-drag: none; }
.rp-info { font-size: .74rem; line-height: 1.35; padding: 4px 5px; white-space: pre-wrap; min-height: 3.6em;
           border-top: 1px solid #555; word-break: break-all; }
.rp-badge { position: absolute; top: 4px; left: 4px; background: rgba(0,0,0,.65); color: #fff;
            font-size: .72rem; padding: 1px 6px; border-radius: 9px; pointer-events: none; }
.rp-board { position: absolute; top: 4px; right: 4px; background: #1f6feb; color: #fff;
            font-size: .7rem; padding: 1px 6px; border-radius: 9px; pointer-events: none; }
.rp-cell.sel .rp-img { outline: 4px solid var(--st-primary-color, #ff4b4b); outline-offset: -4px; }
.rp-cell.over { background: #fff3c4; }
.rp-cell.over .rp-img { outline: 4px dashed #f59f00; outline-offset: -4px; }
.rp-cell.dragging { opacity: .35; }
.rp-cell.empty .rp-img { cursor: default; background: #fafafa; }
.rp-ghost { position: fixed; z-index: 9999; width: 120px; pointer-events: none; opacity: .85;
            box-shadow: 0 6px 18px rgba(0,0,0,.35); border: 2px solid var(--st-primary-color, #ff4b4b);
            border-radius: 4px; transform: translate(-50%, -60%); }
"""

_JS = r"""
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

export default function (component) {
  const { data, setStateValue, parentElement } = component;
  const root = parentElement;
  const pagesEl = root.querySelector('.rp-pages');
  const tip = root.querySelector('.rp-tip');
  const rootNode = root.getRootNode ? root.getRootNode() : document;
  const pick = (x, y) => (rootNode.elementFromPoint ? rootNode.elementFromPoint(x, y) : document.elementFromPoint(x, y));

  const cells = data.cells || [];
  const h = data.header || {};
  let order = cells.map((c) => c.id);
  const byId = Object.fromEntries(cells.map((c) => [c.id, c]));
  const locked = !!data.locked;
  const modeTxt = locked ? '只換照片（說明留在原位）' : '照片和說明一起換';
  const baseTip = `拖拉照片到另一張上即可交換（手機：長按再拖，或點一張再點另一張）｜目前模式：<b>${modeTxt}</b>`;
  tip.innerHTML = baseTip;

  // 每格顯示的文字：鎖定時文字跟「位置」，不鎖定時跟「照片」
  const slotInfo = cells.map((c) => c.info);

  function render() {
    const per = 4;
    const pages = Math.max(1, Math.ceil(order.length / per));
    let html = '';
    for (let p = 0; p < pages; p++) {
      html += `<div class="rp-page"><div class="rp-pno">第 ${p + 1} / ${pages} 頁</div>
        <table class="rp-head">
          <tr><td class="k">工程名稱</td><td colspan="3">${esc(h.project)}</td></tr>
          <tr><td class="k">施工廠商</td><td>${esc(h.contractor)}</td><td class="k">協力廠商</td><td>${esc(h.sub)}</td></tr>
          <tr><td class="k">施作位置</td><td colspan="3">${esc(h.location)}</td></tr>
          <tr><td class="k">自檢項目</td><td colspan="3">${esc(h.item)}</td></tr>
        </table><div class="rp-title">檢 查 照 片</div><div class="rp-grid">`;
      for (let k = 0; k < per; k++) {
        const pos = p * per + k;
        if (pos >= order.length) { html += `<div class="rp-cell empty"><div class="rp-img"></div><div class="rp-info"></div></div>`; continue; }
        const c = byId[order[pos]];
        const info = locked ? slotInfo[pos] : c.info;
        html += `<div class="rp-cell" data-pos="${pos}">
            <div class="rp-img"><img src="data:image/jpeg;base64,${c.thumb}"></div>
            <span class="rp-badge">No.${String(pos + 1).padStart(2, '0')}</span>
            ${c.board ? '<span class="rp-board">白板</span>' : ''}
            <div class="rp-info">${esc(info)}</div></div>`;
      }
      html += `</div></div>`;
    }
    pagesEl.innerHTML = html;
  }

  function commit(a, b) {
    if (a === b) return;
    [order[a], order[b]] = [order[b], order[a]];
    render();
    tip.innerHTML = `⏳ 已交換 No.${String(a + 1).padStart(2, '0')} ↔ No.${String(b + 1).padStart(2, '0')}，同步中…`;
    setStateValue('order', { ids: order.slice(), n: Date.now() + Math.random() });
  }

  render();

  // ---------- 點一張、再點另一張 ----------
  let selected = null;
  const clearSel = () => { root.querySelectorAll('.rp-cell.sel').forEach((e) => e.classList.remove('sel')); selected = null; };
  const cellAt = (x, y) => { const el = pick(x, y); return el && el.closest ? el.closest('.rp-cell[data-pos]') : null; };

  function tapCell(cell) {
    const pos = +cell.dataset.pos;
    if (selected === null) {
      selected = pos; cell.classList.add('sel');
      tip.innerHTML = `已選 No.${String(pos + 1).padStart(2, '0')}，再點另一張即可交換（點同一張取消）`;
    } else if (selected === pos) {
      clearSel(); tip.innerHTML = baseTip;
    } else {
      const a = selected; clearSel(); commit(a, pos);
    }
  }

  // ---------- 拖拉（滑鼠＋手機長按） ----------
  let drag = null;   // {from, ghost, overEl, srcEl}
  let press = null;  // {x, y, timer, cell, touch}

  function startDrag(cell, x, y) {
    clearSel();
    const img = cell.querySelector('img');
    const ghost = img.cloneNode(); ghost.className = 'rp-ghost';
    root.appendChild(ghost);
    cell.classList.add('dragging');
    drag = { from: +cell.dataset.pos, ghost, overEl: null, srcEl: cell };
    if (navigator.vibrate) { try { navigator.vibrate(30); } catch (e) {} }
    moveDrag(x, y);
  }
  function moveDrag(x, y) {
    drag.ghost.style.left = x + 'px'; drag.ghost.style.top = y + 'px';
    const over = cellAt(x, y);
    if (drag.overEl && drag.overEl !== over) drag.overEl.classList.remove('over');
    if (over && over !== drag.srcEl) { over.classList.add('over'); drag.overEl = over; } else drag.overEl = null;
  }
  function endDrag(x, y) {
    const over = cellAt(x, y);
    drag.ghost.remove();
    drag.srcEl.classList.remove('dragging');
    if (drag.overEl) drag.overEl.classList.remove('over');
    const from = drag.from; drag = null;
    if (over && +over.dataset.pos !== from) commit(from, +over.dataset.pos);
  }

  // 滑鼠
  pagesEl.onmousedown = (e) => {
    const cell = e.target.closest('.rp-cell[data-pos]');
    if (!cell || e.button !== 0 || !e.target.closest('.rp-img')) return;
    e.preventDefault();
    press = { x: e.clientX, y: e.clientY, cell, touch: false };
    const mm = (ev) => {
      if (!drag && press && Math.hypot(ev.clientX - press.x, ev.clientY - press.y) > 6) startDrag(press.cell, ev.clientX, ev.clientY);
      if (drag) moveDrag(ev.clientX, ev.clientY);
    };
    const mu = (ev) => {
      window.removeEventListener('mousemove', mm); window.removeEventListener('mouseup', mu);
      if (drag) endDrag(ev.clientX, ev.clientY); else if (press) tapCell(press.cell);
      press = null;
    };
    window.addEventListener('mousemove', mm); window.addEventListener('mouseup', mu);
  };

  // 手機：長按 350ms 開始拖；還沒開始前手指移動＝捲動頁面
  pagesEl.ontouchstart = (e) => {
    const cell = e.target.closest('.rp-cell[data-pos]');
    if (!cell || !e.target.closest('.rp-img') || e.touches.length > 1) return;
    const t = e.touches[0];
    press = { x: t.clientX, y: t.clientY, cell, touch: true, moved: false,
              timer: setTimeout(() => { if (press && !press.moved) startDrag(cell, press.x, press.y); }, 350) };
  };
  pagesEl.addEventListener('touchmove', (e) => {
    const t = e.touches[0];
    if (drag) { e.preventDefault(); moveDrag(t.clientX, t.clientY); return; }
    if (press && Math.hypot(t.clientX - press.x, t.clientY - press.y) > 10) { press.moved = true; clearTimeout(press.timer); }
  }, { passive: false });
  pagesEl.ontouchend = (e) => {
    if (press) clearTimeout(press.timer);
    const t = e.changedTouches[0];
    if (drag) { e.preventDefault(); endDrag(t.clientX, t.clientY); }
    else if (press && !press.moved) { e.preventDefault(); tapCell(press.cell); }
    press = null;
  };
  pagesEl.ontouchcancel = () => {
    if (press) clearTimeout(press.timer);
    if (drag) { drag.ghost.remove(); drag.srcEl.classList.remove('dragging'); if (drag.overEl) drag.overEl.classList.remove('over'); drag = null; }
    press = null;
  };
}
"""

_component = None
if PREVIEW_OK:
    try:
        _component = _cv2.component("site_report_preview", html=_HTML, css=_CSS, js=_JS)
    except Exception:
        PREVIEW_OK = False


def preview_thumb(photo, max_side=360):
    """預覽用小縮圖（base64），每張照片只算一次"""
    if photo.get("pthumb"):
        return photo["pthumb"]
    try:
        img = Image.open(io.BytesIO(photo["file"].getvalue()))
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.thumbnail((max_side, max_side))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=70)
        photo["pthumb"] = base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        photo["pthumb"] = ""
    return photo["pthumb"]


def report_preview(key, header, cells, locked):
    """
    header: {project, contractor, sub, location, item}
    cells:  [{id, thumb(base64), info(文字), board(bool)}]，依目前順序
    回傳新的照片 id 順序（list）；沒有變動回傳 None
    """
    if not PREVIEW_OK or _component is None:
        return None
    seen_key = f"_rp_seen_{key}"
    result = _component(
        key=key,
        data={"header": header, "cells": cells, "locked": bool(locked)},
        default={"order": None},
        on_order_change=lambda: None,
    )
    order = getattr(result, "order", None)
    if not order or not isinstance(order, dict):
        return None
    n = order.get("n")
    if n is None or n == st.session_state.get(seen_key):
        return None
    st.session_state[seen_key] = n
    ids = [str(i) for i in (order.get("ids") or [])]
    current = [c["id"] for c in cells]
    if sorted(ids) != sorted(current) or ids == current:
        return None
    return ids
