"""
手機友善的照片上傳元件

為什麼不用 st.file_uploader：
  手機打開相簿時瀏覽器會被切到背景，App 的連線會斷掉；選完照片回來時連線還在重連，
  st.file_uploader 卻馬上開始上傳，伺服器會拒收「已斷線的連線」送來的檔案，照片就消失了。

這個元件的做法：
  1. 在手機上先把照片縮小成 JPEG（6MB → 約 0.3～0.6MB），傳得快
  2. 透過 App 本身的連線送出，連線斷掉時會自動重送，直到伺服器回覆「收到了」
  3. 讀出照片裡的拍攝時間一起送，照片排序不受縮圖影響
"""
import base64
import datetime

import streamlit as st

try:
    import streamlit.components.v2 as _cv2
    PICKER_OK = True
except Exception:  # 舊版 Streamlit 沒有 components.v2
    _cv2 = None
    PICKER_OK = False

_HTML = """
<div class="pp">
  <button type="button" class="pp-btn">📷 選擇照片</button>
  <input type="file" class="pp-input" accept="image/*" multiple hidden>
  <div class="pp-msg"></div>
</div>
"""

_CSS = """
.pp { font-family: inherit; }
.pp-btn {
  width: 100%; padding: 14px 12px; font-size: 1.05rem; font-weight: 600;
  border-radius: 10px; border: 2px dashed var(--st-primary-color, #ff4b4b);
  background: transparent; color: var(--st-text-color, inherit); cursor: pointer;
}
.pp-btn:disabled { opacity: .5; cursor: wait; }
.pp-msg { margin-top: 6px; font-size: .9rem; min-height: 1.2em; color: var(--st-text-color, inherit); }
"""

_JS = r"""
const toB64 = (blob) => new Promise((res, rej) => {
  const fr = new FileReader();
  fr.onload = () => res(String(fr.result).split(',')[1]);
  fr.onerror = rej;
  fr.readAsDataURL(blob);
});

// 從 JPEG 開頭找拍攝時間（EXIF 的 "YYYY:MM:DD HH:MM:SS"）
async function readExifTime(file) {
  try {
    const buf = await file.slice(0, 131072).arrayBuffer();
    const txt = new TextDecoder('latin1').decode(new Uint8Array(buf));
    const m = txt.match(/(19|20)\d\d:[01]\d:[0-3]\d [0-2]\d:[0-5]\d:[0-5]\d/);
    return m ? m[0] : '';
  } catch (e) { return ''; }
}

async function shrink(file, maxSide) {
  const base = { name: file.name || 'photo.jpg', last_modified: file.lastModified || 0,
                 exif_time: await readExifTime(file) };
  try {
    const bmp = await createImageBitmap(file, { imageOrientation: 'from-image' });
    const s = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
    const c = document.createElement('canvas');
    c.width = Math.round(bmp.width * s); c.height = Math.round(bmp.height * s);
    c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
    if (bmp.close) bmp.close();
    const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', 0.85));
    const name = base.name.replace(/\.[^.]+$/, '') + '.jpg';
    return { ...base, name, b64: await toB64(blob) };
  } catch (e) {
    // 瀏覽器解不開（例如部分 HEIC）→ 傳原檔，交給伺服器轉
    return { ...base, b64: await toB64(file) };
  }
}

export default function (component) {
  const { data, setStateValue, parentElement } = component;
  const root = parentElement;
  const uid = data.uid;
  window.__photoPicker = window.__photoPicker || {};
  const S = (window.__photoPicker[uid] = window.__photoPicker[uid] ||
             { pending: null, tries: 0, timer: null, busy: false, msg: '' });
  S.set = setStateValue;                       // 每次重繪都換成最新的 setStateValue

  const btn = root.querySelector('.pp-btn');
  const input = root.querySelector('.pp-input');
  const msgEl = root.querySelector('.pp-msg');
  const show = (t) => { S.msg = t; msgEl.textContent = t; };
  msgEl.textContent = S.msg;
  btn.disabled = S.busy;

  // 伺服器回覆收到 → 停止重送
  if (S.pending && data.ack === S.pending.id) {
    show(`✅ 已上傳 ${S.pending.files.length} 張`);
    S.pending = null;
    clearInterval(S.timer); S.timer = null;
    S.busy = false; btn.disabled = false;
  }

  const send = () => {
    if (!S.pending) return;
    S.tries += 1;
    if (S.tries > 60) {                         // 約 5 分鐘還送不到就放棄
      clearInterval(S.timer); S.timer = null;
      S.busy = false; btn.disabled = false;
      show('❌ 一直連不上伺服器，請確認網路後重新選擇照片');
      S.pending = null;
      return;
    }
    if (S.tries > 1) show(`📶 等待連線，自動重送中…（第 ${S.tries} 次）`);
    S.set('batch', { ...S.pending, attempt: S.tries });
  };

  btn.onclick = () => { if (!S.busy) input.click(); };
  input.onchange = async () => {
    const files = Array.from(input.files || []);
    input.value = '';
    if (!files.length) return;
    S.busy = true; btn.disabled = true;
    const out = [];
    for (let i = 0; i < files.length; i++) {
      show(`🗜️ 處理照片中 ${i + 1}/${files.length}…`);
      out.push(await shrink(files[i], data.max_side || 2000));
    }
    S.pending = { id: Date.now() + '-' + Math.random().toString(36).slice(2, 8), files: out };
    S.tries = 0;
    show(`⬆️ 傳送 ${out.length} 張中…`);
    send();
    clearInterval(S.timer);
    S.timer = setInterval(send, 5000);
  };
}
"""

_component = None
if PICKER_OK:
    try:
        _component = _cv2.component("site_photo_picker", html=_HTML, css=_CSS, js=_JS)
    except Exception:
        PICKER_OK = False


class PickedPhoto:
    """跟上傳檔一樣的用法：.name / .size / .getvalue()，另外帶著拍照時間"""
    def __init__(self, data, name, ts_info):
        import io
        self._buf = io.BytesIO(data)
        self.name = name
        self.size = len(data)
        self.ts_info = ts_info

    def getvalue(self):
        return self._buf.getvalue()

    def read(self, *a):
        return self._buf.read(*a)

    def seek(self, *a):
        return self._buf.seek(*a)

    def tell(self):
        return self._buf.tell()


def _ts_from_item(it):
    s = str(it.get("exif_time") or "")
    try:
        if s:
            return datetime.datetime.strptime(s, "%Y:%m:%d %H:%M:%S"), 0, "EXIF"
    except ValueError:
        pass
    return None


def photo_picker(key, on_photos, max_side=2000):
    """
    顯示選照片按鈕。收到新的一批照片時呼叫 on_photos([PickedPhoto, ...])，
    「存好之後」才回覆手機收到（中途被打斷的話手機會重送，on_photos 要能處理重複的照片）。
    回傳 True 代表這次有新照片（呼叫端可以 st.rerun 讓手機看到「已上傳」）。
    """
    if not PICKER_OK or _component is None:
        return False
    ack_key = f"_pp_ack_{key}"
    result = _component(
        key=key,
        data={"uid": key, "ack": st.session_state.get(ack_key), "max_side": max_side},
        default={"batch": None},
        on_batch_change=lambda: None,
    )
    batch = getattr(result, "batch", None)
    if not batch or not isinstance(batch, dict):
        return False
    bid = batch.get("id")
    if not bid or bid == st.session_state.get(ack_key):
        return False
    photos = []
    for it in batch.get("files") or []:
        try:
            raw = base64.b64decode(it.get("b64") or "")
        except Exception:
            continue
        if not raw:
            continue
        ts_info = _ts_from_item(it)
        if ts_info is None and it.get("last_modified"):
            # 沒有 EXIF：用手機檔案的修改時間（通常就是拍照時間），換成台灣時間
            try:
                t = (datetime.datetime.fromtimestamp(int(it["last_modified"]) / 1000, datetime.timezone.utc)
                     .replace(tzinfo=None) + datetime.timedelta(hours=8))
                ts_info = (t, 0, "檔案時間")
            except Exception:
                ts_info = None
        photos.append(PickedPhoto(raw, str(it.get("name") or "photo.jpg"), ts_info))
    on_photos(photos)
    st.session_state[ack_key] = bid   # 存好了才回覆收到
    return True
