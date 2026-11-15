/**
 * iFix Laptop - Trang tải ảnh máy từ điện thoại.
 * Chọn máy -> chụp/chọn ảnh -> bấm Tải lên. Ảnh vào thư mục "iFix Laptop - Ảnh máy (cho web)"
 * và link ảnh tự điền vào cột "Ảnh" đúng dòng máy đó trong bảng. Web tự hiện ảnh mới.
 *
 * Cài một lần (trên máy tính): mở bảng Google Sheets -> Tiện ích mở rộng -> Apps Script ->
 * xóa hết code mẫu, dán file này vào -> Lưu -> Triển khai -> Tùy chọn triển khai mới -> Ứng dụng web
 * (Thực thi với tư cách: Tôi; Người có quyền truy cập: Chỉ mình tôi) -> Triển khai -> cấp quyền.
 * Copy link "Ứng dụng web" mở trên điện thoại, bấm "Thêm vào màn hình chính".
 */
const SHEET_ID = "1fmLC2Z9kXulfQtPzyPAt3XamGivs5YjCyfQZmpS5yUI";
const FOLDER_ID = "1h-IloR6VuXPje1DXB2jGnDXXzcz1q75J";

function key_(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/đ/gi, "d").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function sheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function cols_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(key_);
  const name = head.indexOf("tenmay");
  let img = head.indexOf("anh");
  if (img < 0) img = head.indexOf("hinhanh");
  if (name < 0) throw new Error("Không thấy cột 'Tên máy' trong bảng.");
  if (img < 0) { img = head.length; sh.getRange(1, img + 1).setValue("Ảnh"); }
  return { name: name, img: img };
}

/** Danh sách máy cho ô chọn */
function listMachines() {
  const sh = sheet_(), c = cols_(sh), n = sh.getLastRow();
  if (n < 2) return [];
  const rows = sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
  return rows.map((r, i) => ({ row: i + 2, name: String(r[c.name]).trim(), hasImg: !!String(r[c.img] || "").trim() }))
             .filter(m => m.name);
}

/** Nhận ảnh từ điện thoại, lưu vào Drive, điền link vào cột Ảnh */
function uploadPhoto(row, name, dataUrl) {
  const sh = sheet_(), c = cols_(sh);
  // Dòng có thể đã dời (thêm/xóa máy trong lúc chọn ảnh): tìm lại theo tên máy
  let r = row;
  if (String(sh.getRange(r, c.name + 1).getValue()).trim() !== name) {
    const list = listMachines().filter(m => m.name === name);
    if (!list.length) throw new Error("Không tìm thấy máy " + name + " trong bảng.");
    r = list[0].row;
  }
  const m = String(dataUrl).match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!m) throw new Error("File không phải ảnh.");
  const stamp = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd_HHmm");
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name.replace(/[^\w\- ]+/g, "") + " " + stamp + ".jpg");
  const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const link = "https://drive.google.com/file/d/" + file.getId() + "/view";
  sh.getRange(r, c.img + 1).setValue(link);
  return { name: name, link: link, thumb: "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w400" };
}

function doGet() {
  return HtmlService.createHtmlOutput(PAGE)
    .setTitle("iFix - Tải ảnh máy")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

const PAGE = `<!doctype html><html><head><meta charset="utf-8">
<style>
  body { font-family: system-ui, sans-serif; margin: 0; padding: 16px; background: #f1f5f9; color: #0f172a; }
  h1 { font-size: 1.25rem; margin: 4px 0 16px; }
  label { display: block; font-weight: 600; margin: 14px 0 6px; }
  select, button { width: 100%; font-size: 1.05rem; padding: 14px; border-radius: 12px; border: 1px solid #cbd5e1; background: #fff; }
  .pick { display: block; text-align: center; padding: 22px; border: 2px dashed #94a3b8; border-radius: 12px; background: #fff; font-size: 1.05rem; }
  input[type=file] { display: none; }
  #preview { width: 100%; border-radius: 12px; margin-top: 10px; display: none; }
  button { margin-top: 18px; background: #2563eb; color: #fff; border: 0; font-weight: 700; }
  button:disabled { background: #94a3b8; }
  #msg { margin-top: 14px; padding: 12px; border-radius: 10px; display: none; }
  .ok { background: #dcfce7; color: #166534; } .err { background: #fee2e2; color: #991b1b; }
</style></head><body>
<h1>📷 Tải ảnh máy lên web</h1>
<label for="may">1. Chọn máy</label>
<select id="may"><option value="">Đang tải danh sách...</option></select>
<label>2. Chụp hoặc chọn ảnh</label>
<label class="pick" for="file">📸 Bấm để chụp / chọn ảnh</label>
<input id="file" type="file" accept="image/*">
<img id="preview" alt="">
<button id="go" disabled>⬆️ Tải lên</button>
<div id="msg"></div>
<script>
  const sel = document.getElementById("may"), fileIn = document.getElementById("file"),
        prev = document.getElementById("preview"), go = document.getElementById("go"), msg = document.getElementById("msg");
  let data = null;
  const say = (t, ok) => { msg.textContent = t; msg.className = ok ? "ok" : "err"; msg.style.display = "block"; };
  const check = () => go.disabled = !(sel.value && data);
  function load() {
    google.script.run.withSuccessHandler(list => {
      sel.innerHTML = '<option value="">-- Chọn máy --</option>' + list.map(m =>
        '<option value="' + m.row + '">' + m.name.replace(/</g, "&lt;") + (m.hasImg ? " ✓" : "") + '</option>').join("");
    }).withFailureHandler(e => say("Lỗi tải danh sách: " + e.message)).listMachines();
  }
  load();
  sel.onchange = check;
  // Thu nhỏ ảnh còn tối đa 1400px để tải nhanh bằng 4G
  fileIn.onchange = () => {
    const f = fileIn.files[0]; if (!f) return;
    const img = new Image();
    img.onload = () => {
      const s = Math.min(1, 1400 / Math.max(img.width, img.height));
      const c = document.createElement("canvas"); c.width = Math.round(img.width * s); c.height = Math.round(img.height * s);
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      data = c.toDataURL("image/jpeg", 0.85); prev.src = data; prev.style.display = "block"; URL.revokeObjectURL(img.src); check();
    };
    img.onerror = () => say("Không đọc được ảnh này, thử ảnh khác nhé.");
    img.src = URL.createObjectURL(f);
  };
  go.onclick = () => {
    const name = sel.options[sel.selectedIndex].text.replace(/ ✓$/, "");
    go.disabled = true; go.textContent = "Đang tải lên...";
    google.script.run.withSuccessHandler(r => {
      say("✅ Đã gắn ảnh cho " + r.name + ". Khách mở lại web là thấy ảnh mới.", true);
      go.textContent = "⬆️ Tải lên"; data = null; fileIn.value = ""; prev.style.display = "none"; load(); check();
    }).withFailureHandler(e => { say("Lỗi: " + e.message); go.textContent = "⬆️ Tải lên"; check(); })
      .uploadPhoto(+sel.value, name, data);
  };
</script></body></html>`;
