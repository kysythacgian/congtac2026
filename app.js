const { Document, Packer, Paragraph, TextRun, AlignmentType, TabStopType, TabStopPosition } = docx;

// ===================== TOAST NOTIFICATION =====================
function showToast(title, message, type = 'info', duration = 3000) {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  
  const icons = {
    success: '✅',
    error: '❌',
    warning: '⚠️',
    info: 'ℹ️'
  };
  
  toast.innerHTML = `
    <div class="toast-icon">${icons[type] || icons.info}</div>
    <div class="toast-content">
      <div class="toast-title">${title}</div>
      <div class="toast-message">${message}</div>
    </div>
    <button class="toast-close" onclick="this.closest('.toast').remove()">×</button>
    <div class="toast-progress">
      <div class="toast-progress-bar"></div>
    </div>
  `;
  
  container.appendChild(toast);
  
  const timeout = setTimeout(() => {
    if (toast && toast.parentNode) {
      toast.classList.add('hide');
      setTimeout(() => toast.remove(), 300);
    }
  }, duration);
  
  const closeBtn = toast.querySelector('.toast-close');
  closeBtn.onclick = () => {
    clearTimeout(timeout);
    toast.classList.add('hide');
    setTimeout(() => toast.remove(), 300);
  };
  
  return toast;
}

// ===================== CONFIRM MODAL =====================
function showConfirm({ icon, title, message, confirmText, confirmClass, onConfirm, onCancel }) {
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.innerHTML = `
    <div class="modal-box">
      <div class="modal-icon">${icon}</div>
      <div class="modal-title">${title}</div>
      <div class="modal-msg">${message}</div>
      <div class="modal-actions">
        <button class="modal-btn modal-btn-cancel" id="modalCancel">✖ Hủy</button>
        <button class="modal-btn ${confirmClass}" id="modalConfirm">${confirmText}</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  overlay.querySelector('#modalCancel').onclick = () => {
    overlay.remove();
    if (onCancel) onCancel();
  };
  overlay.querySelector('#modalConfirm').onclick = () => { 
    overlay.remove(); 
    if (onConfirm) onConfirm();
  };
  overlay.addEventListener('click', e => { 
    if (e.target === overlay) {
      overlay.remove();
      if (onCancel) onCancel();
    }
  });
}

// ===================== SET BUTTON LOADING STATE =====================
function setButtonLoading(btn, isLoading, originalText) {
  if (isLoading) {
    btn.disabled = true;
    btn.style.opacity = '0.7';
    btn.innerHTML = '<span>⏳</span> Đang xử lý...';
  } else {
    btn.disabled = false;
    btn.style.opacity = '1';
    btn.innerHTML = originalText;
  }
}


// ===================== CẤU HÌNH BÁO CÁO 2026 =====================
const KY_BAO_CAO = 'từ tháng 7/2026 đến tháng 11/2026';
const PHUONG = 'THANH KHÊ';
const MUC = [
  { id: 'thamMuu',      title: '1. Tham mưu cấp ủy' },
  { id: 'phoiHop',      title: '2. Phối hợp Tổ dân phố, các hội đoàn thể khu dân cư' },
  { id: 'moHinhMoi',    title: '3. Mô hình mới của Chi hội' },
  { id: 'ketQuaHoi',    title: '4. Kết quả công tác Hội', note: '(Cuộc vận động “5 không, 3 sạch, 3 an”; bảo vệ môi trường; an sinh xã hội; Mẹ đỡ đầu; chuyển đổi số…)' },
  { id: 'tuyenTruyen',  title: '5. Công tác tuyên truyền của chi hội', note: '(Tổ chức các điểm tuyên truyền…)' },
  { id: 'hoatDong2010', title: '6. Hoạt động 20/10/2026' },
  { id: 'hoiCapTren',   title: '7. Tham gia hoạt động Hội cấp trên' },
  { id: 'deXuat',       title: '8. Đề xuất' },
];
const DOTS = '..........................................................................................................................................';

function safeName(s) {
  return s.replace(/[đĐ]/g, 'd').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
}
function getFileName(chiHoi) { return `CTH_BC2026_${safeName(chiHoi || 'ChiHoi')}.docx`; }

function checkForm() {
  if (!document.getElementById('chiHoi').value.trim()) {
    showToast('Thiếu thông tin', 'Vui lòng chọn chi hội.', 'warning');
    return false;
  }
  return true;
}

// ===================== HANDLE DOWNLOAD BUTTON =====================
function handleDownload() {
  if (!checkForm()) return;
  const fileName = getFileName(document.getElementById('chiHoi').value.trim().toUpperCase());
  showConfirm({
    icon: '📄',
    title: 'Xác nhận tải file Word',
    message: `Bạn sẽ tải xuống file:<br><b>${fileName}</b>`,
    confirmText: '⬇ Tải xuống',
    confirmClass: 'modal-btn-confirm-blue',
    onConfirm: () => {
      const btn = document.querySelector('.btn-word');
      const originalText = btn.innerHTML;
      setButtonLoading(btn, true, originalText);
      createDoc(false)
        .then(() => showToast('Thành công', 'Đã lưu tại thư mục Download', 'success'))
        .catch(err => showToast('Lỗi', err.message || 'Không thể tải file', 'error'))
        .finally(() => setButtonLoading(btn, false, originalText));
    }
  });
}

// ===================== HANDLE SEND BUTTON =====================
function handleSend() {
  const btn = document.getElementById('sendBtn');
  if (btn.disabled) return;
  if (!checkForm()) return;
  showConfirm({
    icon: '☁️',
    title: 'Xác nhận gửi báo cáo',
    message: 'Gửi báo cáo kết quả công tác Hội năm 2026 về Hội LHPN phường?<br><small>Nếu chi hội đã gửi trước đó, báo cáo mới sẽ thay thế báo cáo cũ.</small>',
    confirmText: '✔ Gửi ngay',
    confirmClass: 'modal-btn-confirm-cyan',
    onConfirm: () => {
      const originalText = btn.innerHTML;
      setButtonLoading(btn, true, originalText);
      createDoc(true)
        .then(() => showToast('Gửi thành công', 'Báo cáo đã được gửi lên phường!', 'success'))
        .catch(err => showToast('Gửi thất bại', err.message || 'Không thể gửi báo cáo', 'error'))
        .finally(() => setButtonLoading(btn, false, originalText));
    }
  });
}

// ===================== UPLOAD LÊN GOOGLE DRIVE + GHI BẢNG TỔNG HỢP =====================
function uploadFile(blob, fileName, data) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Không đọc được file'));
    reader.onloadend = function () {
      const base64 = reader.result.split(',')[1];
      fetch('/api/proxy', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file: base64, filename: fileName, data })
      })
        .then(r => r.json())
        .then(res => res.success ? resolve() : reject(new Error(res.error || 'Upload failed')))
        .catch(reject);
    };
    reader.readAsDataURL(blob);
  });
}

// ===================== TẠO FILE WORD =====================
async function createDoc(isUpload) {
  const chiHoi = document.getElementById('chiHoi').value.trim().toUpperCase() || '.................';
  const values = {};
  MUC.forEach(m => { values[m.id] = document.getElementById(m.id).value.trim(); });

  const today = new Date();
  const dateStr = `${PHUONG}, ngày ${today.getDate()} tháng ${today.getMonth() + 1} năm ${today.getFullYear()}`;

  const bold = (text, extra = {}) => new Paragraph({ children: [new TextRun({ text, bold: true })], spacing: { after: 60 }, ...extra });
  const content = (text) => {
    const lines = (text || DOTS).split(/\r?\n/);
    return lines.map((ln, i) => new Paragraph({
      children: [new TextRun({ text: ln })],
      spacing: { after: i === lines.length - 1 ? 150 : 40 }
    }));
  };

  const body = [];
  MUC.forEach(m => {
    body.push(bold(m.title, { spacing: { before: 120, after: 60 } }));
    if (m.note) body.push(new Paragraph({ children: [new TextRun({ text: m.note, italics: true })], spacing: { after: 60 } }));
    body.push(...content(values[m.id]));
  });

  const doc = new Document({
    styles: { default: { document: { run: { font: 'Times New Roman', size: 26 }, paragraph: { spacing: { line: 360 } } } } },
    sections: [{
      properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
      children: [
        new Paragraph({
          tabStops: [{ type: TabStopType.CENTER, position: 6500 }],
          children: [
            new TextRun({ text: 'HỘI LHPN PHƯỜNG ' + PHUONG, bold: true }),
            new TextRun({ text: '\t' }),
            new TextRun({ text: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', bold: true }),
          ],
        }),
        new Paragraph({
          tabStops: [{ type: TabStopType.CENTER, position: 6500 }],
          children: [
            new TextRun({ text: (/^CHI HỘI/.test(chiHoi) ? '' : 'CHI HỘI ') + chiHoi, bold: true }),
            new TextRun({ text: '\t' }),
            new TextRun({ text: 'Độc lập - Tự do - Hạnh phúc', bold: true, underline: {} }),
          ],
        }),
        new Paragraph({
          tabStops: [{ type: TabStopType.CENTER, position: 6500 }],
          spacing: { after: 200 },
          children: [new TextRun({ text: '\t' }), new TextRun({ text: dateStr, italics: true })],
        }),
        new Paragraph({ children: [new TextRun({ text: 'BÁO CÁO', bold: true, size: 32 })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
        new Paragraph({ children: [new TextRun({ text: 'Kết quả công tác Hội năm 2026', bold: true })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
        new Paragraph({ children: [new TextRun({ text: `(Thời gian: ${KY_BAO_CAO})`, italics: true })], alignment: AlignmentType.CENTER, spacing: { after: 300 } }),
        ...body,
        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 300, after: 200 }, children: [new TextRun({ text: 'TM. CHI HỘI', bold: true })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: chiHoi, bold: true })] }),
      ]
    }]
  });

  const blob = await Packer.toBlob(doc);
  const fileName = getFileName(chiHoi);
  if (isUpload) {
    await uploadFile(blob, fileName, { chiHoi, ...values });
  } else {
    await downloadFile(blob, fileName);
  }
}

// ===================== TẢI FILE (HỖ TRỢ CẢ ANDROID VÀ TRÌNH DUYỆT) =====================
function downloadFile(blob, fileName) {
  return new Promise((resolve, reject) => {
    try {
      // --- Android WebView: dùng bridge Java (AndroidDownload) ---
      const isAndroidBridge = typeof AndroidDownload !== 'undefined';
      if (isAndroidBridge) {
        const reader = new FileReader();
        reader.onload = function(e) {
          try {
            const mimeType = blob.type || 'application/octet-stream';
            AndroidDownload.saveFile(e.target.result, fileName, mimeType);
            resolve();
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = () => reject(new Error('Không đọc được file'));
        reader.readAsDataURL(blob);
        return;
      }

      // --- Trình duyệt thường: dùng FileSaver hoặc thẻ <a> ---
      if (typeof saveAs !== 'undefined') {
        saveAs(blob, fileName);
        resolve();
      } else {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
          resolve();
        }, 2000);
      }
    } catch (e) {
      // --- Fallback cuối: mở data URL ---
      try {
        const reader = new FileReader();
        reader.onload = function(ev) {
          window.location.href = ev.target.result;
          resolve();
        };
        reader.onerror = () => reject(new Error('Không thể tải file'));
        reader.readAsDataURL(blob);
      } catch (err) {
        reject(err);
      }
    }
  });
}

// ===================== TỰ ĐỘNG VIẾT HOA TÊN NGƯỜI KÝ =====================
function setupAutoUppercase() {
  const chiHoiTruongInput = document.getElementById('chiHoiTruong');
  if (chiHoiTruongInput) {
    chiHoiTruongInput.addEventListener('input', function() {
      const start = this.selectionStart;
      const end = this.selectionEnd;
      this.value = this.value.toUpperCase();
      this.setSelectionRange(start, end);
    });
  }
}

// ===================== LOAD DANH SÁCH CHI HỘI =====================
window.addEventListener("DOMContentLoaded", () => {
  setupAutoUppercase();

  const list = document.getElementById("chiHoiList");

  function renderList(filter = "") {
    list.innerHTML = "";
    let count = 0;
    if (typeof DS_CHI_HOI_GS !== 'undefined' && Array.isArray(DS_CHI_HOI_GS)) {
      DS_CHI_HOI_GS.forEach(name => {
        if (!filter || name.includes(filter.toUpperCase())) {
          const li = document.createElement("li");
          li.textContent = name;
          li.addEventListener("mousedown", (e) => {
            e.preventDefault();
            selectChiHoi(name);
          });
          list.appendChild(li);
          count++;
        }
      });
    }
    document.getElementById("chiHoiNoResult").style.display = count === 0 ? "block" : "none";
  }

  renderList();

  document.getElementById("chiHoiSearch").addEventListener("input", function () {
    renderList(this.value);
  });

  document.addEventListener("click", function (e) {
    if (!document.getElementById("chiHoiWrapper").contains(e.target)) {
      closeChiHoiDropdown();
    }
  });
});

function toggleChiHoiDropdown() {
  const dropdown = document.getElementById("chiHoiDropdown");
  const display = document.getElementById("chiHoiDisplay");
  const isOpen = dropdown.style.display !== "none";
  if (isOpen) {
    closeChiHoiDropdown();
  } else {
    dropdown.style.display = "block";
    display.classList.add("open");
    setTimeout(() => document.getElementById("chiHoiSearch").focus(), 50);
  }
}

function closeChiHoiDropdown() {
  document.getElementById("chiHoiDropdown").style.display = "none";
  document.getElementById("chiHoiDisplay").classList.remove("open");
  document.getElementById("chiHoiSearch").value = "";
  const list = document.getElementById("chiHoiList");
  list.innerHTML = "";
  if (typeof DS_CHI_HOI_GS !== 'undefined') {
    DS_CHI_HOI_GS.forEach(name => {
      const li = document.createElement("li");
      li.textContent = name;
      li.addEventListener("mousedown", (e) => { e.preventDefault(); selectChiHoi(name); });
      list.appendChild(li);
    });
  }
  document.getElementById("chiHoiNoResult").style.display = "none";
}

function selectChiHoi(name) {
  document.getElementById("chiHoi").value = name;
  const textEl = document.getElementById("chiHoiText");
  textEl.textContent = name;
  textEl.classList.remove("placeholder");
  document.getElementById("chiHoiDisplay").classList.add("selected");
  closeChiHoiDropdown();
}
