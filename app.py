import streamlit as st
import streamlit.components.v1 as components

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Dashboard Báo Cáo Vận Hành - Triển Khai - Khai Thác",
    page_icon="📊",
    layout="wide"
)

# Ẩn bớt các padding và menu mặc định của Streamlit để giữ giao diện full-screen đẹp mắt
st.markdown("""
    <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 1rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# Giao diện HTML/CSS/JS gốc
html_code = """
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard Báo Cáo Vận Hành - Triển Khai - Khai Thác</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <style>
    :root {
      --primary-color: #0f172a;
      --secondary-color: #2563eb;
      --accent-color: #f59e0b;
      --bg-color: #f8fafc;
      --card-bg: #ffffff;
      --text-color: #334155;
      --border-color: #e2e8f0;
    }

    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: var(--bg-color);
      color: var(--text-color);
      margin: 0;
      padding: 10px;
    }

    .container {
      max-width: 1400px;
      margin: 0 auto;
    }

    header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      background-color: var(--card-bg);
      padding: 20px 30px;
      border-radius: 12px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
      margin-bottom: 25px;
    }

    h1 {
      margin: 0;
      font-size: 24px;
      color: var(--primary-color);
    }

    .file-input-wrapper {
      position: relative;
    }

    .btn-upload {
      background-color: var(--secondary-color);
      color: white;
      padding: 10px 20px;
      border-radius: 8px;
      cursor: pointer;
      font-weight: 600;
      transition: all 0.3s;
      border: none;
      display: inline-block;
    }

    .btn-upload:hover {
      background-color: #1d4ed8;
    }

    input[type="file"] {
      display: none;
    }

    .section-title {
      font-size: 20px;
      font-weight: 700;
      color: var(--primary-color);
      margin: 30px 0 15px 0;
      padding-bottom: 8px;
      border-bottom: 2px solid var(--secondary-color);
    }

    .grid-container {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
      gap: 20px;
      margin-bottom: 25px;
    }

    .card {
      background-color: var(--card-bg);
      border-radius: 12px;
      padding: 20px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
      border: 1px solid var(--border-color);
    }

    .full-width {
      grid-column: 1 / -1;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 10px;
      font-size: 14px;
    }

    th, td {
      padding: 12px 15px;
      text-align: left;
      border-bottom: 1px solid var(--border-color);
    }

    th {
      background-color: #f1f5f9;
      color: var(--primary-color);
      font-weight: 600;
    }

    tr:hover {
      background-color: #f8fafc;
    }

    .badge {
      padding: 4px 8px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 600;
    }

    .badge-success { background-color: #dcfce7; color: #15803d; }
    .badge-warning { background-color: #fef3c7; color: #b45309; }
    .badge-danger { background-color: #fee2e2; color: #b91c1c; }

    .notes-box {
      background-color: #f1f5f9;
      border-left: 4px solid var(--secondary-color);
      padding: 15px;
      border-radius: 0 8px 8px 0;
      margin-top: 15px;
    }

    .notes-box ul {
      margin: 0;
      padding-left: 20px;
    }

    .notes-box li {
      margin-bottom: 6px;
    }

    .chart-container {
      position: relative;
      height: 320px;
      width: 100%;
    }
  </style>
</head>
<body>

  <div class="container">
    <header>
      <h1>BÁO CÁO TỔNG HỢP VẬN HÀNH - TRIỂN KHAI - KHAI THÁC</h1>
      <div class="file-input-wrapper">
        <label for="excelFile" class="btn-upload">📥 Tải file Excel/Google Sheet (.xlsx)</label>
        <input type="file" id="excelFile" accept=".xlsx, .xls">
      </div>
    </header>

    <!-- PHẦN I: VẬN HÀNH -->
    <div class="section-title">I. QUẢN LÝ VẬN HÀNH KPI</div>
    <div class="grid-container">
      <div class="card full-width">
        <h3>Bảng Chỉ Số Vận Hành</h3>
        <div style="overflow-x: auto;">
          <table id="tableVanHanh">
            <thead></thead>
            <tbody></tbody>
          </table>
        </div>
        <div class="notes-box" id="notesVanHanh">
          <strong>Nhận xét:</strong>
          <ul id="listNotesVanHanh"><li>Chưa có dữ liệu.</li></ul>
        </div>
      </div>

      <div class="card">
        <h3>Biểu Đồ Số Lượng Sự Cố & Khách Hàng Ảnh Hưởng</h3>
        <div class="chart-container">
          <canvas id="chartSuCo"></canvas>
        </div>
      </div>

      <div class="card">
        <h3>Biểu Đồ Thời Gian Xử Lý & Gián Đoạn (Phút)</h3>
        <div class="chart-container">
          <canvas id="chartThoiGian"></canvas>
        </div>
      </div>
    </div>

    <!-- PHẦN II: TRIỂN KHAI -->
    <div class="section-title">II. TRIỂN KHAI HẠ TẦNG</div>
    <div class="grid-container">
      <div class="card full-width">
        <h3>Tiến Độ Triển Khai POP / Hạ Tầng</h3>
        <div style="overflow-x: auto;">
          <table id="tableTrienKhai">
            <thead></thead>
            <tbody></tbody>
          </table>
        </div>
        <div class="notes-box" id="notesTrienKhai">
          <strong>Nhận xét:</strong>
          <ul id="listNotesTrienKhai"><li>Chưa có dữ liệu.</li></ul>
        </div>
      </div>
    </div>

    <!-- PHẦN III: KHAI THÁC -->
    <div class="section-title">III. KHAI THÁC CỔNG PORT</div>
    <div class="grid-container">
      <div class="card full-width">
        <h3>Thống Kê Khai Thác Port Theo POP</h3>
        <div style="overflow-x: auto;">
          <table id="tableKhaiThac">
            <thead>
              <tr>
                <th>STT</th>
                <th>Tên POP</th>
                <th>Tổng Port</th>
                <th>Port Sử Dụng</th>
                <th>Port Trống</th>
                <th>Tỷ Lệ (%)</th>
                <th>Đánh Giá Trạng Thái</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>

      <div class="card full-width">
        <h3>Biểu Đồ Tỷ Lệ Sử Dụng Port Theo POP (%)</h3>
        <div class="chart-container">
          <canvas id="chartPortRate"></canvas>
        </div>
      </div>
    </div>
  </div>

  <script>
    let dashboardData = {
      vanhanh: {
        headers: ["STT", "Chỉ số KPIs", "26-W46", "26-W47", "26-W48", "26-W49", "26-T12", "Plan", "+/- Plan"],
        weekHeaders: ["26-W46", "26-W47", "26-W48", "26-W49", "26-T12"],
        items: [],
        notes: []
      },
      trienkhai: {
        headers: ["STT", "POP", "26-W46", "26-W47", "26-W48", "26-W49", "26-T12"],
        weekHeaders: ["26-W46", "26-W47", "26-W48", "26-W49"],
        items: [],
        notes: []
      },
      khaithac: {
        headers: ["STT", "POP", "Tổng port", "Port sử dụng", "Port trống", "Tỷ lệ", "Ghi chú"],
        items: [],
        notes: []
      }
    };

    let chartSuCoInstance = null;
    let chartThoiGianInstance = null;
    let chartPortRateInstance = null;

    document.getElementById('excelFile').addEventListener('change', handleFileUpload);

    function handleFileUpload(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        parseMatrixData(matrix);
      };
      reader.readAsArrayBuffer(file);
    }

    function parseMatrixData(matrix) {
      const newVhItems = [];
      const newVhNotes = [];
      const newTkItems = [];
      const newTkNotes = [];
      const newKtItems = [];
      const newKtNotes = [];

      let vhHeaders = ["STT", "Chỉ Tiêu Vận Hành", "26-W46", "26-W47", "26-W48", "26-W49", "26-T12", "Plan", "+/- Plan"];
      let vhWeekHeaders = ["26-W46", "26-W47", "26-W48", "26-W49", "26-T12"];
      let tkHeaders = ["STT", "POP", "26-W46", "26-W47", "26-W48", "26-W49", "26-T12"];
      let tkWeekHeaders = ["26-W46", "26-W47", "26-W48", "26-W49"];
      let ktHeaders = ["STT", "POP", "Tổng Port", "Port Sử Dụng", "Port Trống", "Tỷ Lệ (%)", "Ghi Chú Trạng Thái"];

      let offset = 0;
      for (let r = 0; r < Math.min(matrix.length, 40); r++) {
        const row = matrix[r] || [];
        if (row[1] !== undefined && row[1] !== null && String(row[1]).trim().toUpperCase() === "STT") {
          offset = 1;
          break;
        }
      }

      let mod2StartRow = -1;
      let mod3StartRow = -1;

      for (let r = 0; r < matrix.length; r++) {
        const rowStr = (matrix[r] || []).join(" ").toLowerCase();
        if (mod2StartRow === -1 && (rowStr.includes("ii. triển khai") || rowStr.includes("ii.triển khai"))) {
          mod2StartRow = r;
        }
        if (mod3StartRow === -1 && (rowStr.includes("iii. khai thác") || rowStr.includes("iii.khai thác"))) {
          mod3StartRow = r;
        }
      }

      if (mod2StartRow === -1) mod2StartRow = 16;
      if (mod3StartRow === -1) mod3StartRow = 30;

      // I. VẬN HÀNH
      for (let r = 0; r < mod2StartRow; r++) {
        const row = matrix[r] || [];
        const firstColStr = row[0 + offset] !== undefined && row[0 + offset] !== null ? String(row[0 + offset]).trim() : "";
        const secondColStr = row[1 + offset] !== undefined && row[1 + offset] !== null ? String(row[1 + offset]).trim() : "";

        if (firstColStr.toUpperCase() === "STT" || secondColStr.toLowerCase().includes("chỉ số") || secondColStr.toLowerCase().includes("chỉ tiêu")) {
          const extracted = row.slice(offset).filter(cell => cell !== undefined && cell !== null && String(cell).trim() !== "");
          if (extracted.length >= 4) {
            vhHeaders = extracted.map(c => String(c).trim());
            vhWeekHeaders = vhHeaders.slice(2, vhHeaders.length - 2);
          }
          continue;
        }

        const sttNum = parseInt(firstColStr);
        if (!isNaN(sttNum) && sttNum > 0 && secondColStr !== "") {
          const isPercentMetric = secondColStr.includes("%") || secondColStr.toLowerCase().includes("saidi") || secondColStr.toLowerCase().includes("tỷ lệ");
          const vals = [];
          const rawVals = [];

          const numWeeks = vhWeekHeaders.length || 5;
          for (let c = 2 + offset; c < 2 + offset + numWeeks; c++) {
            let rawVal = row[c];
            let numVal = parseFloat(rawVal) || 0;
            if (typeof rawVal === 'string' && rawVal.includes('%')) {
              numVal = parseFloat(rawVal.replace('%', ''));
            } else if (isPercentMetric && numVal > 0 && numVal <= 1) {
              numVal = parseFloat((numVal * 100).toFixed(2));
            }
            rawVals.push(numVal);
            vals.push(isPercentMetric ? `${numVal}%` : numVal);
          }

          const planIdx = 2 + offset + numWeeks;
          const diffIdx = planIdx + 1;

          let rawPlanCell = row[planIdx];
          let planVal = rawPlanCell !== undefined ? rawPlanCell : 0;
          let rawPlan = parseFloat(planVal) || 0;
          if (isPercentMetric && rawPlan > 0 && rawPlan <= 1) rawPlan = parseFloat((rawPlan * 100).toFixed(2));
          planVal = isPercentMetric ? `${rawPlan}%` : rawPlan;

          let rawDiffCell = row[diffIdx];
          let diffVal = rawDiffCell !== undefined ? rawDiffCell : 0;
          let rawDiff = parseFloat(diffVal) || 0;
          if (isPercentMetric && Math.abs(rawDiff) > 0 && Math.abs(rawDiff) <= 1) rawDiff = parseFloat((rawDiff * 100).toFixed(2));

          let diffSign = rawDiff > 0 ? '+' : '';
          diffVal = isPercentMetric ? `${diffSign}${rawDiff}%` : `${diffSign}${rawDiff}`;

          newVhItems.push({
            stt: sttNum,
            metric: secondColStr,
            values: vals,
            rawValues: rawVals,
            plan: planVal,
            diff: diffVal,
            rawDiff: rawDiff,
            isPercent: isPercentMetric
          });
        } else {
          const noteText = row.slice(offset).filter(cell => cell !== "" && cell !== null && cell !== undefined).join(" ").trim();
          if (noteText.length > 2 && !noteText.toLowerCase().includes("i. vận hành") && !noteText.toLowerCase().includes("nhận xét")) {
            newVhNotes.push(noteText.replace(/^[-*•]\s*/, ""));
          }
        }
      }

      // II. TRIỂN KHAI
      for (let r = mod2StartRow; r < mod3StartRow; r++) {
        const row = matrix[r] || [];
        const firstColStr = row[0 + offset] !== undefined && row[0 + offset] !== null ? String(row[0 + offset]).trim() : "";
        const secondColStr = row[1 + offset] !== undefined && row[1 + offset] !== null ? String(row[1 + offset]).trim() : "";

        if (firstColStr.toUpperCase() === "STT" || secondColStr.toLowerCase().includes("pop")) {
          const extracted = row.slice(offset).filter(cell => cell !== undefined && cell !== null && String(cell).trim() !== "");
          if (extracted.length >= 3) {
            tkHeaders = extracted.map(c => String(c).trim());
            tkWeekHeaders = tkHeaders.slice(2, tkHeaders.length - 1);
          }
          continue;
        }

        const sttNum = parseInt(firstColStr);
        if (!isNaN(sttNum) && sttNum > 0 && secondColStr !== "") {
          const rawCells = [sttNum, secondColStr];
          const numWeeks = tkWeekHeaders.length || 4;
          for (let c = 2 + offset; c < 2 + offset + numWeeks + 1; c++) {
            rawCells.push(row[c] !== undefined && row[c] !== null ? row[c] : 0);
          }

          newTkItems.push({
            stt: sttNum,
            pop: secondColStr,
            rawCells: rawCells,
            total: parseFloat(row[2 + offset + numWeeks]) || 0
          });
        } else {
          const noteText = row.slice(offset).filter(cell => cell !== "" && cell !== null && cell !== undefined).join(" ").trim();
          if (noteText.length > 2 && !noteText.toLowerCase().includes("ii. triển khai") && !noteText.toLowerCase().includes("nhận xét")) {
            newTkNotes.push(noteText.replace(/^[-*•]\s*/, ""));
          }
        }
      }

      // III. KHAI THÁC
      for (let r = mod3StartRow; r < matrix.length; r++) {
        const row = matrix[r] || [];
        const firstColStr = row[0 + offset] !== undefined && row[0 + offset] !== null ? String(row[0 + offset]).trim() : "";
        const secondColStr = row[1 + offset] !== undefined && row[1 + offset] !== null ? String(row[1 + offset]).trim() : "";

        const sttNum = parseInt(firstColStr);
        if (!isNaN(sttNum) && sttNum > 0 && secondColStr !== "") {
          const totalP = parseFloat(row[2 + offset]) || 0;
          const usedP = parseFloat(row[3 + offset]) || 0;
          const freeP = parseFloat(row[4 + offset]) || 0;
          let rateP = parseFloat(row[5 + offset]) || 0;
          if (typeof row[5 + offset] === 'string' && row[5 + offset].includes('%')) {
            rateP = parseFloat(row[5 + offset].replace('%', ''));
          } else if (rateP > 0 && rateP <= 1) {
            rateP = rateP * 100;
          }

          newKtItems.push({
            stt: sttNum,
            pop: secondColStr,
            total: totalP,
            used: usedP,
            free: freeP,
            rate: parseFloat(rateP.toFixed(2)),
            note: row[6 + offset] || (rateP >= 55 ? "Tỷ lệ khai hiệu quả" : (rateP <= 30 ? "Tỷ lệ khai thác thấp" : "Trung bình"))
          });
        } else {
          const noteText = row.slice(offset).filter(cell => cell !== "" && cell !== null && cell !== undefined).join(" ").trim();
          if (noteText.length > 2 && !noteText.toLowerCase().includes("iii. khai thác") && !noteText.toLowerCase().includes("nhận xét")) {
            newKtNotes.push(noteText.replace(/^[-*•]\s*/, ""));
          }
        }
      }

      dashboardData.vanhanh.headers = vhHeaders;
      dashboardData.vanhanh.weekHeaders = vhWeekHeaders;
      if (newVhItems.length > 0) dashboardData.vanhanh.items = newVhItems;
      if (newVhNotes.length > 0) dashboardData.vanhanh.notes = newVhNotes;

      dashboardData.trienkhai.headers = tkHeaders;
      dashboardData.trienkhai.weekHeaders = tkWeekHeaders;
      if (newTkItems.length > 0) dashboardData.trienkhai.items = newTkItems;
      if (newTkNotes.length > 0) dashboardData.trienkhai.notes = newTkNotes;

      dashboardData.khaithac.headers = ktHeaders;
      if (newKtItems.length > 0) dashboardData.khaithac.items = newKtItems;
      if (newKtNotes.length > 0) dashboardData.khaithac.notes = newKtNotes;

      renderDashboard();
    }

    function renderDashboard() {
      renderTableVanHanh();
      renderTableTrienKhai();
      renderTableKhaiThac();
      renderCharts();
    }

    function renderTableVanHanh() {
      const table = document.getElementById('tableVanHanh');
      const thead = table.querySelector('thead');
      const tbody = table.querySelector('tbody');

      thead.innerHTML = `<tr>${dashboardData.vanhanh.headers.map(h => `<th>${h}</th>`).join('')}</tr>`;

      tbody.innerHTML = dashboardData.vanhanh.items.map(item => `
        <tr>
          <td>${item.stt}</td>
          <td><strong>${item.metric}</strong></td>
          ${item.values.map(v => `<td>${v}</td>`).join('')}
          <td>${item.plan}</td>
          <td style="color: ${item.rawDiff > 0 ? '#dc2626' : '#16a34a'}; font-weight: bold;">${item.diff}</td>
        </tr>
      `).join('');

      const listNotes = document.getElementById('listNotesVanHanh');
      listNotes.innerHTML = dashboardData.vanhanh.notes.length > 0 
        ? dashboardData.vanhanh.notes.map(n => `<li>${n}</li>`).join('')
        : '<li>Không có ghi chú thêm.</li>';
    }

    function renderTableTrienKhai() {
      const table = document.getElementById('tableTrienKhai');
      const thead = table.querySelector('thead');
      const tbody = table.querySelector('tbody');

      thead.innerHTML = `<tr>${dashboardData.trienkhai.headers.map(h => `<th>${h}</th>`).join('')}</tr>`;

      tbody.innerHTML = dashboardData.trienkhai.items.map(item => `
        <tr>
          ${item.rawCells.map(c => `<td>${c}</td>`).join('')}
        </tr>
      `).join('');

      const listNotes = document.getElementById('listNotesTrienKhai');
      listNotes.innerHTML = dashboardData.trienkhai.notes.length > 0 
        ? dashboardData.trienkhai.notes.map(n => `<li>${n}</li>`).join('')
        : '<li>Không có ghi chú thêm.</li>';
    }

    function renderTableKhaiThac() {
      const tbody = document.getElementById('tableKhaiThac').querySelector('tbody');

      tbody.innerHTML = dashboardData.khaithac.items.map(item => {
        let badgeClass = 'badge-warning';
        if (item.rate >= 55) badgeClass = 'badge-success';
        else if (item.rate <= 30) badgeClass = 'badge-danger';

        return `
          <tr>
            <td>${item.stt}</td>
            <td><strong>${item.pop}</strong></td>
            <td>${item.total}</td>
            <td>${item.used}</td>
            <td>${item.free}</td>
            <td><strong>${item.rate}%</strong></td>
            <td><span class="badge ${badgeClass}">${item.note}</span></td>
          </tr>
        `;
      }).join('');
    }

    function renderCharts() {
      const labelsVh = dashboardData.vanhanh.weekHeaders;

      const suCoItem = dashboardData.vanhanh.items.find(i => i.stt === 1) || { rawValues: [] };
      const khgItem = dashboardData.vanhanh.items.find(i => i.stt === 3) || { rawValues: [] };

      if (chartSuCoInstance) chartSuCoInstance.destroy();
      const ctxSuCo = document.getElementById('chartSuCo').getContext('2d');
      chartSuCoInstance = new Chart(ctxSuCo, {
        type: 'line',
        data: {
          labels: labelsVh,
          datasets: [
            { label: 'Số sự cố', data: suCoItem.rawValues, borderColor: '#ef4444', backgroundColor: 'rgba(239,68,68,0.1)', fill: true, tension: 0.3 },
            { label: 'Số KHG ảnh hưởng', data: khgItem.rawValues, borderColor: '#3b82f6', backgroundColor: 'rgba(59,130,246,0.1)', fill: true, tension: 0.3 }
          ]
        },
        options: { responsive: true, maintainAspectRatio: false }
      });

      const xlscItem = dashboardData.vanhanh.items.find(i => i.stt === 4) || { rawValues: [] };
      const gianDoanItem = dashboardData.vanhanh.items.find(i => i.stt === 5) || { rawValues: [] };

      if (chartThoiGianInstance) chartThoiGianInstance.destroy();
      const ctxThoiGian = document.getElementById('chartThoiGian').getContext('2d');
      chartThoiGianInstance = new Chart(ctxThoiGian, {
        type: 'bar',
        data: {
          labels: labelsVh,
          datasets: [
            { label: 'Thời gian XLSC TB (phút)', data: xlscItem.rawValues, backgroundColor: '#f59e0b' },
            { label: 'Thời gian gián đoạn TB (phút)', data: gianDoanItem.rawValues, backgroundColor: '#10b981' }
          ]
        },
        options: { responsive: true, maintainAspectRatio: false }
      });

      const popLabels = dashboardData.khaithac.items.map(i => i.pop);
      const portRates = dashboardData.khaithac.items.map(i => i.rate);

      if (chartPortRateInstance) chartPortRateInstance.destroy();
      const ctxPort = document.getElementById('chartPortRate').getContext('2d');
      chartPortRateInstance = new Chart(ctxPort, {
        type: 'bar',
        data: {
          labels: popLabels,
          datasets: [{
            label: 'Tỷ lệ khai thác (%)',
            data: portRates,
            backgroundColor: portRates.map(r => r >= 55 ? '#10b981' : (r <= 30 ? '#ef4444' : '#f59e0b'))
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: { y: { beginAtZero: true, max: 100 } }
        }
      });
    }
  </script>
</body>
</html>
"""

# Hiển thị toàn bộ Component HTML trên ứng dụng Streamlit
components.html(html_code, height=1800, scrolling=True)
