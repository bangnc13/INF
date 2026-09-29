import streamlit as st
import streamlit.components.v1 as components

# Cấu hình trang Streamlit full chiều rộng giao diện
st.set_page_config(
    page_title="Executive Dashboard - INF TQG",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Nhúng mã CSS đè lề chuẩn của Streamlit để hiển thị Fullscreen mượt mà
st.markdown("""
    <style>
        /* Ẩn bớt padding dư thừa của Streamlit */
        .block-container {
            padding-top: 0rem !important;
            padding-bottom: 0rem !important;
            padding-left: 0rem !important;
            padding-right: 0rem !important;
            max-width: 100% !important;
        }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# Mã nguồn HTML / Tailwind / Chart.js nguyên bản giữ nguyên giao diện & màu sắc
HTML_CODE = """
<!DOCTYPE html>
<html lang="vi" class="h-full bg-slate-900 text-slate-100">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Dashboard INF TQG - Executive Dashboard</title>
  
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      darkMode: 'class',
      theme: {
        extend: {
          colors: {
            vanhanh: '#3b82f6',
            trienkhai: '#10b981',
            khaithac: '#f59e0b',
          }
        }
      }
    }
  </script>

  <!-- SheetJS (XLSX Parser) -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <!-- Chart.js -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>

  <style>
    /* Custom scrollbars */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #1e293b; }
    ::-webkit-scrollbar-thumb { background: #475569; border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #64748b; }

    .glass-card {
      background: rgba(30, 41, 59, 0.75);
      backdrop-filter: blur(12px);
      border: 1px solid rgba(255, 255, 255, 0.08);
    }
  </style>
</head>
<body class="h-full flex flex-col font-sans antialiased text-slate-200 bg-slate-900">

  <!-- Top Navigation Header -->
  <header class="glass-card border-b border-slate-800 sticky top-0 z-30 px-4 lg:px-8 py-3 flex flex-wrap items-center justify-between gap-4">
    <div class="flex items-center gap-3">
      <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 via-emerald-500 to-amber-500 flex items-center justify-center text-white font-bold text-xl shadow-lg">
        D
      </div>
      <div>
        <h1 class="text-lg font-bold text-white tracking-wide">Dashboard INF TQG</h1>
        <p class="text-xs text-slate-400">Make by BangNC13</p>
      </div>
    </div>

    <!-- Source File Info Badge in Header -->
    <div class="hidden xl:flex items-center gap-4 px-4 py-1.5 rounded-xl bg-slate-950/60 border border-slate-800 text-xs">
      <div class="flex items-center gap-2">
        <span class="inline-block w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse" id="gsheetStatusDot"></span>
        <span class="text-slate-400">Nguồn:</span>
        <span id="activeFileName" class="font-medium text-emerald-400 truncate max-w-[140px]">Google Sheet Live</span>
      </div>
      <div class="w-px h-3 bg-slate-800"></div>
      <div class="flex items-center gap-1.5 text-slate-400">
        <span>POPs:</span>
        <span id="totalPopsCount" class="font-bold text-amber-400">0</span>
      </div>
      <div class="w-px h-3 bg-slate-800"></div>
      <div class="flex items-center gap-1.5 text-slate-400">
        <span>Cập nhật:</span>
        <span id="lastUpdatedTime" class="text-slate-300">Đang tải...</span>
      </div>
    </div>

    <!-- Actions & Google Sheet Sync Toolbar -->
    <div class="flex items-center gap-2.5 flex-wrap">
      <button onclick="toggleGSheetModal()" class="flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white shadow-lg shadow-emerald-600/25 transition active:scale-95 border border-emerald-400/30">
        <i data-lucide="sheet" class="w-4 h-4"></i>
        <span>Google Sheet Live</span>
        <span id="syncPulseBadge" class="w-2 h-2 rounded-full bg-white animate-ping hidden"></span>
      </button>

      <button onclick="downloadSampleExcel()" class="hidden sm:flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 transition">
        <i data-lucide="download" class="w-4 h-4 text-emerald-400"></i>
        <span>Tải File Mẫu</span>
      </button>

      <label class="cursor-pointer flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition active:scale-95">
        <i data-lucide="upload-cloud" class="w-4 h-4 text-blue-400"></i>
        <span>Import File Offline</span>
        <input type="file" id="excelFileInput" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleFileUpload(event)" />
      </label>
    </div>
  </header>

  <!-- GOOGLE SHEET LIVE CONFIG MODAL -->
  <div id="gsheetModal" class="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-md hidden flex items-center justify-center p-4">
    <div class="glass-card max-w-lg w-full rounded-2xl p-6 border border-slate-700 shadow-2xl relative space-y-5 animate-fade-in">
      <button onclick="toggleGSheetModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition">
        <i data-lucide="x" class="w-5 h-5"></i>
      </button>

      <div class="flex items-center gap-3">
        <div class="w-10 h-10 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center border border-emerald-500/30">
          <i data-lucide="refresh-cw" class="w-5 h-5" id="syncSpinIcon"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">Kết Nối Google Sheet Live</h3>
          <p class="text-xs text-slate-400">Đồng bộ dữ liệu tự động thời gian thực từ bảng tính</p>
        </div>
      </div>

      <!-- URL Input & Connect Controls -->
      <div class="space-y-2">
        <label class="text-xs font-semibold text-slate-300">Đường dẫn Google Sheet (Spreadsheet URL):</label>
        <div class="flex gap-2">
          <input type="text" id="gsheetUrlInput" value="https://docs.google.com/spreadsheets/d/1UpCxfr4MsecT5M4umP-gN7DTAj6yspJsWXOQRSuAqYo/edit?gid=0#gid=0" placeholder="https://docs.google.com/spreadsheets/d/..." class="flex-1 px-3.5 py-2.5 rounded-xl text-xs bg-slate-900 border border-slate-700 text-slate-100 focus:outline-none focus:border-emerald-500 font-mono" />
          <button onclick="syncGoogleSheetNow()" class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold transition flex items-center gap-1.5 flex-shrink-0 shadow-lg shadow-emerald-600/30">
            <i data-lucide="rotate-cw" class="w-3.5 h-3.5"></i>
            <span>Đồng Bộ Ngay</span>
          </button>
        </div>
      </div>

      <!-- Auto Sync Settings -->
      <div class="glass-card p-4 rounded-xl border border-slate-800 space-y-3">
        <div class="flex items-center justify-between">
          <div class="flex items-center gap-2">
            <i data-lucide="clock" class="w-4 h-4 text-amber-400"></i>
            <span class="text-xs font-medium text-slate-200">Tự động làm mới khi nhập liệu:</span>
          </div>
          <label class="relative inline-flex items-center cursor-pointer">
            <input type="checkbox" id="autoSyncToggle" class="sr-only peer" checked onchange="toggleAutoSync(this.checked)">
            <div class="w-9 h-5 bg-slate-700 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-emerald-500"></div>
          </label>
        </div>

        <div class="flex items-center justify-between text-xs text-slate-400 pt-1 border-t border-slate-800/80">
          <span>Tần suất cập nhật:</span>
          <select id="autoSyncIntervalSelect" onchange="changeSyncInterval(this.value)" class="bg-slate-900 text-slate-200 text-xs rounded-lg border border-slate-700 px-2 py-1 focus:outline-none">
            <option value="10000">10 giây</option>
            <option value="15000" selected>15 giây</option>
            <option value="30000">30 giây</option>
            <option value="60000">1 phút</option>
          </select>
        </div>
      </div>

      <div class="p-3.5 rounded-xl bg-blue-950/40 border border-blue-500/30 text-xs space-y-1.5 text-blue-200">
        <div class="flex items-center gap-1.5 font-bold text-blue-300">
          <i data-lucide="info" class="w-4 h-4 text-blue-400"></i>
          <span>Lưu ý quan trọng để Google Sheet đồng bộ:</span>
        </div>
        <ol class="list-decimal list-inside space-y-1 text-[11px] text-slate-300">
          <li>Mở File Google Sheet của bạn.</li>
          <li>Nút <b>Chia sẻ (Share)</b> ở góc phải trên cùng &rarr; Chọn <b>Bất kỳ ai có đường liên kết (Anyone with the link)</b>.</li>
          <li>Mọi thay đổi số liệu trên Google Sheet sẽ được ứng dụng tự động tải về mà không cần F5!</li>
        </ol>
      </div>

      <div class="flex justify-end pt-2">
        <button onclick="toggleGSheetModal()" class="px-4 py-2 rounded-xl text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 transition">Đóng</button>
      </div>
    </div>
  </div>

  <div class="flex-1 flex flex-col overflow-hidden">
    <!-- Main Content Area -->
    <main class="flex-1 overflow-y-auto p-4 lg:p-8 space-y-6">
      
      <!-- Top Horizontal Module Tabs Bar -->
      <div class="glass-card p-2 rounded-2xl border border-slate-800 shadow-xl overflow-x-auto">
        <nav class="flex items-center gap-2 min-w-max">
          <button onclick="switchTab('all')" id="tab-btn-all" class="nav-tab-btn active flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition-all duration-200 bg-blue-600 text-white shadow-lg shadow-blue-600/25 border border-blue-500/50">
            <i data-lucide="layout-dashboard" class="w-4 h-4"></i>
            <span>Tổng Quan Hệ Thống</span>
          </button>

          <button onclick="switchTab('vanhanh')" id="tab-btn-vanhanh" class="nav-tab-btn flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition-all duration-200 bg-slate-800/80 hover:bg-slate-700/80 text-slate-300 hover:text-white border border-slate-700/50">
            <i data-lucide="cpu" class="w-4 h-4 text-blue-400"></i>
            <span> I. Vận Hành</span>
          </button>

          <button onclick="switchTab('trienkhai')" id="tab-btn-trienkhai" class="nav-tab-btn flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition-all duration-200 bg-slate-800/80 hover:bg-slate-700/80 text-slate-300 hover:text-white border border-slate-700/50">
            <i data-lucide="rocket" class="w-4 h-4 text-emerald-400"></i>
            <span>II. Triển Khai</span>
          </button>

          <button onclick="switchTab('khaithac')" id="tab-btn-khaithac" class="nav-tab-btn flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition-all duration-200 bg-slate-800/80 hover:bg-slate-700/80 text-slate-300 hover:text-white border border-slate-700/50">
            <i data-lucide="trending-up" class="w-4 h-4 text-amber-400"></i>
            <span> III. Khai Thác</span>
          </button>
        </nav>
      </div>

      <!-- Quick Notification Banner -->
      <div id="dataUpdateBanner" class="hidden p-3.5 rounded-xl bg-emerald-950/40 border border-emerald-500/30 text-emerald-300 text-xs flex items-center justify-between animate-fade-in">
        <div class="flex items-center gap-2.5">
          <i data-lucide="check-circle-2" class="w-4 h-4 text-emerald-400 flex-shrink-0"></i>
          <span id="updateBannerMessage">Đã update</span>
        </div>
        <button onclick="document.getElementById('dataUpdateBanner').classList.add('hidden')" class="text-emerald-400 hover:text-emerald-200">
          <i data-lucide="x" class="w-4 h-4"></i>
        </button>
      </div>

      <!-- SECTION: TỔNG QUAN (OVERVIEW) -->
      <section id="section-all" class="tab-content space-y-6">
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div class="glass-card p-6 rounded-2xl relative overflow-hidden group border border-emerald-500/20">
            <div class="absolute right-4 top-4 w-12 h-12 rounded-xl bg-emerald-500/10 text-emerald-400 flex items-center justify-center">
              <i data-lucide="server" class="w-6 h-6"></i>
            </div>
            <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider">RA PORT TRONG THÁNG </p>
            <h3 id="kpi-total-h20-ports" class="text-3xl font-black text-emerald-400 mt-2">325</h3>
          </div>

          <div class="glass-card p-6 rounded-2xl relative overflow-hidden group border border-amber-500/20">
            <div class="absolute right-4 top-4 w-12 h-12 rounded-xl bg-amber-500/10 text-amber-400 flex items-center justify-center">
              <i data-lucide="pie-chart" class="w-6 h-6"></i>
            </div>
            <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider">III: Tỷ Lệ Khai Thác TB</p>
            <h3 id="kpi-khaithac-avg-rate" class="text-3xl font-black text-amber-400 mt-2">58.4%</h3>
            <div class="mt-3 flex items-center gap-2 text-xs">
              <span id="kpi-khaithac-high-count" class="px-2 py-0.5 rounded bg-amber-500/10 text-amber-400 font-medium border border-amber-500/20">18 POPs Tỷ lệ cao</span>
            </div>
          </div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div class="glass-card p-5 rounded-2xl">
            <div class="flex items-center justify-between mb-4">
              <h3 class="font-semibold text-white flex items-center gap-2">
                <i data-lucide="line-chart" class="w-4 h-4 text-blue-400"></i>
                Xu Hướng Biến Động Các Chỉ Tiêu Vận Hành
              </h3>
            </div>
            <div class="h-64">
              <canvas id="overviewVanHanhChart"></canvas>
            </div>
          </div>

          <div class="glass-card p-5 rounded-2xl">
            <div class="flex items-center justify-between mb-4">
              <h3 class="font-semibold text-white flex items-center gap-2">
                <i data-lucide="doughnut" class="w-4 h-4 text-amber-400"></i>
                Tổng Quan Tình Trạng Khai Thác Port
              </h3>
            </div>
            <div class="h-64">
              <canvas id="overviewKhaiThacPieChart"></canvas>
            </div>
          </div>
        </div>
      </section>

      <!-- SECTION: MODUL I. VẬN HÀNH -->
      <section id="section-vanhanh" class="tab-content hidden space-y-6">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-800 pb-4">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2.5 py-1 rounded-md text-xs font-semibold bg-blue-500/20 text-blue-400 border border-blue-500/30"> I</span>
              <h2 class="text-xl font-bold text-white">Quản Lý Vận Hành</h2>
            </div>
          </div>
        </div>

        <div class="glass-card p-5 rounded-2xl">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
            <h3 class="font-semibold text-white flex items-center gap-2">
              <i data-lucide="activity" class="w-4 h-4 text-blue-400"></i>
              So Sánh Tiến Độ Các Chỉ Tiêu Vận Hành Theo Tuần (Đủ 7 Tiêu Chí)
            </h3>
          </div>
          <div class="h-80 relative overflow-x-auto">
            <canvas id="vanhanhDetailChart"></canvas>
          </div>
        </div>

        <div class="glass-card rounded-2xl overflow-hidden">
          <div class="p-4 border-b border-slate-800 font-semibold text-sm text-slate-200 flex justify-between items-center">
            <span>Bảng Chỉ Số Vận Hành (7 Tiêu Chí)</span>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-center text-xs text-slate-300 border-collapse border border-slate-800">
              <thead id="thead-vanhanh" class="bg-slate-800/80 text-slate-400 uppercase tracking-wider"></thead>
              <tbody id="table-body-vanhanh" class="divide-y divide-slate-800"></tbody>
            </table>
          </div>
        </div>

        <div class="glass-card p-5 rounded-2xl border-l-4 border-l-blue-500">
          <h4 class="font-bold text-white text-sm flex items-center gap-2 mb-3">
            <i data-lucide="file-text" class="w-4 h-4 text-blue-400"></i>
            Nhận Xét & Đánh Giá Vận Hành
          </h4>
          <ul id="vanhanh-notes-list" class="space-y-2 text-xs text-slate-300 list-disc list-inside"></ul>
        </div>
      </section>

      <!-- SECTION: II. TRIỂN KHAI -->
      <section id="section-trienkhai" class="tab-content hidden space-y-6">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-800 pb-4">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2.5 py-1 rounded-md text-xs font-semibold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30"> II</span>
              <h2 class="text-xl font-bold text-white">Quản Lý Triển Khai</h2>
            </div>
          </div>
        </div>

        <div class="glass-card rounded-2xl overflow-hidden">
          <div class="p-4 border-b border-slate-800 font-semibold text-sm text-slate-200 flex justify-between items-center">
            <span>Bảng Dữ Liệu Triển Khai</span>
            <span id="trienkhai-table-count" class="text-xs text-emerald-400"></span>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-center text-xs text-slate-300 border-collapse border border-slate-800">
              <thead id="thead-trienkhai" class="bg-slate-800/80 text-slate-400 uppercase tracking-wider"></thead>
              <tbody id="table-body-trienkhai" class="divide-y divide-slate-800"></tbody>
            </table>
          </div>
        </div>

        <div class="glass-card p-5 rounded-2xl border-l-4 border-l-emerald-500">
          <h4 class="font-bold text-white text-sm flex items-center gap-2 mb-3">
            <i data-lucide="file-text" class="w-4 h-4 text-emerald-400"></i>
            Nhận Xét & Đánh Giá Triển Khai
          </h4>
          <ul id="trienkhai-notes-list" class="space-y-2 text-xs text-slate-300 list-disc list-inside"></ul>
        </div>
      </section>

      <!-- SECTION: MODUL III. KHAI THÁC -->
      <section id="section-khaithac" class="tab-content hidden space-y-6">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-800 pb-4">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2.5 py-1 rounded-md text-xs font-semibold bg-amber-500/20 text-amber-400 border border-amber-500/30"> III</span>
              <h2 class="text-xl font-bold text-white">Quản Lý Khai Thác</h2>
            </div>
          </div>
          <div class="flex items-center gap-2">
            <input type="text" id="search-khaithac" onkeyup="filterKhaiThacTable()" placeholder="Tìm kiếm mã POP..." class="px-3 py-1.5 rounded-lg text-xs bg-slate-800 border border-slate-700 text-slate-200 focus:outline-none focus:border-amber-500" />
          </div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div class="glass-card p-5 rounded-2xl">
            <h3 class="font-semibold text-white mb-4 flex items-center gap-2">
              <i data-lucide="trending-up" class="w-4 h-4 text-amber-400"></i>
              Top 5 POP Có Tỷ Lệ Khai Thác Cao Nhất (%)
            </h3>
            <div class="h-64">
              <canvas id="khaithacTopChart"></canvas>
            </div>
          </div>

          <div class="glass-card p-5 rounded-2xl border border-rose-500/20">
            <h3 class="font-semibold text-white mb-4 flex items-center gap-2">
              <i data-lucide="trending-down" class="w-4 h-4 text-rose-400"></i>
              Top 5 POP Tỷ Lệ Khai Thác Thấp Nhất (%)
            </h3>
            <div class="h-64">
              <canvas id="khaithacLowestChart"></canvas>
            </div>
          </div>

          <div class="glass-card p-5 rounded-2xl flex flex-col justify-between">
            <div>
              <h3 class="font-semibold text-white mb-2 flex items-center gap-2">
                <i data-lucide="pie-chart" class="w-4 h-4 text-blue-400"></i>
                Tỉ lệ khai thác
              </h3>
              <p class="text-xs text-slate-400 mb-4">Chi nhánh Tuyên Quang</p>
            </div>
            <div class="h-48">
              <canvas id="khaithacPortStructureChart"></canvas>
            </div>
          </div>
        </div>

        <div class="glass-card rounded-2xl overflow-hidden">
          <div class="p-4 border-b border-slate-800 font-semibold text-sm text-slate-200 flex justify-between items-center">
            <span>Danh Sách POP Khai Thác</span>
            <span id="khaithac-table-count" class="text-xs text-amber-400"></span>
          </div>
          <div class="overflow-x-auto max-h-[480px]">
            <table class="w-full text-center text-xs text-slate-300 border-collapse border border-slate-800">
              <thead id="thead-khaithac" class="bg-slate-800/90 text-slate-400 uppercase tracking-wider sticky top-0 backdrop-blur-md"></thead>
              <tbody id="table-body-khaithac" class="divide-y divide-slate-800"></tbody>
            </table>
          </div>
        </div>

        <div class="glass-card p-5 rounded-2xl border-l-4 border-l-amber-500">
          <h4 class="font-bold text-white text-sm flex items-center gap-2 mb-3">
            <i data-lucide="file-text" class="w-4 h-4 text-amber-400"></i>
            Nhận Xét & Đánh Giá Khai Thác
          </h4>
          <ul id="khaithac-notes-list" class="space-y-2 text-xs text-slate-300 list-disc list-inside"></ul>
        </div>
      </section>

    </main>
  </div>

  <script>
    const barValuePlugin = {
      id: 'barValuePlugin',
      afterDatasetsDraw(chart) {
        if (chart.config.id !== 'vanhanhDetailChart') return;
        const { ctx } = chart;
        chart.data.datasets.forEach((dataset, i) => {
          const meta = chart.getDatasetMeta(i);
          if (!meta.hidden) {
            meta.data.forEach((element, index) => {
              const rawVal = dataset.data[index];
              if (rawVal !== null && rawVal !== undefined && rawVal !== 0) {
                ctx.save();
                ctx.fillStyle = '#cbd5e1';
                ctx.font = '9px Inter, sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                const isPercent = dashboardData.vanhanh.items[index]?.isPercent;
                const text = isPercent ? rawVal + '%' : rawVal;
                ctx.fillText(text, element.x, element.y - 2);
                ctx.restore();
              }
            });
          }
        });
      }
    };

    Chart.register(barValuePlugin);

    let chartInstances = {};

    let gsheetConfig = {
      url: "https://docs.google.com/spreadsheets/d/1UpCxfr4MsecT5M4umP-gN7DTAj6yspJsWXOQRSuAqYo/edit?gid=0#gid=0",
      autoSync: true,
      intervalMs: 15000,
      timerId: null,
      lastHash: "",
      isSyncing: false
    };

    let dashboardData = {
      totalPortH20: 325,
      vanhanh: {
        headers: ["STT", "Chỉ Tiêu Vận Hành", "26-W37", "26-W38", "26-W39", "26-W40", "Plan", "+ / - Thực Hiện"],
        weekHeaders: ["26-W37", "26-W38", "26-W39", "26-W40"],
        items: [],
        notes: []
      },
      trienkhai: {
        headers: ["STT", "Hạng Mục / POP", "26-W37", "26-W38", "26-W39", "26-W40", "Lũy Kế", "Đơn Vị", "Ghi Chú"],
        weekHeaders: ["26-W37", "26-W38", "26-W39", "26-W40"],
        items: [],
        notes: []
      },
      khaithac: {
        headers: ["STT", "POP", "Tổng Port", "Port Sử Dụng", "Port Trống", "Tỷ Lệ (%)", "Ghi Chú Trạng Thái"],
        items: [],
        notes: []
      }
    };

    const initialMockData = {
      totalPortH20: 325,
      vanhanh: {
        headers: ["STT", "Chỉ Tiêu Vận Hành", "26-W37", "26-W38", "26-W39", "26-W40", "Plan", "+ / - Thực Hiện"],
        weekHeaders: ["26-W37", "26-W38", "26-W39", "26-W40"],
        items: [
          { stt: 1, metric: "Số lượng sự cố (Sđ)", values: [24, 26, 27, 24], rawValues: [24, 26, 27, 24], plan: 31, diff: "-7", rawDiff: -7, unit: "Sđ" },
          { stt: 2, metric: "KHG Ảnh hưởng/sự cố (Slg)", values: [21.00, 16.00, 17.20, 14.20], rawValues: [21.00, 16.00, 17.20, 14.20], plan: 16.00, diff: "-1.8", rawDiff: -1.8, unit: "Slg" },
          { stt: 3, metric: "Số lượng KHG ảnh hưởng (Slg)", values: [280, 255, 490, 325], rawValues: [280, 255, 490, 325], plan: 300, diff: "+25", rawDiff: 25, unit: "Slg" },
          { stt: 4, metric: "Thời gian XL SC trung bình (phút)", values: [88.5, 180.2, 128.0, 168.5], rawValues: [88.5, 180.2, 128.0, 168.5], plan: 110, diff: "+58.5", rawDiff: 58.5, unit: "phút" },
          { stt: 5, metric: "Thời gian gián đoạn TB (phút)", values: [98.0, 182.0, 164.0, 212.0], rawValues: [98.0, 182.0, 164.0, 212.0], plan: 217, diff: "-5.0", rawDiff: -5.0, unit: "phút" },
          { stt: 6, metric: "SAIDI (phút) / SLA (%)", values: ["0.79%", "1.32%", "2.92%", "2.01%"], rawValues: [0.79, 1.32, 2.92, 2.01], plan: "5.00%", diff: "-2.99%", rawDiff: -2.99, unit: "%", isPercent: true },
          { stt: 7, metric: "Tỷ lệ KHG ảnh hưởng (%)", values: ["8.33%", "11.54%", "14.81%", "12.50%"], rawValues: [8.33, 11.54, 14.81, 12.50], plan: "5.00%", diff: "+7.50%", rawDiff: 7.50, unit: "%", isPercent: true }
        ],
        notes: [
          "Trong tuần 26-W40: Đã kiểm soát đầy đủ 7 chỉ tiêu vận hành hệ thống.",
          "Chỉ tiêu SLA đạt mức 2.01%, hoàn thành tốt hơn so với kế hoạch (5.0%).",
          "Số lượng KHG ảnh hưởng giảm mạnh so với tuần cao điểm 26-W39.",
          "Cần tiếp tục duy trì tiến độ xử lý sự cố nhanh và tối ưu thời gian gián đoạn."
        ]
      },
      trienkhai: {
        headers: ["STT", "Hạng Mục / POP", "26-W37", "26-W38", "26-W39", "26-W40", "Lũy Kế 4 Tuần", "Đơn Vị", "Ghi Chú"],
        weekHeaders: ["26-W37", "26-W38", "26-W39", "26-W40"],
        items: [
          { stt: 1, pop: "TCGF012", values: [0, 48, 0, 0], w37: 0, w38: 48, w39: 0, w40: 0, total: 48, unit: "Port", note: "Đã hoàn thành" },
          { stt: 2, pop: "Nâng cấp Port, M1", values: [72, 88, 65, 100], w37: 72, w38: 88, w39: 65, w40: 100, total: 325, unit: "Port", note: "Đang duy trì" },
          { stt: 3, pop: "Bàn giao Hạ Tầng Mới", values: [10, 15, 20, 25], w37: 10, w38: 15, w39: 20, w40: 25, total: 70, unit: "POP", note: "Đang mở rộng" }
        ],
        notes: [
          "Tiến độ triển khai trong tuần 26-W40 tăng trưởng tích cực, đặc biệt ở hạng mục Nâng cấp Port M1 đạt mốc kỷ lục 100 Port/tuần.",
          "Lũy kế 4 tuần (W37-W40) đã hoàn thành mở rộng tổng cộng 325 Port M1 và đưa vào vận hành 70 trạm POP hạ tầng mới.",
          "POP trọng điểm TCGF012 đã hoàn thành thi công và bàn giao đúng tiến độ cam kết từ tuần W38 (48 Port).",
          "Kế hoạch tiếp theo: Tập trung phối hợp nghiệm thu các trạm hạ tầng mở rộng, ưu tiên cấp dung lượng cho khu vực nhu cầu phát triển cao."
        ]
      },
      khaithac: {
        headers: ["STT", "POP", "Tổng Port", "Port Sử Dụng", "Port Trống", "Tỷ Lệ (%)", "Ghi Chú Trạng Thái"],
        items: [
          { stt: 1, pop: "TCGP001", total: 1570, used: 912, free: 658, rate: 58.09, note: "Tỷ lệ khai thác cao" },
          { stt: 2, pop: "TCGP002", total: 1432, used: 824, free: 608, rate: 57.54, note: "Tỷ lệ khai thác cao" },
          { stt: 3, pop: "TCGP003", total: 1750, used: 1010, free: 740, rate: 57.71, note: "Tỷ lệ khai thác cao" },
          { stt: 4, pop: "TCGP004", total: 2672, used: 1654, free: 1018, rate: 61.90, note: "Tỷ lệ khai thác cao" },
          { stt: 5, pop: "TCGP005", total: 2654, used: 1654, free: 1000, rate: 62.32, note: "Tỷ lệ khai thác cao" },
          { stt: 6, pop: "TCGP006", total: 2524, used: 1654, free: 870, rate: 65.53, note: "Tỷ lệ khai thác cao" },
          { stt: 7, pop: "TCGP007", total: 1434, used: 924, free: 510, rate: 64.44, note: "Tỷ lệ khai thác cao" },
          { stt: 8, pop: "TCGP008", total: 2912, used: 2185, free: 727, rate: 75.03, note: "Tỷ lệ khai thác cao" },
          { stt: 9, pop: "TCGP009", total: 2028, used: 973, free: 1055, rate: 47.98, note: "Trung bình" },
          { stt: 10, pop: "TCGP010", total: 1496, used: 1077, free: 419, rate: 71.99, note: "Tỷ lệ khai thác cao" }
        ],
        notes: [
          "Tỷ lệ khai thác toàn hệ thống trung bình đạt 58.4%, duy trì hiệu quả sử dụng tài nguyên ổn định ở mức an toàn.",
          "Nhiều POP đạt hiệu suất khai thác rất cao trên 65% như: TCGP008 (75.03%), TCGP010 (71.99%), TCGP006 (65.53%).",
          "Đã xác định các POP có tỷ lệ khai thác còn thấp (như TCGP009 đạt 47.98%) để ưu tiên quy hoạch phát triển thuê bao mới.",
          "Khuyến nghị: Đẩy mạnh kinh doanh tại các vùng phủ của trạm có tỷ lệ trống cao, đồng thời chủ động lập kế hoạch nâng cấp cho các trạm chạm ngưỡng 75%."
        ]
      }
    };

    window.onload = function() {
      lucide.createIcons();
      dashboardData = JSON.parse(JSON.stringify(initialMockData));
      renderDashboard();

      syncGoogleSheetNow();

      if (gsheetConfig.autoSync) {
        startAutoSync();
      }
    };

    function toggleGSheetModal() {
      const modal = document.getElementById('gsheetModal');
      if (modal) modal.classList.toggle('hidden');
    }

    function extractGoogleSheetDetails(urlOrId) {
      if (!urlOrId) return null;
      let str = urlOrId.trim();
      const idMatch = str.match(/\/d\/([a-zA-Z0-9-_]+)/);
      const gidMatch = str.match(/[?&]gid=([0-9]+)/);

      const sheetId = idMatch ? idMatch[1] : (str.length > 20 && !str.includes('/') ? str : null);
      const gid = gidMatch ? gidMatch[1] : '0';

      if (!sheetId) return null;
      return { sheetId, gid };
    }

    async function fetchAndParseGoogleSheet(url) {
      const details = extractGoogleSheetDetails(url);
      if (!details) {
        throw new Error("Đường dẫn Google Sheet không đúng định dạng!");
      }

      const csvUrl = `https://docs.google.com/spreadsheets/d/${details.sheetId}/gviz/tq?tqx=out:csv&gid=${details.gid}&t=${Date.now()}`;
      
      const response = await fetch(csvUrl);
      if (!response.ok) {
        throw new Error("Không thể truy cập Google Sheet!");
      }

      const csvText = await response.text();
      if (!csvText || csvText.trim().length === 0) {
        throw new Error("Bảng tính Google Sheet rỗng!");
      }

      const workbook = XLSX.read(csvText, { type: 'string' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
      
      return { matrix, rawCsv: csvText };
    }

    async function syncGoogleSheetNow(silent = false) {
      if (gsheetConfig.isSyncing) return;
      gsheetConfig.isSyncing = true;

      const urlInput = document.getElementById('gsheetUrlInput');
      const inputUrl = urlInput ? urlInput.value.trim() : gsheetConfig.url;
      if (inputUrl) gsheetConfig.url = inputUrl;

      const spinIcon = document.getElementById('syncSpinIcon');
      const statusDot = document.getElementById('gsheetStatusDot');
      const pulseBadge = document.getElementById('syncPulseBadge');

      if (spinIcon) spinIcon.classList.add('animate-spin');
      if (pulseBadge) pulseBadge.classList.remove('hidden');

      try {
        const { matrix, rawCsv } = await fetchAndParseGoogleSheet(gsheetConfig.url);

        const currentHash = rawCsv.length + "_" + rawCsv.slice(0, 100) + rawCsv.slice(-100);
        
        if (currentHash !== gsheetConfig.lastHash) {
          gsheetConfig.lastHash = currentHash;
          parseMatrixData(matrix);

          const banner = document.getElementById('dataUpdateBanner');
          const bannerMsg = document.getElementById('updateBannerMessage');
          if (banner && bannerMsg) {
            bannerMsg.innerText = `Đã tự động cập nhật dữ liệu từ Google Sheet!`;
            banner.classList.remove('hidden');
          }
        }

        const fileNameEl = document.getElementById('activeFileName');
        if (fileNameEl) fileNameEl.innerText = "Google Sheet Live";

        const timeEl = document.getElementById('lastUpdatedTime');
        if (timeEl) {
          const now = new Date();
          timeEl.innerText = now.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        }

        if (statusDot) {
          statusDot.className = "inline-block w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse";
        }
      } catch (err) {
        console.warn("Lỗi đồng bộ Google Sheet:", err);
        if (statusDot) {
          statusDot.className = "inline-block w-2.5 h-2.5 rounded-full bg-rose-500";
        }
      } finally {
        gsheetConfig.isSyncing = false;
        if (spinIcon) spinIcon.classList.remove('animate-spin');
        if (pulseBadge) pulseBadge.classList.add('hidden');
      }
    }

    function startAutoSync() {
      stopAutoSync();
      gsheetConfig.timerId = setInterval(() => {
        syncGoogleSheetNow(true);
      }, gsheetConfig.intervalMs);
    }

    function stopAutoSync() {
      if (gsheetConfig.timerId) {
        clearInterval(gsheetConfig.timerId);
        gsheetConfig.timerId = null;
      }
    }

    function toggleAutoSync(enabled) {
      gsheetConfig.autoSync = enabled;
      if (enabled) {
        startAutoSync();
        syncGoogleSheetNow(true);
      } else {
        stopAutoSync();
      }
    }

    function changeSyncInterval(ms) {
      gsheetConfig.intervalMs = parseInt(ms) || 15000;
      if (gsheetConfig.autoSync) {
        startAutoSync();
      }
    }

    function renderTable(tbodyId, items, rowRenderer) {
      const tbody = document.getElementById(tbodyId);
      if (!tbody) return;
      if (!items || items.length === 0) {
        tbody.innerHTML = `<tr><td colspan="10" class="p-4 text-center text-slate-500 italic">Không có dữ liệu</td></tr>`;
        return;
      }
      tbody.innerHTML = items.map(item => `<tr class="hover:bg-slate-800/50 transition-colors">${rowRenderer(item)}</tr>`).join('');
    }

    function renderRateBadge(rate, note) {
      let badgeClass = "bg-slate-800 text-slate-300 border-slate-700";
      if (rate >= 60) {
        badgeClass = "bg-emerald-500/10 text-emerald-400 border-emerald-500/30";
      } else if (rate >= 35) {
        badgeClass = "bg-amber-500/10 text-amber-400 border-amber-500/30";
      } else {
        badgeClass = "bg-rose-500/10 text-rose-400 border-rose-500/30";
      }
      return `<span class="px-2 py-0.5 rounded text-[11px] font-medium border ${badgeClass}">${note || (rate + '%')}</span>`;
    }

    function formatDiffTwoDecimals(diffVal, isPercent = false) {
      if (diffVal === undefined || diffVal === null || diffVal === "") return "";
      let str = String(diffVal).trim();
      if (str.includes('%')) return str;

      let cleanStr = str.replace(/\+/g, '').replace(',', '.');
      let num = parseFloat(cleanStr);
      if (isNaN(num)) return str;
      
      let rounded = Math.round(num * 100) / 100;
      let formattedStr = rounded.toString().replace('.', ',');
      let sign = rounded > 0 ? '+' : '';
      return isPercent ? `${sign}${formattedStr}%` : `${sign}${formattedStr}`;
    }

    function switchTab(tabId) {
      const contents = document.querySelectorAll('.tab-content');
      contents.forEach(el => el.classList.add('hidden'));

      if (tabId === 'all') {
        document.getElementById('section-all').classList.remove('hidden');
      } else {
        const target = document.getElementById(`section-${tabId}`);
        if (target) target.classList.remove('hidden');
      }

      const buttons = document.querySelectorAll('.nav-tab-btn');
      buttons.forEach(btn => {
        btn.classList.remove('bg-blue-600', 'text-white', 'shadow-lg', 'shadow-blue-600/25', 'border-blue-500/50');
        btn.classList.add('bg-slate-800/80', 'text-slate-300', 'hover:bg-slate-700/80', 'hover:text-white', 'border-slate-700/50');
      });

      const activeBtn = document.getElementById(`tab-btn-${tabId}`);
      if (activeBtn) {
        activeBtn.classList.remove('bg-slate-800/80', 'text-slate-300', 'hover:bg-slate-700/80', 'hover:text-white');
        activeBtn.classList.add('bg-blue-600', 'text-white', 'shadow-lg', 'shadow-blue-600/25', 'border-blue-500/50');
      }
    }

    function filterKhaiThacTable() {
      const query = (document.getElementById('search-khaithac')?.value || '').toLowerCase().trim();
      const filtered = dashboardData.khaithac.items.filter(item => 
        item.pop.toLowerCase().includes(query) || (item.note && item.note.toLowerCase().includes(query))
      );

      renderTable('table-body-khaithac', filtered, (item) => `
        <td class="p-3 text-center border border-slate-800 font-semibold text-slate-400">${item.stt}</td>
        <td class="p-3 text-center border border-slate-800 font-bold text-white">${item.pop}</td>
        <td class="p-3 text-center border border-slate-800 font-mono text-slate-300">${item.total.toLocaleString('vi-VN')}</td>
        <td class="p-3 text-center border border-slate-800 font-mono text-emerald-400 font-semibold">${item.used.toLocaleString('vi-VN')}</td>
        <td class="p-3 text-center border border-slate-800 font-mono text-slate-400">${item.free.toLocaleString('vi-VN')}</td>
        <td class="p-3 text-center border border-slate-800 font-mono font-bold ${item.rate >= 50 ? 'text-amber-400' : (item.rate < 30 ? 'text-rose-400' : 'text-slate-300')}">${item.rate}%</td>
        <td class="p-3 text-center border border-slate-800">${renderRateBadge(item.rate, item.note)}</td>
      `);

      const countEl = document.getElementById('khaithac-table-count');
      if (countEl) countEl.innerText = `Hiển thị ${filtered.length} POPs`;
    }

    function handleFileUpload(event) {
      const file = event.target.files[0];
      if (!file) return;

      const fileNameEl = document.getElementById('activeFileName');
      if (fileNameEl) fileNameEl.innerText = file.name;

      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

          parseMatrixData(matrix);

          const banner = document.getElementById('dataUpdateBanner');
          const bannerMsg = document.getElementById('updateBannerMessage');
          if (banner && bannerMsg) {
            bannerMsg.innerText = `Đã import "${file.name}" thành công`;
            banner.classList.remove('hidden');
          }

          const timeEl = document.getElementById('lastUpdatedTime');
          if (timeEl) {
            const now = new Date();
            timeEl.innerText = now.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' });
          }
        } catch (err) {
          console.error("Lỗi khi đọc file Excel:", err);
        }
      };

      reader.readAsArrayBuffer(file);
      event.target.value = '';
    }

    function parseMatrixData(matrix) {
      const newVhItems = [];
      const newVhNotes = [];
      const newTkItems = [];
      const newTkNotes = [];
      const newKtItems = [];
      const newKtNotes = [];

      let vhHeaders = ["STT", "Chỉ Tiêu Vận Hành", "26-W37", "26-W38", "26-W39", "26-W40", "Plan", "+ / - Thực Hiện"];
      let vhWeekHeaders = ["26-W37", "26-W38", "26-W39", "26-W40"];

      let tkHeaders = ["STT", "Hạng Mục / POP", "26-W37", "26-W38", "26-W39", "26-W40", "Lũy Kế", "Đơn Vị", "Ghi Chú"];
      let tkWeekHeaders = ["26-W37", "26-W38", "26-W39", "26-W40"];

      let ktHeaders = ["STT", "POP", "Tổng Port", "Port Sử Dụng", "Port Trống", "Tỷ Lệ (%)", "Ghi Chú Trạng Thái"];

      if (matrix[19]) {
        let valH20 = matrix[19][7] !== undefined && matrix[19][7] !== "" ? matrix[19][7] : matrix[19][6];
        if (valH20 !== undefined && valH20 !== null && valH20 !== "") {
          let parsedVal = parseFloat(String(valH20).replace(/,/g, ''));
          dashboardData.totalPortH20 = !isNaN(parsedVal) ? parsedVal : valH20;
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

      // MODUL I: VẬN HÀNH
      for (let r = 0; r < mod2StartRow; r++) {
        const row = matrix[r] || [];
        const firstColStr = (row[0] !== undefined && row[0] !== null) ? row[0].toString().trim() : "";
        const secondColStr = (row[1] !== undefined && row[1] !== null) ? row[1].toString().trim() : "";

        if (firstColStr.toUpperCase() === "STT" || secondColStr.toLowerCase().includes("chỉ tiêu")) {
          const extracted = row.filter(cell => cell !== undefined && cell !== null && cell.toString().trim() !== "");
          if (extracted.length >= 4) {
            vhHeaders = extracted.map(c => String(c).trim());
            vhWeekHeaders = vhHeaders.slice(2, Math.max(3, vhHeaders.length - 2));
          }
          continue;
        }

        const sttNum = parseInt(firstColStr);
        if (!isNaN(sttNum) && sttNum > 0 && secondColStr !== "") {
          const isPercentMetric = secondColStr.includes("%") || secondColStr.toLowerCase().includes("sla") || secondColStr.toLowerCase().includes("tỷ lệ");
          const vals = [];
          const rawVals = [];

          for (let c = 2; c < Math.max(3, vhHeaders.length - 2); c++) {
            let rawVal = row[c];
            let numVal = parseFloat(rawVal) || 0;
            if (typeof rawVal === 'string' && rawVal.includes('%')) {
              numVal = parseFloat(rawVal.replace('%', ''));
            } else if (isPercentMetric && numVal > 0 && numVal <= 1) {
              numVal = parseFloat((numVal * 100).toFixed(2));
            }
            rawVals.push(numVal);

            if (isPercentMetric) {
              vals.push(`${numVal}%`);
            } else {
              vals.push(numVal);
            }
          }

          const planIdx = vhHeaders.length - 2 >= 2 ? vhHeaders.length - 2 : 6;
          const diffIdx = vhHeaders.length - 1 >= 2 ? vhHeaders.length - 1 : 7;

          let rawPlanCell = row[planIdx];
          let planVal = rawPlanCell;
          let rawPlan = 0;

          if (typeof rawPlanCell === 'string') {
            rawPlan = parseFloat(rawPlanCell.replace(/%/g, '').replace(/\+/g, '').replace(',', '.')) || 0;
            planVal = rawPlanCell.trim();
            if (isPercentMetric && !planVal.includes('%')) planVal = `${planVal}%`;
          } else {
            rawPlan = parseFloat(rawPlanCell) || 0;
            if (isPercentMetric && rawPlan > 0 && rawPlan <= 1) {
              rawPlan = parseFloat((rawPlan * 100).toFixed(2));
            }
            planVal = isPercentMetric ? `${rawPlan.toString().replace('.', ',')}%` : rawPlan;
          }

          let rawDiffCell = row[diffIdx];
          let diffVal = rawDiffCell;
          let rawDiff = 0;

          if (typeof rawDiffCell === 'string') {
            rawDiff = parseFloat(rawDiffCell.replace(/%/g, '').replace(/\+/g, '').replace(',', '.')) || 0;
            if (rawDiffCell.trim().startsWith('-')) rawDiff = -Math.abs(rawDiff);
            diffVal = rawDiffCell.trim();
            if (isPercentMetric && !diffVal.includes('%')) diffVal = `${diffVal}%`;
          } else {
            rawDiff = parseFloat(rawDiffCell) || 0;
            if (isPercentMetric && Math.abs(rawDiff) > 0 && Math.abs(rawDiff) <= 1) {
              rawDiff = parseFloat((rawDiff * 100).toFixed(2));
            }
            let diffSign = rawDiff > 0 ? '+' : '';
            let formattedStr = Math.abs(rawDiff).toString().replace('.', ',');
            diffVal = isPercentMetric ? `${diffSign}${formattedStr}%` : `${diffSign}${formattedStr}`;
          }

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
          const noteText = row.filter(cell => cell !== "" && cell !== null).join(" ").trim();
          if (noteText.length > 5 && !noteText.toLowerCase().includes("i. vận hành") && !noteText.toLowerCase().includes("ii. triển khai")) {
            newVhNotes.push(noteText.replace(/^[-*•]\s*/, ""));
          }
        }
      }

      // MODUL II & III PARSING
      for (let r = mod2StartRow; r < mod3StartRow; r++) {
        const row = matrix[r] || [];
        const firstColStr = (row[0] !== undefined && row[0] !== null) ? row[0].toString().trim() : "";
        const secondColStr = (row[1] !== undefined && row[1] !== null) ? row[1].toString().trim() : "";

        if (firstColStr.toUpperCase() === "STT" || secondColStr.toLowerCase().includes("pop") || secondColStr.toLowerCase().includes("hạng mục")) {
          const extracted = row.filter(cell => cell !== undefined && cell !== null && cell.toString().trim() !== "");
          if (extracted.length >= 3) {
            tkHeaders = extracted.map(c => String(c).trim());
            tkWeekHeaders = tkHeaders.filter((h, idx) => idx >= 2 && !h.toLowerCase().includes("lũy kế") && !h.toLowerCase().includes("đơn vị") && !h.toLowerCase().includes("ghi chú") && !h.toLowerCase().includes("tổng") && !h.toLowerCase().includes("note"));
          }
          continue;
        }

        const sttNum = parseInt(firstColStr);
        if (!isNaN(sttNum) && sttNum > 0 && secondColStr !== "") {
          const rawCells = [];
          for (let c = 0; c < tkHeaders.length; c++) {
            rawCells.push(row[c] !== undefined && row[c] !== null ? row[c] : "");
          }

          const vals = [];
          const numWeeks = tkWeekHeaders.length || 4;
          for (let c = 2; c < 2 + numWeeks; c++) {
            vals.push(parseFloat(row[c]) || 0);
          }

          newTkItems.push({
            stt: sttNum,
            pop: secondColStr,
            values: vals,
            rawCells: rawCells,
            total: parseFloat(row[2 + numWeeks]) || parseFloat(row[6]) || 0,
            unit: row[3 + numWeeks] || row[7] || "Port",
            note: row[4 + numWeeks] || row[8] || "Đang thực hiện"
          });
        } else {
          const noteText = row.filter(cell => cell !== "" && cell !== null).join(" ").trim();
          if (noteText.length > 5 && !noteText.toLowerCase().includes("ii. triển khai") && !noteText.toLowerCase().includes("iii. khai thác")) {
            newTkNotes.push(noteText.replace(/^[-*•]\s*/, ""));
          }
        }
      }

      for (let r = mod3StartRow; r < matrix.length; r++) {
        const row = matrix[r] || [];
        const firstColStr = (row[0] !== undefined && row[0] !== null) ? row[0].toString().trim() : "";
        const secondColStr = (row[1] !== undefined && row[1] !== null) ? row[1].toString().trim() : "";

        const sttNum = parseInt(firstColStr);
        if (!isNaN(sttNum) && sttNum > 0 && secondColStr !== "") {
          const totalP = parseFloat(row[2]) || 0;
          const usedP = parseFloat(row[3]) || 0;
          const freeP = parseFloat(row[4]) || 0;
          let rateP = parseFloat(row[5]) || 0;
          if (typeof row[5] === 'string' && row[5].includes('%')) {
            rateP = parseFloat(row[5].replace('%', ''));
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
            note: row[6] || (rateP >= 50 ? "Tỷ lệ khai thác cao" : (rateP < 30 ? "Thấp - Cần tối ưu" : "Trung bình"))
          });
        } else {
          const noteText = row.filter(cell => cell !== "" && cell !== null).join(" ").trim();
          if (noteText.length > 5 && !noteText.toLowerCase().includes("iii. khai thác")) {
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

    function renderTableHeader(theadId, headers) {
      const thead = document.getElementById(theadId);
      if (!thead) return;

      const ths = headers.map((h) => {
        return `<th class="p-3 text-center border border-slate-700/60 bg-slate-800/90 font-semibold text-slate-300">${h}</th>`;
      }).join('');

      thead.innerHTML = `<tr>${ths}</tr>`;
    }

    function renderDashboard() {
      const totalH20El = document.getElementById('kpi-total-h20-ports');
      if (totalH20El) {
        let val = dashboardData.totalPortH20;
        totalH20El.innerText = typeof val === 'number' ? val.toLocaleString('vi-VN') : val;
      }

      const vhHeaders = dashboardData.vanhanh.headers;
      const vhItems = dashboardData.vanhanh.items;
      const vhNotes = dashboardData.vanhanh.notes;

      renderTableHeader('thead-vanhanh', vhHeaders);

      renderTable('table-body-vanhanh', vhItems, (item) => {
        const valCells = (item.values || []).map((val) => {
          return `<td class="p-3 text-center border border-slate-800 font-mono text-slate-200">${val}</td>`;
        }).join('');

        const formattedDiff = formatDiffTwoDecimals(item.diff, item.isPercent);

        return `
          <td class="p-3 text-center border border-slate-800 font-semibold text-slate-400">${item.stt}</td>
          <td class="p-3 text-center border border-slate-800 font-medium text-white">${item.metric}</td>
          ${valCells}
          <td class="p-3 text-center border border-slate-800 font-mono text-slate-400">${item.plan}</td>
          <td class="p-3 text-center border border-slate-800 font-mono ${item.rawDiff <= 0 || (typeof formattedDiff === 'string' && formattedDiff.startsWith('-')) ? 'text-emerald-400' : 'text-amber-400'}">${formattedDiff}</td>
        `;
      });

      const vhNotesContainer = document.getElementById('vanhanh-notes-list');
      if (vhNotesContainer) {
        if (vhNotes.length > 0) {
          vhNotesContainer.innerHTML = vhNotes.map(n => `<li class="leading-relaxed">${n}</li>`).join('');
        } else {
          vhNotesContainer.innerHTML = `<li class="text-slate-500 italic">Không có ghi chú nhận xét nào.</li>`;
        }
      }

      const tkHeaders = dashboardData.trienkhai.headers;
      const tkItems = dashboardData.trienkhai.items;
      const tkNotes = dashboardData.trienkhai.notes;
      renderTableHeader('thead-trienkhai', tkHeaders);
      renderTable('table-body-trienkhai', tkItems, (item) => {
        if (item.rawCells && item.rawCells.length > 0) {
          return item.rawCells.map((cellVal, colIdx) => {
            let colHeader = (tkHeaders[colIdx] || '').toLowerCase();
            let colorClass = 'text-slate-200';
            if (colIdx === 0) colorClass = 'text-slate-400 font-semibold';
            else if (colHeader.includes('lũy kế') || colHeader.includes('tổng')) colorClass = 'text-emerald-400 font-bold font-mono';
            else if (colHeader.includes('ghi chú')) return `<td class="p-3 text-center border border-slate-800"><span class="px-2 py-0.5 rounded text-[11px] bg-emerald-500/10 text-emerald-400 border border-emerald-500/30">${cellVal}</span></td>`;

            return `<td class="p-3 text-center border border-slate-800 ${colorClass}">${cellVal}</td>`;
          }).join('');
        }
        return '';
      });

      const tkNotesContainer = document.getElementById('trienkhai-notes-list');
      if (tkNotesContainer) {
        if (tkNotes && tkNotes.length > 0) {
          tkNotesContainer.innerHTML = tkNotes.map(n => `<li class="leading-relaxed">${n}</li>`).join('');
        } else {
          tkNotesContainer.innerHTML = `<li class="text-slate-500 italic">Không có ghi chú nhận xét nào.</li>`;
        }
      }

      const ktHeaders = dashboardData.khaithac.headers;
      const ktItems = dashboardData.khaithac.items;
      const ktNotes = dashboardData.khaithac.notes;
      renderTableHeader('thead-khaithac', ktHeaders);
      renderTable('table-body-khaithac', ktItems, (item) => `
        <td class="p-3 text-center border border-slate-800 font-semibold text-slate-400">${item.stt}</td>
        <td class="p-3 text-center border border-slate-800 font-bold text-white">${item.pop}</td>
        <td class="p-3 text-center border border-slate-800 font-mono text-slate-300">${item.total.toLocaleString('vi-VN')}</td>
        <td class="p-3 text-center border border-slate-800 font-mono text-emerald-400 font-semibold">${item.used.toLocaleString('vi-VN')}</td>
        <td class="p-3 text-center border border-slate-800 font-mono text-slate-400">${item.free.toLocaleString('vi-VN')}</td>
        <td class="p-3 text-center border border-slate-800 font-mono font-bold ${item.rate >= 50 ? 'text-amber-400' : (item.rate < 30 ? 'text-rose-400' : 'text-slate-300')}">${item.rate}%</td>
        <td class="p-3 text-center border border-slate-800">${renderRateBadge(item.rate, item.note)}</td>
      `);

      const ktNotesContainer = document.getElementById('khaithac-notes-list');
      if (ktNotesContainer) {
        if (ktNotes && ktNotes.length > 0) {
          ktNotesContainer.innerHTML = ktNotes.map(n => `<li class="leading-relaxed">${n}</li>`).join('');
        } else {
          ktNotesContainer.innerHTML = `<li class="text-slate-500 italic">Không có ghi chú nhận xét nào.</li>`;
        }
      }

      document.getElementById('totalPopsCount').innerText = `${ktItems.length} POPs`;

      const ktAvgRate = ktItems.length > 0 ? (ktItems.reduce((acc, curr) => acc + curr.rate, 0) / ktItems.length).toFixed(1) : 0;
      document.getElementById('kpi-khaithac-avg-rate').innerText = `${ktAvgRate}%`;

      renderAllCharts();
    }

    function renderAllCharts() {
      Object.values(chartInstances).forEach(chart => chart.destroy());
      chartInstances = {};

      const chartDefaults = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { labels: { color: '#94a3b8', font: { family: 'Inter', size: 11 } } },
          tooltip: {
            callbacks: {
              label: function(context) {
                let label = context.dataset.label || '';
                if (label) label += ': ';
                let idx = context.dataIndex;
                let isPercent = dashboardData.vanhanh.items[idx]?.isPercent;
                if (context.parsed.y !== null) {
                  label += isPercent ? context.parsed.y + '%' : context.parsed.y;
                }
                return label;
              }
            }
          }
        },
        scales: {
          x: { 
            ticks: { 
              color: '#94a3b8',
              font: { family: 'Inter', size: 10 },
              maxRotation: 25,
              minRotation: 0
            }, 
            grid: { color: 'rgba(255, 255, 255, 0.05)' } 
          },
          y: { ticks: { color: '#64748b' }, grid: { color: 'rgba(255, 255, 255, 0.05)' } }
        }
      };

      const vhWeekLabels = dashboardData.vanhanh.weekHeaders;

      const ctxOvVh = document.getElementById('overviewVanHanhChart')?.getContext('2d');
      if (ctxOvVh && dashboardData.vanhanh.items.length > 0) {
        chartInstances.ovVh = new Chart(ctxOvVh, {
          type: 'line',
          data: {
            labels: vhWeekLabels,
            datasets: dashboardData.vanhanh.items.slice(0, 7).map((item, idx) => {
              const colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4'];
              const dataPoints = item.rawValues || item.values;
              return {
                label: item.metric,
                data: dataPoints,
                borderColor: colors[idx % colors.length],
                borderWidth: 2,
                tension: 0.3,
                fill: false
              };
            })
          },
          options: chartDefaults
        });
      }

      const ctxOvKt = document.getElementById('overviewKhaiThacPieChart')?.getContext('2d');
      if (ctxOvKt && dashboardData.khaithac.items.length > 0) {
        const totalUsed = dashboardData.khaithac.items.reduce((acc, c) => acc + c.used, 0);
        const totalFree = dashboardData.khaithac.items.reduce((acc, c) => acc + c.free, 0);

        chartInstances.ovKt = new Chart(ctxOvKt, {
          type: 'doughnut',
          data: {
            labels: ['Port Đã Sử Dụng', 'Port Còn Trống'],
            datasets: [{
              data: [totalUsed, totalFree],
              backgroundColor: ['#f59e0b', '#334155'],
              borderWidth: 0
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8' } } }
          }
        });
      }

      const ctxVhDet = document.getElementById('vanhanhDetailChart')?.getContext('2d');
      if (ctxVhDet && dashboardData.vanhanh.items.length > 0) {
        const weekColors = ['#1d4ed8', '#2563eb', '#3b82f6', '#60a5fa', '#93c5fd', '#c084fc'];
        
        chartInstances.vhDet = new Chart(ctxVhDet, {
          id: 'vanhanhDetailChart',
          type: 'bar',
          data: {
            labels: dashboardData.vanhanh.items.map(i => i.metric),
            datasets: vhWeekLabels.map((weekName, wIdx) => ({
              label: weekName,
              data: dashboardData.vanhanh.items.map(i => {
                if (i.rawValues && i.rawValues[wIdx] !== undefined) return i.rawValues[wIdx];
                return i.values ? parseFloat(i.values[wIdx]) || 0 : 0;
              }),
              backgroundColor: weekColors[wIdx % weekColors.length],
              borderRadius: 4
            }))
          },
          options: {
            ...chartDefaults,
            plugins: {
              ...chartDefaults.plugins,
              legend: {
                position: 'top',
                labels: { color: '#cbd5e1', font: { family: 'Inter', size: 12 } }
              }
            }
          }
        });
      }

      const ctxKtTop = document.getElementById('khaithacTopChart')?.getContext('2d');
      if (ctxKtTop && dashboardData.khaithac.items.length > 0) {
        const topPops = [...dashboardData.khaithac.items].sort((a, b) => b.rate - a.rate).slice(0, 5);
        chartInstances.ktTop = new Chart(ctxKtTop, {
          type: 'bar',
          data: {
            labels: topPops.map(i => i.pop),
            datasets: [{
              label: 'Tỷ Lệ Khai Thác (%)',
              data: topPops.map(i => i.rate),
              backgroundColor: '#f59e0b',
              borderRadius: 6
            }]
          },
          options: chartDefaults
        });
      }

      const ctxKtLowest = document.getElementById('khaithacLowestChart')?.getContext('2d');
      if (ctxKtLowest && dashboardData.khaithac.items.length > 0) {
        const lowestPops = [...dashboardData.khaithac.items].sort((a, b) => a.rate - b.rate).slice(0, 5);
        chartInstances.ktLowest = new Chart(ctxKtLowest, {
          type: 'bar',
          data: {
            labels: lowestPops.map(i => i.pop),
            datasets: [{
              label: 'Tỷ Lệ Khai Thác Thấp (%)',
              data: lowestPops.map(i => i.rate),
              backgroundColor: '#f43f5e',
              borderRadius: 6
            }]
          },
          options: chartDefaults
        });
      }

      const ctxKtStruct = document.getElementById('khaithacPortStructureChart')?.getContext('2d');
      if (ctxKtStruct && dashboardData.khaithac.items.length > 0) {
        const totalUsed = dashboardData.khaithac.items.reduce((acc, c) => acc + c.used, 0);
        const totalFree = dashboardData.khaithac.items.reduce((acc, c) => acc + c.free, 0);

        chartInstances.ktStruct = new Chart(ctxKtStruct, {
          type: 'pie',
          data: {
            labels: ['Port Sử Dụng', 'Port Trống'],
            datasets: [{
              data: [totalUsed, totalFree],
              backgroundColor: ['#10b981', '#475569']
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8' } } }
          }
        });
      }
    }

    function downloadSampleExcel() {
      const wb = XLSX.utils.book_new();
      const matrix = [];

      matrix[0] = ["I. VẬN HÀNH"];
      matrix[1] = ["STT", "Chỉ tiêu", "26-W37", "26-W38", "26-W39", "26-W40", "Plan", "+ / - Thực hiện"];
      matrix[2] = [1, "Số lượng sự cố (Sđ)", 24, 26, 27, 24, 31, -7];
      matrix[3] = [2, "KHG Ảnh hưởng/sự cố (Slg)", 21.00, 16.00, 17.20, 14.20, 16.00, -1.80];
      matrix[4] = [3, "Số lượng KHG ảnh hưởng (Slg)", 280, 255, 490, 325, 300, 25];
      matrix[5] = [4, "Thời gian XL SC trung bình (phút)", 88.5, 180.2, 128.0, 168.5, 110, 58.5];
      matrix[6] = [5, "Thời gian gián đoạn TB (phút)", 98.0, 182.0, 164.0, 212.0, 217, -5.0];
      matrix[7] = [6, "SAIDI (phút) / SLA (%)", "0.79%", "1.32%", "2.92%", "2.01%", "5.00%", "-2.99%"];
      matrix[8] = [7, "Tỷ lệ KHG ảnh hưởng (%)", "8.33%", "11.54%", "14.81%", "12.50%", "5.00%", "+7.50%"];
      matrix[9] = ["Nhận xét:"];
      matrix[10] = ["- Trong tuần 26-W40: Đã theo dõi đầy đủ 7 chỉ tiêu vận hành."];

      matrix[16] = ["II. TRIỂN KHAI"];
      matrix[17] = ["STT", "POP / Hạng mục", "26-W37", "26-W38", "26-W39", "26-W40", "Lũy kế 4 tuần", "Đơn vị", "Ghi chú"];
      matrix[18] = [1, "TCGF012", 0, 48, 0, 0, 48, "Port", "Hoàn thành"];
      matrix[19] = [2, "Nâng cấp Port, M1", 72, 88, 65, 100, 325, "Port", "Đang duy trì"];

      matrix[30] = ["III. KHAI THÁC"];
      matrix[31] = ["STT", "POP", "Tổng port", "Port sử dụng", "Port rỗng", "Tỷ lệ", "Ghi chú"];
      matrix[32] = [1, "TCGP001", 1570, 912, 658, "58.1%", "Tỷ lệ khai thác cao"];

      const ws = XLSX.utils.aoa_to_sheet(matrix);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      XLSX.writeFile(wb, "Book1_Sample_Format.xlsx");
    }
  </script>
</body>
</html>
"""

# Render Component lên Streamlit với chiều cao linh hoạt
components.html(HTML_CODE, height=1200, scrolling=True)
