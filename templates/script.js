let currentTab = "attendance";
let currentFilter = "hôm nay";
let currentSearch = "";
let currentSort = { column: 'CreationTime', order: 'desc' };
let startDate = "", endDate = "";
let rawData = [];
let displayData = [];
let currentPage = 1;
let rowsPerPage = 10;
let currentLeaveFilter = "hôm nay";
let currentLeaveSearch = "";
let currentLeaveSort = { column: 'CheckinTime', order: 'desc' };
let leaveStartDate = "", leaveEndDate = "";
let leaveDateType = "CheckinTime";
let rawLeaveData = [];
let currentLeavePage = 1;
let leaveRowsPerPage = 10;

function showMessage(msg, type="success", isModal=false) {
  const el = isModal ? document.getElementById("modalMessage") : document.getElementById("messageArea");
  el.innerHTML = `<div class="msg ${type}">${msg}</div>`;
  el.style.display = "block";
  setTimeout(() => {
    el.innerHTML = "";
    el.style.display = "none";
  }, 4000);
}

function showDashboard() {
  document.getElementById("loginModal").style.display = "none";
  document.getElementById("greeting").style.display = "block";
  document.getElementById("logoutBtn").style.display = "inline-block";
  document.getElementById("dataTabs").style.display = "flex";
  switchTab(currentTab);
}

function hideDashboard() {
  document.getElementById("greeting").style.display = "none";
  document.getElementById("logoutBtn").style.display = "none";
  document.getElementById("dataTabs").style.display = "none";
  document.getElementById("leaveSection").style.display = "none";
  document.getElementById("filters").style.display = "none";
  document.getElementById("dateFilter").style.display = "none";
  document.getElementById("searchFilter").style.display = "none";
  document.getElementById("tableContainer").style.display = "none";
  document.getElementById("pagination").style.display = "none";
  document.getElementById("dataBody").innerHTML = "";
  document.getElementById("leaveDataBody").innerHTML = "";
}

function resetExportOptions() {
  document.getElementById('exportOptionsAttendance').value = "";
  document.getElementById('exportOptionsLeave').value = "";
  document.getElementById("exportMonthAttendance").style.display = "none";
  document.getElementById("exportYearAttendance").style.display = "none";
  document.getElementById("exportMonthLeave").style.display = "none";
  document.getElementById("exportYearLeave").style.display = "none";
  const labelAtt = document.getElementById("exportDateLabelAttendance");
  const labelLeave = document.getElementById("exportDateLabelLeave");
  if (labelAtt) labelAtt.style.display = "none";
  if (labelLeave) labelLeave.style.display = "none";
  const currentMonth = new Date().getMonth() + 1;
  const currentYear = new Date().getFullYear();
  document.getElementById('exportMonthAttendance').value = currentMonth;
  document.getElementById('exportYearAttendance').value = currentYear;
  document.getElementById('exportMonthLeave').value = currentMonth;
  document.getElementById('exportYearLeave').value = currentYear;
}

function switchTab(tab) {
  currentTab = tab;
  document.getElementById("tabAttendance").classList.toggle("active", tab === "attendance");
  document.getElementById("tabLeave").classList.toggle("active", tab === "leave");
  resetExportOptions();
  if (tab === "attendance") {
    document.getElementById("leaveSection").style.display = "none";
    document.getElementById("filters").style.display = "flex";
    document.getElementById("dateFilter").style.display = "flex";
    document.getElementById("searchFilter").style.display = "flex";
    document.getElementById("tableContainer").style.display = "block";
    document.getElementById("pagination").style.display = "block";
    if (!rawData.length) loadData(); else renderTable();
  } else {
    document.getElementById("filters").style.display = "none";
    document.getElementById("dateFilter").style.display = "none";
    document.getElementById("searchFilter").style.display = "none";
    document.getElementById("tableContainer").style.display = "none";
    document.getElementById("pagination").style.display = "none";
    document.getElementById("leaveSection").style.display = "block";
    if (!rawLeaveData.length) loadLeaveData(); else renderLeaveTable();
  }
}

async function login() {
  const email = document.getElementById("loginEmail").value.trim();
  const pass = document.getElementById("loginPass").value;
  if (!email || !pass) return showMessage("⚠️ Nhập email và mật khẩu", "error", true);
  const formData = new FormData();
  formData.append("email", email);
  formData.append("password", pass);
  const res = await fetch("/login", { method: "POST", body: formData });
  const data = await res.json();
  if (res.ok && data.success) {
    localStorage.setItem("adminEmail", email);
    localStorage.setItem("adminUser", data.username || data.email);
    document.getElementById("greeting").innerHTML = `Xin chào, ${data.username || email}!`;
    showDashboard();
    showMessage("✅ Đăng nhập thành công", "success");
    applyFilter("hôm nay");
  } else {
    showMessage(data.message || "🚫 Sai thông tin", "error", true);
  }
}

function logout() {
  localStorage.clear();
  document.getElementById("loginEmail").value = "";
  document.getElementById("loginPass").value = "";
  document.getElementById("loginModal").style.display = "flex";
  hideDashboard();
  showMessage("👋 Đã đăng xuất", "success");
}

document.getElementById("loginBtn").onclick = login;
document.getElementById("logoutBtn").onclick = logout;

function processDataWithHours(data) {
    if (!data || data.length === 0) return [];
    function parseTime(timeStr) {
        if (!timeStr) return null;
        const [datePart, timePart] = timeStr.split(' ');
        if (!datePart || !timePart) return null;
        const [day, month, year] = datePart.split('/');
        const [hours, minutes, seconds] = timePart.split(':');
        try {
            return new Date(year, month - 1, day, hours, minutes, seconds);
        } catch (e) {
            return null;
        }
    }
    const employeesData = data.reduce((acc, record) => {
        acc[record.EmployeeId] = acc[record.EmployeeId] || [];
        acc[record.EmployeeId].push(record);
        return acc;
    }, {});
    const augmentedData = [];
    for (const empId in employeesData) {
        const records = employeesData[empId];
        const monthsData = records.reduce((acc, record) => {
            const date = record.CheckinDate;
            if (!date) return acc;
            const monthKey = date.substring(0, 7);
            acc[monthKey] = acc[monthKey] || [];
            acc[monthKey].push(record);
            return acc;
        }, {});
        for (const monthKey in monthsData) {
            let monthlyTotalSeconds = 0;
            const monthRecords = monthsData[monthKey];
            const daysData = monthRecords.reduce((acc, record) => {
                acc[record.CheckinDate] = acc[record.CheckinDate] || [];
                acc[record.CheckinDate].push(record);
                return acc;
            }, {});
            const sortedDays = Object.keys(daysData).sort();
            const dailyHoursMap = {};
            for (const day of sortedDays) {
                let firstCheckin = null, lastCheckout = null;
                daysData[day].forEach(rec => {
                    const recTime = parseTime(rec.CreationTime);
                    if (!recTime) return;
                    if (rec.CheckType === 'checkin' && (!firstCheckin || recTime < firstCheckin)) {
                        firstCheckin = recTime;
                    } else if (rec.CheckType === 'checkout' && (!lastCheckout || recTime > lastCheckout)) {
                        lastCheckout = recTime;
                    }
                });
                let dailySeconds = 0;
                if (firstCheckin && lastCheckout && lastCheckout > firstCheckin) {
                    dailySeconds = (lastCheckout - firstCheckin) / 1000;
                }
                monthlyTotalSeconds += dailySeconds;
                const dailyH = Math.floor(dailySeconds / 3600);
                const dailyM = Math.floor((dailySeconds % 3600) / 60);
                const monthlyH = Math.floor(monthlyTotalSeconds / 3600);
                const monthlyM = Math.floor((monthlyTotalSeconds % 3600) / 60);
                dailyHoursMap[day] = {
                    daily: `${dailyH}h ${dailyM}m`,
                    monthly: `${monthlyH}h ${monthlyM}m`,
                    _dailySeconds: dailySeconds,
                    _monthlySeconds: monthlyTotalSeconds,
                };
            }
            monthRecords.forEach(rec => {
                const hours = dailyHoursMap[rec.CheckinDate];
                augmentedData.push({
                    ...rec,
                    DailyHours: hours ? hours.daily : '0h 0m',
                    MonthlyHours: hours ? hours.monthly : 'N/A',
                    _dailySeconds: hours ? hours._dailySeconds : 0,
                    _monthlySeconds: hours ? hours._monthlySeconds : 0
                });
            });
        }
    }
    return augmentedData;
}

async function loadData() {
  const email = localStorage.getItem("adminEmail");
  if (!email) {
    document.getElementById("loginModal").style.display = "flex";
    return showMessage("🚫 Chưa đăng nhập!", "error");
  }
  let url = `/api/attendances?email=${encodeURIComponent(email)}&filter=${encodeURIComponent(currentFilter)}`;
  if (startDate && endDate) url += `&startDate=${startDate}&endDate=${endDate}`;
  if (currentSearch) url += `&search=${encodeURIComponent(currentSearch)}`;
  const res = await fetch(url);
  const data = await res.json();
  if (!res.ok) return showMessage(data.error || "❌ Lỗi tải dữ liệu", "error");
  rawData = data;
  displayData = processDataWithHours(rawData);
  currentPage = 1;
  renderTable();
}

function renderTable() {
  const tbody = document.getElementById("dataBody");
  tbody.innerHTML = "";
  let data = [...displayData];
  if (currentSort.column && currentSort.order) {
    data.sort((a,b) => {
        let col = currentSort.column;
        if (col === 'DailyHours') col = '_dailySeconds';
        if (col === 'MonthlyHours') col = '_monthlySeconds';
        let A = a[col] ?? (typeof a[col] === 'number' ? 0 : "");
        let B = b[col] ?? (typeof b[col] === 'number' ? 0 : "");
        if (typeof A === 'string') A = A.toLowerCase();
        if (typeof B === 'string') B = B.toLowerCase();
        if (A < B) return currentSort.order === "asc" ? -1 : 1;
        if (A > B) return currentSort.order === "asc" ? 1 : -1;
        return 0;
    });
  }
  const startIdx = (currentPage-1)*rowsPerPage;
  const pageData = data.slice(startIdx, startIdx+rowsPerPage);
  pageData.forEach(r => {
    const img = r.FaceImage || r.PhotoURL ? `<img src="${r.FaceImage || r.PhotoURL}" class="checkin-photo">` : "";
    const mapLink = (r.Latitude && r.Longitude)
      ? `<a href="https://maps.google.com/?q=${r.Latitude},${r.Longitude}" target="_blank" class="map-link">Xem</a>`
      : "";
    const checkTypeFormatted = r.CheckType === 'checkin' ? 'Check In' : (r.CheckType === 'checkout' ? 'Check Out' : (r.CheckType || ""));
    tbody.innerHTML += `
      <tr>
        <td>${r.EmployeeId||""}</td>
        <td>${r.EmployeeName||""}</td>
        <td>${r.CheckinDate||""}</td>
        <td>${r.CheckinTime||""}</td>
        <td>${checkTypeFormatted}</td>
        <td>${r.ProjectId||""}</td>
        <td>${Array.isArray(r.Tasks) ? r.Tasks.join(', ') : (r.Tasks || "")}</td>
        <td>${img}</td>
        <td>${r.Address||""}</td>
        <td>${mapLink}</td>
        <td>${r.CheckinNote||""}</td>
      </tr>
    `;
  });
  renderPagination(data.length);
  highlightFilterButton();
}

async function loadLeaveData() {
  const email = localStorage.getItem("adminEmail");
  if (!email) {
    document.getElementById("loginModal").style.display = "flex";
    return showMessage("🚫 Chưa đăng nhập!", "error");
  }
  let url = `/api/leaves?email=${encodeURIComponent(email)}&filter=${encodeURIComponent(currentLeaveFilter)}&dateType=${encodeURIComponent(leaveDateType)}`;
  if (leaveStartDate && leaveEndDate) url += `&startDate=${leaveStartDate}&endDate=${leaveEndDate}`;
  if (currentLeaveSearch) url += `&search=${encodeURIComponent(currentLeaveSearch)}`;
  const res = await fetch(url);
  const data = await res.json();
  if (!res.ok) return showMessage(data.error || "❌ Lỗi tải dữ liệu nghỉ phép", "error");
  rawLeaveData = data;
  currentLeavePage = 1;
  renderLeaveTable();
}

function renderLeaveTable() {
  const tbody = document.getElementById("leaveDataBody");
  tbody.innerHTML = "";
  let data = [...rawLeaveData];
  if (currentLeaveSort.column && currentLeaveSort.order) {
    data.sort((a,b) => {
      let A = a[currentLeaveSort.column]||"", B = b[currentLeaveSort.column]||"";
      if (A<B) return currentLeaveSort.order==="asc" ? -1 : 1;
      if (A>B) return currentLeaveSort.order==="asc" ? 1 : -1;
      return 0;
    });
  }
  const startIdx = (currentLeavePage-1)*leaveRowsPerPage;
  const pageData = data.slice(startIdx, startIdx+leaveRowsPerPage);
  pageData.forEach(r => {
    tbody.innerHTML += `
      <tr>
        <td>${r.EmployeeId||""}</td>
        <td>${r.EmployeeName||""}</td>
        <td>${r.CheckinDate||""}</td>
        <td>${r.CheckinTime||""}</td>
        <td>${Array.isArray(r.Tasks) ? r.Tasks.join(', ') : (r.Tasks || "")}</td>
        <td>${r.ApprovalDate1||""}</td>
        <td>${r.Status1||""}</td>
        <td>${r.ApprovalDate2||""}</td>
        <td>${r.Status2||""}</td>
        <td>${r.Note||""}</td>
      </tr>
    `;
  });
  renderLeavePagination(data.length);
  highlightLeaveFilterButton();
}

function renderPagination(total) {
  const pages = Math.ceil(total / rowsPerPage);
  const pagDiv = document.getElementById("pagination");
  pagDiv.innerHTML = "";
  for (let i=1; i<=pages; i++) {
    pagDiv.innerHTML += `<button class="${i===currentPage?'active':''}" onclick="gotoPage(${i})">${i}</button>`;
  }
}

function renderLeavePagination(total) {
  const pages = Math.ceil(total / leaveRowsPerPage);
  const pagDiv = document.getElementById("leavePagination");
  pagDiv.innerHTML = "";
  for (let i=1; i<=pages; i++) {
    pagDiv.innerHTML += `<button class="${i===currentLeavePage?'active':''}" onclick="gotoLeavePage(${i})">${i}</button>`;
  }
}

function gotoPage(page) { currentPage = page; renderTable(); }
function gotoLeavePage(page) { currentLeavePage = page; renderLeaveTable(); }

function sortTable(col) {
  if (currentSort.column===col) {
    currentSort.order = currentSort.order==="asc" ? "desc" : (currentSort.order==="desc"?null:"asc");
  } else {
    currentSort.column = col;
    currentSort.order = "asc";
  }
  document.querySelectorAll("#dataTable th").forEach(th => th.classList.remove("asc","desc"));
  if (currentSort.order)
    document.querySelector(`#dataTable th[onclick="sortTable('${col}')"]`).classList.add(currentSort.order);
  renderTable();
}

function sortLeaveTable(col) {
  if (currentLeaveSort.column===col) {
    currentLeaveSort.order = currentLeaveSort.order==="asc" ? "desc" : (currentLeaveSort.order==="desc"?null:"asc");
  } else {
    currentLeaveSort.column = col;
    currentLeaveSort.order = "asc";
  }
  document.querySelectorAll("#leaveTable th").forEach(th => th.classList.remove("asc","desc"));
  if (currentLeaveSort.order)
    document.querySelector(`#leaveTable th[onclick="sortLeaveTable('${col}')"]`).classList.add(currentLeaveSort.order);
  renderLeaveTable();
}

function applyFilter(f) {
  currentFilter = f;
  startDate = endDate = "";
  document.getElementById("startDate").value = "";
  document.getElementById("endDate").value = "";
  loadData();
}

function applyLeaveFilter(f) {
  currentLeaveFilter = f;
  leaveStartDate = leaveEndDate = "";
  document.getElementById("leaveStartDate").value = "";
  document.getElementById("leaveEndDate").value = "";
  loadLeaveData();
}

function applyLeaveDateTypeFilter() {
  leaveDateType = document.getElementById("leaveDateTypeFilter").value;
  loadLeaveData();
}

function applyDateRange() {
  startDate = document.getElementById("startDate").value;
  endDate = document.getElementById("endDate").value;
  currentFilter = "custom";
  loadData();
}

function applyLeaveDateRange() {
  leaveStartDate = document.getElementById("leaveStartDate").value;
  leaveEndDate = document.getElementById("leaveEndDate").value;
  currentLeaveFilter = "custom";
  loadLeaveData();
}

function applySearch() {
  currentSearch = document.getElementById("searchBox").value.trim();
  loadData();
}

function applyLeaveSearch() {
  currentLeaveSearch = document.getElementById("leaveSearchBox").value.trim();
  loadLeaveData();
}

function handleExportClick() {
  const email = localStorage.getItem("adminEmail");
  let exportOptionsEl;
  let url = "";
  if (!email) {
    document.getElementById("loginModal").style.display = "flex";
    return showMessage("⚠️ Chưa đăng nhập", "error");
  }
  let month, year;
  if (currentTab === "attendance") {
    exportOptionsEl = document.getElementById('exportOptionsAttendance');
    month = document.getElementById('exportMonthAttendance').value;
    year = document.getElementById('exportYearAttendance').value;
  } else {
    exportOptionsEl = document.getElementById('exportOptionsLeave');
    month = document.getElementById('exportMonthLeave').value;
    year = document.getElementById('exportYearLeave').value;
  }
  const exportType = exportOptionsEl ? exportOptionsEl.value : "";
  if (exportType === "") {
    showMessage("⚠️ Vui lòng chọn loại dữ liệu cần xuất!", "error");
    return;
  }
  if (!month || !year) {
    showMessage("⚠️ Vui lòng chọn tháng và năm!", "error");
    return;
  }
  if (exportType === "attendance") {
    url = `/api/export-excel?email=${encodeURIComponent(email)}&filter=${encodeURIComponent(currentFilter)}`;
    if (startDate && endDate) url += `&startDate=${startDate}&endDate=${endDate}`;
    if (currentSearch) url += `&search=${encodeURIComponent(currentSearch)}`;
  } else if (exportType === "leave") {
    url = `/api/export-leaves-excel?email=${encodeURIComponent(email)}&filter=${encodeURIComponent(currentLeaveFilter)}&dateType=${encodeURIComponent(leaveDateType)}`;
    if (leaveStartDate && leaveEndDate) url += `&startDate=${leaveStartDate}&endDate=${leaveEndDate}`;
    if (currentLeaveSearch) url += `&search=${encodeURIComponent(currentLeaveSearch)}`;
  } else if (exportType === "combined") {
    let activeFilter, activeStartDate, activeEndDate, activeSearch, activeDateType;
    if (currentTab === 'attendance') {
        activeFilter = currentFilter;
        activeStartDate = startDate;
        activeEndDate = endDate;
        activeSearch = currentSearch;
        activeDateType = '';
    } else {
        activeFilter = currentLeaveFilter;
        activeStartDate = leaveStartDate;
        activeEndDate = leaveEndDate;
        activeSearch = currentLeaveSearch;
        activeDateType = leaveDateType;
    }
    url = `/api/export-combined-excel?email=${encodeURIComponent(email)}`;
    url += `&filter=${encodeURIComponent(activeFilter)}`;
    if (activeStartDate && activeEndDate) {
        url += `&startDate=${activeStartDate}&endDate=${activeEndDate}`;
    }
    if (activeSearch) {
        url += `&search=${encodeURIComponent(activeSearch)}`;
    }
    if (activeDateType) {
        url += `&dateType=${encodeURIComponent(activeDateType)}`;
    }
  }
  if (month) url += `&month=${encodeURIComponent(month)}`;
  if (year) url += `&year=${encodeURIComponent(year)}`;
  if (url) {
    window.location.href = url;
    resetExportOptions();
  }
}

function updateRowsPerPage() {
  rowsPerPage = parseInt(document.getElementById("rowsPerPage").value);
  currentPage = 1;
  renderTable();
}

function updateLeaveRowsPerPage() {
  leaveRowsPerPage = parseInt(document.getElementById("leaveRowsPerPage").value);
  currentLeavePage = 1;
  renderLeaveTable();
}

function refreshData() { loadData(); }
function refreshLeaveData() { loadLeaveData(); }

function resetFilter() {
  currentSearch = "";
  startDate = endDate = "";
  currentFilter = "hôm nay";
  document.getElementById("searchBox").value = "";
  document.getElementById("startDate").value = "";
  document.getElementById("endDate").value = "";
  highlightFilterButton();
  loadData();
  resetExportOptions();
}

function resetLeaveFilter() {
  currentLeaveSearch = "";
  leaveStartDate = leaveEndDate = "";
  currentLeaveFilter = "hôm nay";
  leaveDateType = "CheckinTime";
  document.getElementById("leaveSearchBox").value = "";
  document.getElementById("leaveStartDate").value = "";
  document.getElementById("leaveEndDate").value = "";
  document.getElementById("leaveDateTypeFilter").value = "CheckinTime";
  highlightLeaveFilterButton();
  loadLeaveData();
  resetExportOptions();
}

function highlightFilterButton() {
  document.querySelectorAll("#filters button").forEach(btn => btn.classList.remove("active"));
  const map = {
    "hôm nay":"btn-homnay","tuần":"btn-tuan","tháng":"btn-thang",
    "năm":"btn-nam","tất cả":"btn-tatca"
  };
  if (map[currentFilter]) document.getElementById(map[currentFilter]).classList.add("active");
}

function highlightLeaveFilterButton() {
  document.querySelectorAll("#leaveSection .controls:nth-child(1) button").forEach(btn => btn.classList.remove("active"));
  const map = {
    "hôm nay":"leave-btn-homnay","tuần":"leave-btn-tuan","tháng":"leave-btn-thang",
    "năm":"leave-btn-nam","tất cả":"leave-btn-tatca"
  };
  if (map[currentLeaveFilter]) document.getElementById(map[currentLeaveFilter]).classList.add("active");
}

function populateYears() {
  const currentYear = new Date().getFullYear();
  const years = [];
  for (let y = currentYear - 5; y <= currentYear; y++) {
    years.push(y);
  }
  const yearOptions = years.map(y => `<option value="${y}">${y}</option>`).join('');
  document.getElementById('exportYearAttendance').innerHTML += yearOptions;
  document.getElementById('exportYearLeave').innerHTML += yearOptions;
}

window.onload = () => {
  populateYears();
  const currentMonth = new Date().getMonth() + 1;
  const currentYear = new Date().getFullYear();
  document.getElementById('exportMonthAttendance').value = currentMonth;
  document.getElementById('exportYearAttendance').value = currentYear;
  document.getElementById('exportMonthLeave').value = currentMonth;
  document.getElementById('exportYearLeave').value = currentYear;
  const email = localStorage.getItem("adminEmail");
  if (email) {
    const username = localStorage.getItem("adminUser");
    document.getElementById("greeting").innerHTML = `Xin chào, ${username}!`;
    showDashboard();
    applyFilter("hôm nay");
  } else {
    hideDashboard();
    document.getElementById("loginModal").style.display = "flex";
  }
};

document.getElementById("loginPass").addEventListener("keypress", (e) => {
  if (e.key === "Enter") login();
});

document.getElementById("exportOptionsAttendance").addEventListener("change", function() {
    const monthSelect = document.getElementById("exportMonthAttendance");
    const yearSelect = document.getElementById("exportYearAttendance");
    const dateLabel = document.getElementById("exportDateLabelAttendance");
    if (this.value !== "") {
        dateLabel.style.display = "inline-block";
        monthSelect.style.display = "inline-block";
        yearSelect.style.display = "inline-block";
    } else {
        dateLabel.style.display = "none";
        monthSelect.style.display = "none";
        yearSelect.style.display = "none";
    }
});

document.getElementById("exportOptionsLeave").addEventListener("change", function() {
    const monthSelect = document.getElementById("exportMonthLeave");
    const yearSelect = document.getElementById("exportYearLeave");
    const dateLabel = document.getElementById("exportDateLabelLeave");
    if (this.value !== "") {
        dateLabel.style.display = "inline-block";
        monthSelect.style.display = "inline-block";
        yearSelect.style.display = "inline-block";
    } else {
        dateLabel.style.display = "none";
        monthSelect.style.display = "none";
        yearSelect.style.display = "none";
    }
});
