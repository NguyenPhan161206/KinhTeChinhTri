window.macroState = { inflation: 0, interest: 0 };
window.currentKPIDept = "Marketing";
window.currentDailyDept = "Marketing";

function updateUI(month) {
    const data = window.staffData[month];
    const fin = calculateFinancials(month);
    document.getElementById('month-display').textContent = month.toString().padStart(2, '0');
    document.getElementById('stat-staff').textContent = fin.totalStaff;
    document.getElementById('stat-m').textContent = Math.round(fin.mVal).toLocaleString() + " Tr";
    document.getElementById('stat-profit').textContent = ((fin.mVal / 2000) * 100).toFixed(1) + "%";
    
    const card = document.getElementById('month-content-card');
    let riskHTML = data.risk !== "Bình thường" ? `<div class="mt-4 p-4 rounded-xl bg-rose-50 border border-rose-100"><p class="text-rose-600 font-bold text-xs mb-1">⚠️ RỦI RO: ${data.risk.toUpperCase()}</p><p class="text-[10px] text-rose-800">${data.desc}</p></div>` : "";
    
    card.innerHTML = `
        <div class="mb-4"><span class="text-[10px] font-bold text-slate-400 uppercase">Sự kiện</span><h4 class="text-xl font-bold text-slate-800">${data.event}</h4></div>
        
        <div class="grid grid-cols-2 gap-4 mb-6">
            <div class="p-4 rounded-xl bg-white border border-slate-100 shadow-sm">
                <p class="text-[10px] text-slate-400 font-bold uppercase mb-1">Tổng Thặng Dư (m)</p>
                <p class="text-xl font-black text-indigo-600">${Math.round(fin.totalSurplus).toLocaleString()} Tr</p>
                <p class="text-[9px] text-slate-500 mt-1 italic">Tạo ra từ sức lao động nhân sự.</p>
            </div>
            <div class="p-4 rounded-xl bg-rose-50 border border-rose-100 shadow-sm">
                <p class="text-[10px] text-rose-400 font-bold uppercase mb-1">Lợi tức (z) trả Ngân hàng</p>
                <p class="text-xl font-black text-rose-600">-${Math.round(fin.interestZ).toLocaleString()} Tr</p>
                <p class="text-[9px] text-rose-500 mt-1 italic">Chi phí sử dụng đòn bẩy tài chính.</p>
            </div>
        </div>

        <div class="bg-indigo-50 p-4 rounded-xl border border-indigo-100">
            <h5 class="text-[10px] font-bold text-indigo-900 mb-1 uppercase tracking-widest">Góc nhìn Mác-xít & Đòn bẩy</h5>
            <p class="text-[11px] text-indigo-700 italic">Doanh nghiệp vay thêm 4 Tỷ VNĐ để mở rộng quy mô. Lợi nhuận doanh nghiệp thực nhận là <strong>m_dn = ${Math.round(fin.mVal)} Tr</strong> sau khi đã trích <strong>z</strong> cho Ngân hàng.</p>
        </div>
        ${riskHTML}
    `;

    if (window.cvPieChart) {
        window.cvPieChart.data.datasets[0].data = [fin.cVal, fin.totalSalary];
        window.cvPieChart.update();
    }

    updateSalaryTable(month);
}

function updateSalaryTable(month) {
    const d = window.staffData[month];
    const tbody = document.getElementById('salary-table-body');
    if (!tbody) return;

    // Cấu trúc lương (Triệu VNĐ) - Khớp 100% với Excel
    const salaryConfig = {
        'Marketing': [
            { pos: 'Manager', base: 40, kpi: 10, count: 1 },
            { pos: 'Performance', base: 25, kpi: 5, count: Math.ceil(d.mkt * 0.3) },
            { pos: 'Content', base: 15, kpi: 3, count: Math.ceil(d.mkt * 0.4) },
            { pos: 'Creative', base: 15, kpi: 3, count: Math.floor(d.mkt * 0.3) }
        ],
        'Sales': [
            { pos: 'Team lead', base: 30, kpi: 10, count: 1 },
            { pos: 'Defendence', base: 20, kpi: 5, count: 1 },
            { pos: 'Sales member', base: 15, kpi: 5, count: d.sales - 2 }
        ],
        'Tech': [
            { pos: 'Automation', base: 35, kpi: 7, count: Math.ceil(d.tech * 0.3) },
            { pos: 'Fullstack', base: 30, kpi: 6, count: Math.ceil(d.tech * 0.5) },
            { pos: 'Devops', base: 30, kpi: 6, count: Math.floor(d.tech * 0.2) }
        ],
        'Ops': [
            { pos: 'Ops manager', base: 30, kpi: 5, count: 1 },
            { pos: 'Delivery', base: 15, kpi: 2, count: Math.ceil((d.ops-1) * 0.5) },
            { pos: 'Support', base: 15, kpi: 2, count: Math.floor((d.ops-1) * 0.5) }
        ]
    };

    let html = "";
    for (const [dept, roles] of Object.entries(salaryConfig)) {
        roles.forEach((role, index) => {
            const total = (role.base + role.kpi) * role.count;
            html += `
                <tr class="border-b border-slate-50 hover:bg-slate-50/50 transition-colors">
                    <td class="py-4 text-indigo-600 font-bold">${index === 0 ? dept : ''}</td>
                    <td class="py-4 text-slate-700">${role.pos}</td>
                    <td class="py-4 text-center text-slate-500">${role.count}</td>
                    <td class="py-4 text-right text-slate-700 font-mono">${role.base.toLocaleString()} Tr</td>
                    <td class="py-4 text-right text-emerald-600 font-mono">${role.kpi.toLocaleString()} Tr</td>
                    <td class="py-4 text-right text-indigo-600 font-bold font-mono">${total.toLocaleString()} Tr</td>
                </tr>
            `;
        });
    }
    tbody.innerHTML = html;
}

function setKPIDept(dept) {
    window.currentKPIDept = dept;
    document.querySelectorAll('.dept-btn').forEach(b => {
        b.classList.remove('toggle-active');
        b.classList.add('bg-slate-800', 'text-slate-400');
    });
    const btn = document.getElementById('dept-' + dept);
    btn.classList.add('toggle-active');
    btn.classList.remove('bg-slate-800', 'text-slate-400');
    updateKPIUI(parseInt(document.getElementById('kpi-month-slider').value));
}

function updateKPIUI(month) {
    document.getElementById('kpi-month-display').textContent = month.toString().padStart(2, '0');
    const deptData = window.kpiData[month][window.currentKPIDept];
    const grid = document.getElementById('kpi-content-grid');
    
    grid.innerHTML = deptData.map(item => `
        <div class="glass-card p-6 rounded-3xl shadow-sm border-t-4 border-indigo-500 hover:shadow-md transition-all">
            <div class="flex justify-between items-start mb-4">
                <h4 class="font-bold text-slate-800 text-lg">${item.position}</h4>
                <span class="kpi-tag">KPI THÁNG ${month}</span>
            </div>
            <div class="space-y-3">
                ${item.kpi.split(/\\n|\n/).map(line => line.trim() ? `
                    <div class="flex gap-3 items-start">
                        <span class="text-indigo-500 mt-1">✓</span>
                        <p class="text-sm text-slate-600 leading-relaxed">${line.trim()}</p>
                    </div>
                ` : '').join('')}
            </div>
        </div>
    `).join('');
}

function setDailyDept(dept) {
    window.currentDailyDept = dept;
    document.querySelectorAll('.daily-dept-btn').forEach(b => {
        b.classList.remove('toggle-active');
        b.classList.add('bg-slate-800', 'text-slate-400');
    });
    const btn = document.getElementById('daily-dept-' + dept);
    btn.classList.add('toggle-active');
    btn.classList.remove('bg-slate-800', 'text-slate-400');
    updateDailyUI();
}

function updateDailyUI() {
    const deptData = window.dailyJobData[window.currentDailyDept];
    const grid = document.getElementById('daily-content-grid');
    grid.innerHTML = deptData.map(item => `
        <div class="glass-card p-6 rounded-3xl shadow-sm border-t-4 border-emerald-500 hover:shadow-md transition-all">
            <h4 class="font-bold text-slate-800 text-lg mb-4">${item.position}</h4>
            <div class="space-y-3">
                ${item.tasks.split('\n').map(task => `
                    <div class="flex gap-3 items-start">
                        <span class="text-emerald-500 mt-1">▹</span>
                        <p class="text-sm text-slate-600 leading-relaxed">${task}</p>
                    </div>
                `).join('')}
            </div>
        </div>
    `).join('');
}

function setMacro(type, val) {
    window.macroState[type] = val;
    document.querySelectorAll(`button[id^="${type.substring(0,3)}"]`).forEach(b => b.classList.remove('toggle-active', 'bg-white'));
    document.getElementById(`${type.substring(0,3)}-${val}`).classList.add('toggle-active');
    
    const desc = document.getElementById('macro-impact-desc');
    if (desc) {
        if(window.macroState.inflation === 1 && window.macroState.interest === 1) desc.textContent = "Cảnh báo: Lạm phát cao kéo lương v tăng, cộng thêm lãi suất cao làm bào mòn lợi nhuận m.";
        else if(window.macroState.inflation === 1) desc.textContent = "Lạm phát cao khiến chi phí tái sản xuất sức lao động (v) tăng 15%.";
        else if(window.macroState.interest === 1) desc.textContent = "Lãi suất vay 12% khiến một phần lớn thặng dư m bị chuyển hóa thành lợi tức z cho ngân hàng.";
        else desc.textContent = "Kịch bản ổn định: Lạm phát thấp và Lãi suất ưu đãi giúp tối ưu hóa tích lũy tư bản.";
    }

    if (window.growthChart) {
        window.growthChart.data.datasets[0].data = Array.from({length: 12}, (_, i) => calculateFinancials(i + 1).mVal);
        window.growthChart.update();
    }
    updateUI(parseInt(document.getElementById('month-slider').value));
}

function changeView(viewId) {
    document.querySelectorAll('section').forEach(s => s.classList.add('hidden'));
    document.getElementById('view-' + viewId).classList.remove('hidden');
    document.querySelectorAll('.nav-btn').forEach(b => {
        b.classList.remove('bg-indigo-600', 'text-white', 'shadow-lg');
        b.classList.add('text-slate-400');
    });
    document.getElementById('nav-' + viewId).classList.add('bg-indigo-600', 'text-white', 'shadow-lg');
    
    if(viewId === 'plan' && !window.growthChart) { 
        setTimeout(() => {
            initCharts();
            updateUI(1);
        }, 100); 
    }
    
    if(viewId === 'kpi') {
        updateKPIUI(parseInt(document.getElementById('kpi-month-slider').value));
    }
    if(viewId === 'daily') {
        updateDailyUI();
    }
}