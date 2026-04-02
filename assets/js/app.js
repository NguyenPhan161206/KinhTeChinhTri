document.addEventListener('DOMContentLoaded', () => {
    // Khởi tạo các slider nếu tồn tại
    const monthSlider = document.getElementById('month-slider');
    if (monthSlider) {
        monthSlider.addEventListener('input', e => updateUI(parseInt(e.target.value)));
    }

    const kpiMonthSlider = document.getElementById('kpi-month-slider');
    if (kpiMonthSlider) {
        kpiMonthSlider.addEventListener('input', e => updateKPIUI(parseInt(e.target.value)));
    }

    // Mặc định hiển thị view giới thiệu
    changeView('about');
});