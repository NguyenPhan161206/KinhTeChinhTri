window.growthChart = null;
window.cvPieChart = null;

function initCharts() {
    const ctxGrowth = document.getElementById('growthChart').getContext('2d');
    window.growthChart = new Chart(ctxGrowth, {
        type: 'line',
        data: {
            labels: Array.from({length: 12}, (_, i) => `Tháng ${i+1}`),
            datasets: [
                { label: 'Giá trị m (Triệu)', data: Array.from({length: 12}, (_, i) => calculateFinancials(i + 1).mVal), borderColor: '#4f46e5', backgroundColor: 'rgba(79, 70, 229, 0.1)', fill: true, tension: 0.4, yAxisID: 'y' },
                { label: 'Nhân sự v', data: Array.from({length: 12}, (_, i) => calculateFinancials(i + 1).totalStaff), type: 'bar', backgroundColor: '#e2e8f0', borderRadius: 6, yAxisID: 'y1' }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { position: 'left', grid: { display: false } }, y1: { position: 'right', grid: { display: false }, min: 25, max: 45 } } }
    });

    const pieCanvas = document.getElementById('cvPieChart');
    if (pieCanvas) {
        const ctxPie = pieCanvas.getContext('2d');
        window.cvPieChart = new Chart(ctxPie, {
            type: 'doughnut',
            data: {
                labels: ['Tư bản bất biến (c)', 'Tư bản khả biến (v)'],
                datasets: [{
                    data: [1400, 600],
                    backgroundColor: ['#1e293b', '#4f46e5'],
                    borderWidth: 0,
                    hoverOffset: 10
                }]
            },
            options: { 
                cutout: '70%',
                plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, padding: 20 } } }
            }
        });
    }
}