let activeChart = null;

function drawChart(labels, data, chartType = 'pie', chartTitle = 'Grafik') {
    if (activeChart) { activeChart.destroy(); activeChart = null; }
    const chartHTML = `
        <div style="margin-bottom:10px;text-align:center;">
            <select id="chartTypeSelector" style="padding:5px 10px;">
                <option value="pie">Pasta</option>
                <option value="bar">Çubuk</option>
                <option value="line">Çizgi</option>
                <option value="doughnut">Donut</option>
            </select>
        </div>
        <canvas id="myChart" width="400" height="300"></canvas>
    `;
    openModal(chartHTML);

    const ctx = document.getElementById('myChart').getContext('2d');

    activeChart = new Chart(ctx, {
        type: chartType,
        data: {
            labels: labels,
            datasets: [{
                label: chartTitle,
                data: data,
                backgroundColor: [
                    '#ff6384', '#36a2eb', '#ffce56', '#4bc0c0', '#9966ff', '#ff9f40'
                ],
                borderWidth: 2
            }]
        },
        options: {
            responsive: false,
            plugins: {
                legend: { display: true, position: 'bottom' },
                title: { display: true, text: chartTitle },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.raw;
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percent = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                            // Noktadan sonra üç basamak (50.000 gibi)
                            const valueString = value.toLocaleString('tr-TR');
                            return `${context.label}: ${valueString} (%${percent})`;
                        }
                    }
                }
            }
        }
    });

    document.getElementById('chartTypeSelector').value = chartType;
    document.getElementById('chartTypeSelector').addEventListener('change', function () {
        const selectedType = this.value;
        drawChart(labels, data, selectedType, chartTitle);
    });
}
window.drawChart = drawChart;
