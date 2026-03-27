// Lógica Principal del Dashboard

let dashboardData = {};
let coursesChartInstance = null;

// Configuración visual de Chart.js para temas oscuros
Chart.defaults.color = '#94a3b8';
Chart.defaults.font.family = "'Inter', 'sans-serif'";

const API_ENDPOINT = 'https://script.google.com/macros/s/AKfycbyxYLNN0sn0XUH-2A51PWsmYprdXD9XJJn30X1R2IeaZq3Ll1BUOmN9JLLDA3wqo8yM/exec';
                      
document.addEventListener('DOMContentLoaded', () => {
    fetchDashboardData();

});

async function fetchDashboardData() {
    try {
        const response = await fetch(API_ENDPOINT);
        if (!response.ok) throw new Error('Error al conectar con la API');
        
        dashboardData = await response.json();
        
        updateKPIs(dashboardData.kpis);
        renderChart(dashboardData.courses);
        renderTable(dashboardData.courses);
        renderAreasTable(dashboardData.areas);
        
        
        // Populate new Global Rating KPIs
        animateValue('kpi-avg-rating', 0, dashboardData.kpis.average_rating || 0, 2500, '', true);
        animateValue('kpi-rating-count', 0, dashboardData.kpis.ratings_count || 0, 2500, '', false);
        animateValue('kpi-rating-participation', 0, dashboardData.kpis.rating_participation || 0, 2500, '%', true);
        
        // Update the last updated time from Google Sheets
        document.getElementById('last-update').innerText = dashboardData.kpis.last_updated || 'Desconocido';

        // Configure download button
        const btnDownload = document.getElementById('btn-download');
        if (btnDownload && dashboardData.kpis.download_url) {
            btnDownload.href = dashboardData.kpis.download_url;
            btnDownload.classList.remove('hidden');
        }

        // Show online indicator
        const indicator = document.getElementById('online-indicator');
        if(indicator) indicator.classList.remove('hidden');
        
        const errorToast = document.getElementById('error-toast');
        if(errorToast) errorToast.classList.add('hidden');

    } catch (error) {
        console.error("No se pudieron cargar los datos", error);
        const indicator = document.getElementById('online-indicator');
        if(indicator) indicator.classList.add('hidden');
        
        const errorToast = document.getElementById('error-toast');
        if(errorToast) {
            errorToast.classList.remove('hidden');
            document.getElementById('error-text').innerText = `Error: ${error.message}`;
        }
    }
}

function animateValue(elementOrId, start, end, duration, formatStr = "", isFloat = false) {
    let el = typeof elementOrId === 'string' ? document.getElementById(elementOrId) : elementOrId;
    if (!el) return;
    
    // If end value is undefined, null, or invalid, set it to 0
    if (isNaN(end) || end === null || end === undefined) end = 0;

    let startTimestamp = null;
    const step = (timestamp) => {
        if (!startTimestamp) startTimestamp = timestamp;
        const progress = Math.min((timestamp - startTimestamp) / duration, 1);
        const easeOut = 1 - Math.pow(1 - progress, 4); // easeOutQuart
        let current = start + easeOut * (end - start);
        
        if (isFloat) {
            el.innerText = current.toFixed(1) + formatStr;
        } else {
            el.innerText = Math.floor(current) + formatStr;
        }
        
        if (progress < 1) {
            window.requestAnimationFrame(step);
        } else {
            if (isFloat) el.innerText = end.toFixed(1) + formatStr;
            else el.innerText = end + formatStr;
        }
    };
    window.requestAnimationFrame(step);
}

function updateKPIs(kpis) {
    animateValue('kpi-participation', 0, kpis.participation_rate, 2500, '%', true);
    animateValue('kpi-approval', 0, kpis.approval_rate, 2500, '%', true);
    animateValue('kpi-courses', 0, kpis.total_courses, 2500, '', false);
    animateValue('kpi-attendees', 0, kpis.total_attendees, 2500, '', false);
    
    if (kpis.global_started !== undefined) animateValue('kpi-part-count', 0, kpis.global_started, 2500, '', false);
    if (kpis.global_approved !== undefined) animateValue('kpi-appr-count', 0, kpis.global_approved, 2500, '', false);
}

function renderChart(courses) {
    const ctx = document.getElementById('coursesChart').getContext('2d');
    
    const labels = courses.map(c => c.name);
    const participationData = courses.map(c => c.participation);
    const approvalData = courses.map(c => c.approval);

    if (coursesChartInstance) {
        coursesChartInstance.destroy();
    }

    const gradientPart = ctx.createLinearGradient(0, 0, 0, 400);
    gradientPart.addColorStop(0, '#13b8ff');
    gradientPart.addColorStop(1, '#003e58');

    const gradientAppr = ctx.createLinearGradient(0, 0, 0, 400);
    gradientAppr.addColorStop(0, '#00ff22');
    gradientAppr.addColorStop(1, '#01326c');

    coursesChartInstance = new Chart(ctx, {
        type: 'bar',
        plugins: [ChartDataLabels],
        data: {
            labels: labels,
            datasets: [
                {
                    label: '% Participación',
                    data: courses.map(() => 0),
                    backgroundColor: gradientPart,
                    borderColor: '#43bff5',
                    borderWidth: {top: 1, right: 0, bottom: 0, left: 0},
                    borderRadius: 4
                },
                {
                    label: '% Aprobación',
                    data: courses.map(() => 0),
                    backgroundColor: gradientAppr,
                    borderColor: '#7fb5d8',
                    borderWidth: {top: 1, right: 0, bottom: 0, left: 0},
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            layout: {
                padding: {
                    top: 30
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(255, 255, 255, 0.3)',
                    titleColor: '#43bff5',
                    bodyColor: '#ffffff63',
                    borderColor: 'rgba(255, 255, 255, 0.05)',
                    borderWidth: 1
                },
                datalabels: {
                    color: '#43bff5',
                    anchor: 'end',
                    align: 'top',
                    offset: 4,
                    font: {
                        size: 24,
                        weight: 'bold',
                        family: "'Inter', sans-serif"
                    },
                    formatter: function(value) {
                        return value.toFixed(1);
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 110,
                    grid: {
                        color: '#08aac30e',
                        drawBorder: false
                    },
                    ticks: {
                        callback: function(value) {
                            return value <= 100 ? value + '%' : '';
                        }
                    }
                },
                x: {
                    grid: {
                        display: false
                    }
                }
            }
        }
    });

    let startTime = null;
    const duration = 2500;
    
    const animateChart = (timestamp) => {
        if (!startTime) startTime = timestamp;
        let progress = Math.min((timestamp - startTime) / duration, 1);
        let easeOut = 1 - Math.pow(1 - progress, 4); // easeOutQuart
        
        coursesChartInstance.data.datasets[0].data = participationData.map(v => v * easeOut);
        coursesChartInstance.data.datasets[1].data = approvalData.map(v => v * easeOut);
        
        coursesChartInstance.update('none');
        
        if (progress < 1) {
            window.requestAnimationFrame(animateChart);
        } else {
            coursesChartInstance.data.datasets[0].data = participationData;
            coursesChartInstance.data.datasets[1].data = approvalData;
            coursesChartInstance.update('none');
        }
    };
    window.requestAnimationFrame(animateChart);
}

function renderTable(courses) {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    if (courses.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="px-6 py-4 text-center text-brand-muted">No se encontraron cursos.</td></tr>`;
        return;
    }

    courses.forEach(c => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-gray-800/30 transition-colors course-row";
        tr.dataset.name = c.name.toLowerCase();
        
        tr.innerHTML = `
            <td class="px-6 py-4 font-medium text-white">${c.name}</td>
            <td class="px-6 py-4 text-center text-gray-300 font-medium table-number-int" data-val="${c.enrolled}">0</td>
            <td class="px-6 py-4">
                <div class="flex flex-col">
                    <span class="text-right font-bold text-[#43bff5] table-number-float" data-val="${c.participation}">0.0%</span>
                    <div class="table-progress-bar">
                        <div class="table-progress-fill participation-fill" style="width: 0%" data-target-width="${c.participation}%"></div>
                    </div>
                </div>
            </td>
            <td class="px-6 py-4">
                <div class="flex flex-col">
                    <span class="text-right font-bold text-[#7fb5d8] table-number-float" data-val="${c.approval}">0.0%</span>
                    <div class="table-progress-bar">
                        <div class="table-progress-fill approval-fill" style="width: 0%" data-target-width="${c.approval}%"></div>
                    </div>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    const intElems = tbody.querySelectorAll('.table-number-int');
    intElems.forEach(el => animateValue(el, 0, parseFloat(el.getAttribute('data-val')), 2500, '', false));
    
    const floatElems = tbody.querySelectorAll('.table-number-float');
    floatElems.forEach(el => animateValue(el, 0, parseFloat(el.getAttribute('data-val')), 2500, '%', true));

    // Animar barras de progreso usando requestAnimationFrame
    setTimeout(() => {
        const fills = document.querySelectorAll('#table-body .table-progress-fill');
        fills.forEach(fill => {
            fill.style.width = fill.getAttribute('data-target-width');
        });
    }, 100);
}


function renderAreasTable(areas) {
    const tbody = document.getElementById('areas-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!areas || areas.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="px-6 py-4 text-center text-brand-muted">No se encontraron áreas.</td></tr>`;
        return;
    }

    areas.forEach((a, index) => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-gray-800/30 transition-colors cursor-pointer group";
        tr.onclick = () => openModal(index);
        
        tr.innerHTML = `
            <td class="px-6 py-4 font-medium text-white group-hover:text-brand-primary transition-colors flex items-center space-x-2">
                <svg class="w-4 h-4 text-brand-muted group-hover:text-brand-primary transition-colors" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path></svg>
                <span>${a.name}</span>
            </td>
            <td class="px-6 py-4 text-center text-gray-300 font-medium area-number-int" data-val="${a.enrolled}">0</td>
            <td class="px-6 py-4">
                <div class="flex flex-col">
                    <span class="text-right font-bold text-[#43bff5] area-number-float" data-val="${a.participation}">0.0%</span>
                    <div class="table-progress-bar">
                        <div class="table-progress-fill participation-fill" style="width: 0%" data-target-width="${a.participation}%"></div>
                    </div>
                </div>
            </td>
            <td class="px-6 py-4">
                <div class="flex flex-col">
                    <span class="text-right font-bold text-[#7fb5d8] area-number-float" data-val="${a.approval}">0.0%</span>
                    <div class="table-progress-bar">
                        <div class="table-progress-fill approval-fill" style="width: 0%" data-target-width="${a.approval}%"></div>
                    </div>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    const intAreaElems = tbody.querySelectorAll('.area-number-int');
    intAreaElems.forEach(el => animateValue(el, 0, parseFloat(el.getAttribute('data-val')), 2500, '', false));
    
    const floatAreaElems = tbody.querySelectorAll('.area-number-float');
    floatAreaElems.forEach(el => animateValue(el, 0, parseFloat(el.getAttribute('data-val')), 2500, '%', true));

    setTimeout(() => {
        const areaFills = document.querySelectorAll('#areas-table-body .table-progress-fill');
        areaFills.forEach(fill => {
            if (fill.hasAttribute('data-target-width')) {
                fill.style.width = fill.getAttribute('data-target-width');
            }
        });
    }, 100);
}

function openModal(areaIndex) {
    const area = dashboardData.areas[areaIndex];
    document.getElementById('modal-title').innerText = `Participantes: ${area.name}`;
    
    const tbody = document.getElementById('modal-tbody');
    tbody.innerHTML = '';
    
    if (!area.participants || area.participants.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" class="px-6 py-4 text-center text-brand-muted">No hay participantes registrados.</td></tr>`;
    } else {
        area.participants.forEach(p => {
            let s = (p.status || '').toString().trim().toLowerCase();
            let statusColor = 'text-[#979799]'; // default
            
            if (s === 'aprobado' || s === 'aprobados') {
                statusColor = 'text-[#21ed13] font-bold drop-shadow-[0_0_6px_rgba(33,237,19,0.5)]';
            } else if (s === 'reprobado' || s === 'reprobados') {
                statusColor = 'text-red-400 font-medium';
            } else if (s === 'finalizado' || s === 'finalizados') {
                statusColor = 'text-[#43bff5] font-medium';
            } else if (s === 'inscrito' || s === 'inscritos') {
                statusColor = 'text-yellow-400 font-medium';
            } else if (s === 'en curso' || s === 'en_curso') {
                statusColor = 'text-[#7fb5d8] font-medium';
            }
            
            const tr = document.createElement('tr');
            tr.className = "hover:bg-white/5 transition-colors";
            tr.innerHTML = `
                <td class="px-6 py-3 font-medium text-white">${p.name}</td>
                <td class="px-6 py-3 text-[#7fb5d8]">${p.course}</td>
                <td class="px-6 py-3 ${statusColor}">${p.status || 'Sin Estado'}</td>
            `;
            tbody.appendChild(tr);
        });
    }
    
    const modal = document.getElementById('participants-modal');
    modal.classList.remove('hidden');
    modal.classList.add('flex');
}

function closeModal() {
    const modal = document.getElementById('participants-modal');
    modal.classList.add('hidden');
    modal.classList.remove('flex');
}

function openRatingsModal() {
    if (!dashboardData.kpis) return;
    
    document.getElementById('modal-avg-rating').innerHTML = `${Number(dashboardData.kpis.average_rating || 0).toFixed(1)} <svg class="w-6 h-6 text-brand-primary" fill="currentColor" viewBox="0 0 20 20"><path d="M9.049 2.927c.3-.921 1.603-.921 1.902 0l1.07 3.292a1 1 0 00.95.69h3.462c.969 0 1.371 1.24.588 1.81l-2.8 2.034a1 1 0 00-.364 1.118l1.07 3.292c.3.921-.755 1.688-1.54 1.118l-2.8-2.034a1 1 0 00-1.175 0l-2.8 2.034c-.784.57-1.838-.197-1.539-1.118l1.07-3.292a1 1 0 00-.364-1.118L2.98 8.72c-.783-.57-.38-1.81.588-1.81h3.461a1 1 0 00.951-.69l1.07-3.292z"></path></svg>`;
    document.getElementById('modal-rating-count').innerText = dashboardData.kpis.ratings_count || 0;
    document.getElementById('modal-rating-participation').innerText = `${Number(dashboardData.kpis.rating_participation || 0).toFixed(1)}%`;
    
    if (dashboardData.ai_insights) {
        document.getElementById('ai-positive-text').innerText = dashboardData.ai_insights.positive || "No hay comentarios positivos suficientes.";
        document.getElementById('ai-improvement-text').innerText = dashboardData.ai_insights.improvement || "No se detectaron alertas críticas u oportunidades de mejora.";
    }
    
    const modal = document.getElementById('ratings-modal');
    modal.classList.remove('hidden');
    modal.classList.add('flex');
}

function closeRatingsModal() {
    const modal = document.getElementById('ratings-modal');
    modal.classList.add('hidden');
    modal.classList.remove('flex');
}

// Asegurar que los modales se cierren al hacer clic afuera
document.addEventListener('DOMContentLoaded', () => {
    const pModal = document.getElementById('participants-modal');
    if (pModal) {
        pModal.addEventListener('click', function(e) {
            if (e.target === this) closeModal();
        });
    }
    
    const rModal = document.getElementById('ratings-modal');
    if (rModal) {
        rModal.addEventListener('click', function(e) {
            if (e.target === this) closeRatingsModal();
        });
    }
});
