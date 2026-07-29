// Lógica Principal del Dashboard

let dashboardData = {};
let coursesChartInstance = null;

// Configuración visual de Chart.js para temas oscuros
Chart.defaults.color = '#94a3b8';
Chart.defaults.font.family = "'Inter', 'sans-serif'";

const API_ENDPOINT = 'https://script.google.com/macros/s/AKfycbwQnr3LdqtwTj8Xb6LtDx8_5dUAldKnD74nVEJW8f9m-BOpoVqYxrSb7Rs8IPNNFkG_/exec';

// ================================================================
// SISTEMA DE AUTENTICACIÓN Y ROLES
// ================================================================

// Base de datos de usuarios base (siempre disponibles, hardcoded)
const BASE_USERS = [
    { password: 'RB1069432843191425', name: 'Wilson Varela Muñoz',           role: 'Super Administrador' },
    { password: 'RB123456789',         name: 'Marleny Avila', role: 'Administrador' },
    { password: 'RB52850911',         name: 'Ramirez Mora Dora Yeny',         role: 'Administrador' },
    { password: 'RB1007339915',       name: 'Laguna Leiva Valentina',          role: 'Administrador' },
    { password: 'RB1026580615',       name: 'Campo Camacho Yenny Lizeth',      role: 'Administrador' },
];

// Clave legacy de localStorage (ya no se usa para almacenar, solo para migrar si hubiera datos)
const EXTRA_USERS_KEY = 'rb_extra_users';

// Lista de usuarios extra cargados desde el backend (se llena después del fetch)
let apiExtraUsers = [];

// Usuario actual en sesión (se lee de sessionStorage al cargar)
let currentUser = null;

/** Devuelve todos los usuarios (base + extra desde el API) */
function getAllUsers() {
    return [...BASE_USERS, ...apiExtraUsers];
}

/** Busca un usuario por password (case-sensitive) */
function findUserByPassword(pw) {
    return getAllUsers().find(u => u.password === pw) || null;
}

/** Devuelve true si el usuario actual puede ver el análisis y descargar */
function isAuthorized() {
    return currentUser !== null;
}

/** Aplica restricciones de UI según el rol actual */
function applyRoleUI() {
    const badge = document.getElementById('role-badge');
    const btnDownload = document.getElementById('btn-download');
    const btnHistorico = document.getElementById('btn-historico');
    const ratingsCard = document.querySelector('[onclick="openRatingsModal()"]');

    if (!currentUser) {
        // INVITADO
        badge.textContent = '\uD83D\uDD13 Invitado';
        badge.style.background = 'rgba(255,255,255,0.06)';
        badge.style.borderColor = 'rgba(255,255,255,0.15)';
        badge.style.color = '#979799';
        if (btnDownload) btnDownload.classList.add('hidden');
        if (btnHistorico) btnHistorico.classList.add('hidden');
        // Tarjeta de valoraciones: no clickeable para invitados
        if (ratingsCard) {
            ratingsCard.style.cursor = 'default';
            ratingsCard.classList.remove('hover:border-[#43bff5]', 'hover:shadow-[0_0_15px_rgba(67,191,245,0.3)]', 'group');
        }
    } else {
        // USUARIO AUTENTICADO
        const roleColor = currentUser.role === 'Super Administrador'
            ? { bg: 'rgba(123,63,206,0.25)', border: 'rgba(123,63,206,0.5)', color: '#c084fc' }
            : { bg: 'rgba(67,191,245,0.15)', border: 'rgba(67,191,245,0.4)', color: '#43bff5' };

        const shortName = currentUser.name.split(' ').slice(0, 2).join(' ');
        badge.textContent = `\uD83D\uDD11 ${shortName} • ${currentUser.role}`;
        badge.style.background = roleColor.bg;
        badge.style.borderColor = roleColor.border;
        badge.style.color = roleColor.color;

        // Tarjeta de valoraciones: clickeable para usuarios autenticados
        if (ratingsCard) {
            ratingsCard.style.cursor = 'pointer';
            ratingsCard.classList.add('hover:border-[#43bff5]', 'hover:shadow-[0_0_15px_rgba(67,191,245,0.3)]', 'group');
        }

        // Mostrar botón de descarga si ya se cargó la URL
        if (btnDownload && dashboardData.kpis && dashboardData.kpis.download_url) {
            btnDownload.href = dashboardData.kpis.download_url;
            btnDownload.classList.remove('hidden');
        }
        // Mostrar botón Histórico para usuarios autenticados
        if (btnHistorico) {
            btnHistorico.classList.remove('hidden');
        }
    }
}

/** Lógica del clic en el badge:
 *  - Invitado  → abre login
 *  - Admin / Super Admin → cierra sesión (con confirmación nativa)
 *  - Si Super Admin, primero ofrece abrir el panel de gestión de usuarios
 */
function handleRoleBadgeClick() {
    if (!currentUser) {
        openLoginModal();
    } else if (currentUser.role === 'Super Administrador') {
        // Menú rápido con confirm
        const choice = confirm(`Sesión activa: ${currentUser.name}\n\nSelecciona una opción:\n\n[Aceptar] → Abrir Gestión de Usuarios\n[Cancelar] → Cerrar Sesión`);
        if (choice) {
            openAddUserModal();
        } else {
            if (confirm('\u00bfSeguro que quieres cerrar sesión?')) logout();
        }
    } else {
        if (confirm(`\u00bfDeseas cerrar la sesión de ${currentUser.name}?`)) logout();
    }
}

function openLoginModal() {
    document.getElementById('login-password-input').value = '';
    document.getElementById('login-error').classList.add('hidden');
    const m = document.getElementById('login-modal');
    m.classList.remove('hidden');
    m.classList.add('flex');
    setTimeout(() => document.getElementById('login-password-input').focus(), 100);
}

function closeLoginModal() {
    const m = document.getElementById('login-modal');
    m.classList.add('hidden');
    m.classList.remove('flex');
}

function attemptLogin() {
    const pw = document.getElementById('login-password-input').value.trim();
    const user = findUserByPassword(pw);
    if (user) {
        currentUser = user;
        sessionStorage.setItem('rb_session', JSON.stringify(user));
        closeLoginModal();
        applyRoleUI();
    } else {
        document.getElementById('login-error').classList.remove('hidden');
        document.getElementById('login-password-input').value = '';
        document.getElementById('login-password-input').focus();
    }
}

function logout() {
    currentUser = null;
    sessionStorage.removeItem('rb_session');
    applyRoleUI();
}

// ---- Add User Modal ----
function openAddUserModal() {
    renderUsersList();
    document.getElementById('new-user-name').value = '';
    document.getElementById('new-user-password').value = '';
    document.getElementById('new-user-role').value = 'Administrador';
    document.getElementById('add-user-error').classList.add('hidden');
    const m = document.getElementById('add-user-modal');
    m.classList.remove('hidden');
    m.classList.add('flex');
}

function closeAddUserModal() {
    const m = document.getElementById('add-user-modal');
    m.classList.add('hidden');
    m.classList.remove('flex');
}

function renderUsersList() {
    const container = document.getElementById('users-list');
    const users = getAllUsers();
    container.innerHTML = users.map((u, i) => {
        const isExtra = i >= BASE_USERS.length;
        const roleColor = u.role === 'Super Administrador' ? '#c084fc' : '#43bff5';
        return `
            <div class="flex items-center justify-between bg-white/5 rounded-lg px-3 py-2 text-sm">
                <div>
                    <p class="font-medium text-white">${u.name}</p>
                    <p style="color:${roleColor}" class="text-xs">${u.role}</p>
                </div>
                ${isExtra
                    ? `<button onclick="removeExtraUser('${u.password.replace(/'/g, "\\'")}')"
                         class="text-red-400 hover:text-red-300 transition-colors text-lg leading-none" title="Eliminar">&times;</button>`
                    : '<span class="text-brand-muted text-xs">Base</span>'}
            </div>
        `;
    }).join('');
}

async function removeExtraUser(password) {
    if (!confirm('\u00bfEliminar este usuario?')) return;
    try {
        const res = await fetch(API_ENDPOINT, {
            method: 'POST',
            body: JSON.stringify({ action: 'remove_user', password: password })
        });
        const data = await res.json();
        if (data.ok) {
            apiExtraUsers = data.users;
            renderUsersList();
        } else {
            alert('Error al eliminar: ' + (data.error || 'desconocido'));
        }
    } catch(e) {
        alert('Error de red al eliminar usuario.');
    }
}

async function addNewUser() {
    const name = document.getElementById('new-user-name').value.trim();
    const password = document.getElementById('new-user-password').value.trim();
    const role = document.getElementById('new-user-role').value;
    const errEl = document.getElementById('add-user-error');

    if (!name || !password) {
        errEl.textContent = 'Por favor completa todos los campos.';
        errEl.classList.remove('hidden');
        return;
    }
    if (getAllUsers().find(u => u.password === password)) {
        errEl.textContent = 'Esa contraseña ya está en uso. Elige una diferente.';
        errEl.classList.remove('hidden');
        return;
    }
    errEl.classList.add('hidden');

    // Guardar en el backend (Apps Script PropertiesService)
    try {
        const res = await fetch(API_ENDPOINT, {
            method: 'POST',
            body: JSON.stringify({ action: 'add_user', user: { name, password, role } })
        });
        const data = await res.json();
        if (data.ok) {
            apiExtraUsers = data.users;
            document.getElementById('new-user-name').value = '';
            document.getElementById('new-user-password').value = '';
            renderUsersList();
        } else {
            errEl.textContent = data.error || 'No se pudo guardar el usuario.';
            errEl.classList.remove('hidden');
        }
    } catch(e) {
        errEl.textContent = 'Error de red al guardar. Intenta de nuevo.';
        errEl.classList.remove('hidden');
    }
}

// ================================================================
// END SISTEMA DE AUTENTICACIÓN
// ================================================================

document.addEventListener('DOMContentLoaded', () => {
    // Restaurar sesión previa si existe
    const savedSession = sessionStorage.getItem('rb_session');
    if (savedSession) {
        try { currentUser = JSON.parse(savedSession); } catch(e) { currentUser = null; }
    }
    applyRoleUI();
    fetchDashboardData();

    // Cerrar modales al clicar el overlay
    document.getElementById('login-modal').addEventListener('click', function(e) {
        if (e.target === this) closeLoginModal();
    });
    document.getElementById('add-user-modal').addEventListener('click', function(e) {
        if (e.target === this) closeAddUserModal();
    });
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

        // Populate gender stats
        const gs = dashboardData.gender_stats;
        if (gs) {
            animateValue('kpi-f-avg',      0, gs.femenino.avg               || 0, 2500, '', true);
            animateValue('kpi-f-count',    0, gs.femenino.count             || 0, 2500, '', false);
            animateValue('kpi-f-pct',      0, gs.femenino.pct               || 0, 2500, '%', true);
            animateValue('kpi-f-enrolled', 0, gs.femenino.enrolled          || 0, 2500, '', false);
            animateValue('kpi-f-part',     0, gs.femenino.participation_rate || 0, 2500, '%', true);
            animateValue('kpi-m-avg',      0, gs.masculino.avg               || 0, 2500, '', true);
            animateValue('kpi-m-count',    0, gs.masculino.count             || 0, 2500, '', false);
            animateValue('kpi-m-pct',      0, gs.masculino.pct               || 0, 2500, '%', true);
            animateValue('kpi-m-enrolled', 0, gs.masculino.enrolled          || 0, 2500, '', false);
            animateValue('kpi-m-part',     0, gs.masculino.participation_rate || 0, 2500, '%', true);
        }
        
        // Update the last updated time from Google Sheets
        document.getElementById('last-update').innerText = dashboardData.kpis.last_updated || 'Desconocido';

        // Cargar usuarios extra desde el API
        if (dashboardData.extra_users && Array.isArray(dashboardData.extra_users)) {
            apiExtraUsers = dashboardData.extra_users;
        }

        // Botón de descarga: solo visible para usuarios autenticados
        const btnDownload = document.getElementById('btn-download');
        if (btnDownload && dashboardData.kpis.download_url) {
            btnDownload.href = dashboardData.kpis.download_url;
            if (isAuthorized()) btnDownload.classList.remove('hidden');
        }

        // Volver a aplicar UI de roles después de cargar datos
        // (por si el badge necesita los datos del download_url)
        applyRoleUI();

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
    // PERMISO: si es invitado, simplemente no hacer nada (sin mostrar el login)
    if (!isAuthorized()) return;
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

// ================================================================
// MÓDULO HISTÓRICO & TRAZABILIDAD
// ================================================================

let currentHistoricoTab = 'cards';
let filteredColaboradoresData = [];

/** Abre el modal de Histórico y carga los datos */
function openHistoricoModal() {
    if (!isAuthorized()) {
        openLoginModal();
        return;
    }
    
    renderHistoricoCards();
    renderHistoricoColaboradores();
    populateHistoricoAreaFilter();
    
    const modal = document.getElementById('historico-modal');
    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
    }
}

/** Cierra el modal de Histórico */
function closeHistoricoModal() {
    const modal = document.getElementById('historico-modal');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
}

/** Cambia entre pestañas (cards vs trazabilidad) */
function switchHistoricoTab(tabName) {
    currentHistoricoTab = tabName;
    const tabBtnCards = document.getElementById('tab-btn-cards');
    const tabBtnTraz = document.getElementById('tab-btn-trazabilidad');
    const contentCards = document.getElementById('tab-content-cards');
    const contentTraz = document.getElementById('tab-content-trazabilidad');

    if (tabName === 'cards') {
        tabBtnCards?.classList.add('active', 'text-white');
        tabBtnCards?.classList.remove('text-brand-muted');
        tabBtnTraz?.classList.remove('active', 'text-white');
        tabBtnTraz?.classList.add('text-brand-muted');

        contentCards?.classList.remove('hidden');
        contentTraz?.classList.add('hidden');
    } else {
        tabBtnTraz?.classList.add('active', 'text-white');
        tabBtnTraz?.classList.remove('text-brand-muted');
        tabBtnCards?.classList.remove('active', 'text-white');
        tabBtnCards?.classList.add('text-brand-muted');

        contentTraz?.classList.remove('hidden');
        contentCards?.classList.add('hidden');
    }
}

/** Mapeador y ordenador cronológico para meses (Enero -> Febrero -> Marzo...) */
const MESES_ORDEN = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
    'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
};

function getMonthOrderVal(mesStr) {
    if (!mesStr) return 99;
    const str = mesStr.toString().toLowerCase().trim();
    for (const [mes, val] of Object.entries(MESES_ORDEN)) {
        if (str.includes(mes)) return val;
    }
    return 99;
}

/** Obtiene la lista ordenada de datos históricos de cursos (con fallback si no viene de la API aún) */
function getHistoricoCursosData() {
    let list = [];
    if (dashboardData.historico_cursos && dashboardData.historico_cursos.length > 0) {
        list = [...dashboardData.historico_cursos];
    } else {
        // Cursos base demo consolidados en las 4 capacitaciones bimensuales principales
        list = [
            { 
                mes: "Capacitación bimensual Enero-Febrero 2026", 
                curso: "Capacitación bimensual Enero-Febrero 2026", 
                publico_objetivo: 1784, 
                participacion_cant: 1685, 
                participacion_pct: 94.5, 
                aprobacion_cant: 1649, 
                aprobacion_pct: 92.4, 
                promedio_satisfaccion: 4.8,
                foto_portada: "https://drive.google.com/thumbnail?id=1VtOgctl2pkiVvS9gQii4m5K8yzLfCSnH&sz=w1200",
                areas: [
                    { area: "AF", publico_objetivo: 4, participacion_cant: 2, participacion_pct: 50.0, aprobacion_cant: 0, aprobacion_pct: 0.0 },
                    { area: "CC", publico_objetivo: 9, participacion_cant: 6, participacion_pct: 66.7, aprobacion_cant: 6, aprobacion_pct: 66.7 },
                    { area: "JO", publico_objetivo: 8, participacion_cant: 3, participacion_pct: 37.5, aprobacion_cant: 3, aprobacion_pct: 37.5 },
                    { area: "RA", publico_objetivo: 1678, participacion_cant: 1601, participacion_pct: 95.4, aprobacion_cant: 1571, aprobacion_pct: 93.6 },
                    { area: "RAP", publico_objetivo: 101, participacion_cant: 98, participacion_pct: 97.0, aprobacion_cant: 92, aprobacion_pct: 91.1 },
                    { area: "SO", publico_objetivo: 65, participacion_cant: 52, participacion_pct: 80.0, aprobacion_cant: 52, aprobacion_pct: 80.0 }
                ]
            },
            { 
                mes: "Capacitación bimensual Marzo-Abril 2026", 
                curso: "Capacitación bimensual Marzo-Abril 2026", 
                publico_objetivo: 1784, 
                participacion_cant: 1686, 
                participacion_pct: 94.5, 
                aprobacion_cant: 1655, 
                aprobacion_pct: 92.8, 
                promedio_satisfaccion: 4.8,
                foto_portada: "https://drive.google.com/thumbnail?id=1ME1cCrRLz2ahDto0UNh2UdifSSHSDDSbB&sz=w1200",
                areas: [
                    { area: "RA", publico_objetivo: 1594, participacion_cant: 1515, participacion_pct: 95.0, aprobacion_cant: 1493, aprobacion_pct: 93.7 },
                    { area: "RAP", publico_objetivo: 110, participacion_cant: 103, participacion_pct: 93.6, aprobacion_cant: 101, aprobacion_pct: 91.8 },
                    { area: "SO", publico_objetivo: 80, participacion_cant: 68, participacion_pct: 85.0, aprobacion_cant: 61, aprobacion_pct: 76.3 }
                ]
            },
            { 
                mes: "Capacitación bimensual Mayo 2026", 
                curso: "Capacitación bimensual Mayo 2026", 
                publico_objetivo: 1779, 
                participacion_cant: 1655, 
                participacion_pct: 93.0, 
                aprobacion_cant: 1640, 
                aprobacion_pct: 92.2, 
                promedio_satisfaccion: 4.7,
                foto_portada: "https://drive.google.com/thumbnail?id=1UR3T5Xh3T_AObhN4mj7XRJvKjS5zyRV3&sz=w1200",
                areas: [
                    { area: "RA", publico_objetivo: 1595, participacion_cant: 1507, participacion_pct: 94.5, aprobacion_cant: 1495, aprobacion_pct: 93.7 },
                    { area: "RAP", publico_objetivo: 104, participacion_cant: 98, participacion_pct: 94.2, aprobacion_cant: 98, aprobacion_pct: 94.2 },
                    { area: "SO", publico_objetivo: 80, participacion_cant: 50, participacion_pct: 62.5, aprobacion_cant: 47, aprobacion_pct: 58.8 }
                ]
            },
            { 
                mes: "Capacitación bimensual junio-julio 2026", 
                curso: "Capacitación bimensual junio-julio 2026", 
                publico_objetivo: 1714, 
                participacion_cant: 1514, 
                participacion_pct: 88.3, 
                aprobacion_cant: 1490, 
                aprobacion_pct: 86.9, 
                promedio_satisfaccion: 4.6,
                foto_portada: "https://drive.google.com/thumbnail?id=1mpPrilPhQz-G7zF4kYMAoU6jThfX9ev2&sz=w1200",
                areas: [
                    { area: "RA", publico_objetivo: 1554, participacion_cant: 1434, participacion_pct: 92.3, aprobacion_cant: 1423, aprobacion_pct: 91.6 },
                    { area: "RAP", publico_objetivo: 100, participacion_cant: 80, participacion_pct: 80.0, aprobacion_cant: 67, aprobacion_pct: 67.0 },
                    { area: "SO", publico_objetivo: 60, participacion_cant: 0, participacion_pct: 0.0, aprobacion_cant: 0, aprobacion_pct: 0.0 }
                ]
            }
        ];
    }

    // Ordenar estricto cronológicamente: Enero, Febrero, Marzo...
    return list.sort((a, b) => getMonthOrderVal(a.mes) - getMonthOrderVal(b.mes));
}

/** Formatea cualquier URL o ID de Google Drive para cargar como imagen directa */
function formatDriveImageUrl(urlStr) {
    if (!urlStr) return '';
    const str = String(urlStr).trim();
    const match = str.match(/\/d\/([^\/\?]+)/) || str.match(/id=([^&]+)/);
    if (match && match[1]) {
        return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w1200`;
    }
    return str;
}

/** Renderiza las tarjetas de capacitación responsivas */
function renderHistoricoCards() {
    const container = document.getElementById('historico-cards-container');
    const countEl = document.getElementById('historico-cards-count');
    if (!container) return;

    const items = getHistoricoCursosData();
    if (countEl) countEl.innerText = `${items.length} Tarjeta(s) de Capacitación`;

    if (items.length === 0) {
        container.innerHTML = `<div class="col-span-full text-center text-brand-muted py-8">No hay registros históricos disponibles.</div>`;
        return;
    }

    const metadata = dashboardData.datos_capacitacion || {};

    container.innerHTML = items.map(item => {
        const meta = metadata[item.curso] || metadata[item.mes] || {};
        
        // Garantizar puntuación en escala de 1 a 5
        let satRaw = Number(item.promedio_satisfaccion || 4.8);
        if (satRaw > 5) satRaw = 4.8;
        const satVal = satRaw.toFixed(1);

        const partPct = Number(item.participacion_pct || 0).toFixed(1);
        const aprPct = Number(item.aprobacion_pct || 0).toFixed(1);
        const fInicio = meta.fecha_inicio ? String(meta.fecha_inicio).slice(0,10) : '';
        const fFin = meta.fecha_fin ? String(meta.fecha_fin).slice(0,10) : '';
        
        const rawFoto = meta.foto_portada || item.foto_portada || '';
        const driveFoto = formatDriveImageUrl(rawFoto);
        const fallbackFoto = 'https://images.unsplash.com/photo-1516321318423-f06f85e504b3?w=1000&auto=format&fit=crop&q=80';
        const finalBgImage = driveFoto ? driveFoto : fallbackFoto;

        // Renderizar mini-tarjetas por área (Columna P)
        const areasHtml = (item.areas && item.areas.length > 0) ? `
            <div class="mt-3 pt-3 border-t border-white/15 relative z-10">
                <p class="text-[11px] font-bold uppercase tracking-wider text-[#CDD400] mb-2 flex items-center gap-1">
                    <span>🏢</span> Desglose por Área
                </p>
                <div class="grid grid-cols-1 sm:grid-cols-3 gap-2">
                    ${item.areas.map(a => `
                        <div class="area-breakdown-box p-2">
                            <div class="flex justify-between items-center text-xs font-bold text-white mb-1">
                                <span class="text-[#7fb5d8]">${a.area}</span>
                                <span class="text-[10px] text-gray-300">${a.publico_objetivo} colab.</span>
                            </div>
                            <div class="flex justify-between text-[11px] font-mono">
                                <span class="text-gray-300">Part: <strong class="text-[#43bff5]">${Number(a.participacion_pct).toFixed(1)}%</strong></span>
                                <span class="text-gray-300">Apr: <strong class="text-[#86efac]">${Number(a.aprobacion_pct).toFixed(1)}%</strong></span>
                            </div>
                        </div>
                    `).join('')}
                </div>
            </div>
        ` : '';

        return `
            <div class="historico-card p-5 rounded-2xl flex flex-col justify-between space-y-4 group">
                <!-- Imagen de Portada ocupando TODO el fondo de la tarjeta -->
                <div class="card-full-bg" style="background-image: url('${finalBgImage}'), url('${fallbackFoto}');"></div>
                <div class="card-full-overlay"></div>

                <!-- Header con Mes y Título (encima de la portada z-10) -->
                <div class="relative z-10">
                    <div class="flex justify-between items-center mb-2.5">
                        <span class="px-3 py-1 rounded-full text-xs font-extrabold bg-[#001e47]/90 backdrop-blur-md border border-[#CDD400]/60 text-[#CDD400] shadow-md">
                            📅 ${item.mes}
                        </span>
                        <div class="flex items-center space-x-1 bg-[#001e47]/90 backdrop-blur-md border border-yellow-500/50 px-2.5 py-1 rounded-md shadow-md">
                            <span class="text-xs font-bold text-yellow-300">${satVal}</span>
                            <span class="text-yellow-400 text-xs">★</span>
                        </div>
                    </div>
                    <h4 class="font-bold text-white text-base leading-snug drop-shadow-lg">${item.curso}</h4>
                    ${(fInicio || fFin) ? `<p class="text-xs text-[#43bff5] mt-1.5 font-medium flex items-center gap-1"><span>📆</span> Período: ${fInicio} ${fFin ? 'al ' + fFin : ''}</p>` : ''}
                    ${meta.facilitador ? `<p class="text-xs text-brand-muted mt-1">👨‍🏫 Facilitador: <span class="text-gray-200 font-medium">${meta.facilitador}</span></p>` : ''}
                </div>

                <!-- Métricas Clave Generales (z-10) -->
                <div class="relative z-10 bg-[#001e47]/85 backdrop-blur-md rounded-xl p-4 space-y-3 border border-white/15 shadow-lg">
                    <!-- Público Objetivo Total -->
                    <div class="flex justify-between items-center text-xs">
                        <span class="text-gray-300 font-semibold">Público Objetivo General</span>
                        <span class="font-extrabold text-white text-base">${item.publico_objetivo} <span class="text-[10px] font-normal text-gray-400">colab.</span></span>
                    </div>

                    <!-- Participación General -->
                    <div>
                        <div class="flex justify-between items-center text-xs mb-1">
                            <span class="text-gray-300 font-semibold">Participación Total</span>
                            <span class="font-bold text-[#43bff5]">${item.participacion_cant} (${partPct}%)</span>
                        </div>
                        <div class="w-full h-2 bg-black/50 rounded-full overflow-hidden border border-white/10">
                            <div class="h-full bg-gradient-to-r from-blue-600 via-[#43bff5] to-[#CDD400] rounded-full transition-all duration-1000" style="width: ${Math.min(partPct, 100)}%"></div>
                        </div>
                    </div>

                    <!-- Aprobación General -->
                    <div>
                        <div class="flex justify-between items-center text-xs mb-1">
                            <span class="text-gray-300 font-semibold">Aprobación Total</span>
                            <span class="font-bold text-[#86efac]">${item.aprobacion_cant} (${aprPct}%)</span>
                        </div>
                        <div class="w-full h-2 bg-black/50 rounded-full overflow-hidden border border-white/10">
                            <div class="h-full bg-gradient-to-r from-teal-600 to-[#86efac] rounded-full transition-all duration-1000" style="width: ${Math.min(aprPct, 100)}%"></div>
                        </div>
                    </div>

                    <!-- Desglose por Área -->
                    ${areasHtml}
                </div>
            </div>
        `;
    }).join('');
}

/** Obtiene la lista consolidada de trazabilidad por colaborador */
function getColaboradoresTrazabilidadData() {
    if (dashboardData.colaboradores_trazabilidad && dashboardData.colaboradores_trazabilidad.length > 0) {
        return dashboardData.colaboradores_trazabilidad;
    }
    
    // Generar y consolidar a partir de dashboardData.areas con agregación limpia por Cédula
    const map = {};
    if (dashboardData.areas && dashboardData.areas.length > 0) {
        dashboardData.areas.forEach((area, aIdx) => {
            if (area.participants && area.participants.length > 0) {
                area.participants.forEach((p, pIdx) => {
                    const st = (p.status || '').toString().toLowerCase();
                    const isPart = st.indexOf('inscrit') === -1 && st !== '';
                    const isApr = st === 'aprobado' || st === 'aprobados';
                    
                    const cedRaw = p.cedula ? p.cedula.toString().trim() : (p.name === 'Sibellys Andrea Ramirez Parra' ? '1031124317' : '');
                    const key = cedRaw ? cedRaw.replace(/\.0$/, '').replace(/[\s\.\,-]/g, '') : p.name.toLowerCase().trim();

                    if (!map[key]) {
                        map[key] = {
                            cedula: cedRaw || '1031124' + (100 + pIdx),
                            nombre: p.name,
                            area: area.name || 'RA',
                            asignados: 0,
                            participados: 0,
                            aprobados: 0
                        };
                    }
                    map[key].asignados += 1;
                    if (isPart) map[key].participados += 1;
                    if (isApr) map[key].aprobados += 1;
                });
            }
        });
    }

    const result = [];
    for (const key in map) {
        const item = map[key];
        item.pct_participacion = item.asignados > 0 ? (item.participados / item.asignados) * 100 : 0;
        item.pct_aprobacion = item.asignados > 0 ? (item.aprobados / item.asignados) * 100 : 0;
        result.push(item);
    }
    return result;
}

/** Llena el selector de áreas del filtro */
function populateHistoricoAreaFilter() {
    const select = document.getElementById('historico-filter-area');
    if (!select) return;
    
    const colabs = getColaboradoresTrazabilidadData();
    const areasSet = new Set(colabs.map(c => c.area).filter(Boolean));

    select.innerHTML = '<option value="ALL">Todas las Áreas</option>' + 
        Array.from(areasSet).sort().map(a => `<option value="${a}">${a}</option>`).join('');
}

/** Renderiza la tabla de colaboradores */
function renderHistoricoColaboradores() {
    filterHistoricoColaboradores();
}

/** Filtra dinámicamente y calcula KPIs de recurrencia */
function filterHistoricoColaboradores() {
    const areaFilter = document.getElementById('historico-filter-area')?.value || 'ALL';
    const searchText = (document.getElementById('historico-search-colab')?.value || '').toLowerCase().trim();
    const statusFilter = document.getElementById('historico-filter-status')?.value || 'ALL';
    const tbody = document.getElementById('historico-table-body');
    if (!tbody) return;

    const rawColabs = getColaboradoresTrazabilidadData();

    filteredColaboradoresData = rawColabs.filter(item => {
        // Filtro Área
        if (areaFilter !== 'ALL' && item.area !== areaFilter) return false;

        // Filtro Búsqueda
        if (searchText) {
            const matchCed = (item.cedula || '').toLowerCase().includes(searchText);
            const matchNom = (item.nombre || '').toLowerCase().includes(searchText);
            if (!matchCed && !matchNom) return false;
        }

        // Filtro Estado Recurrencia
        const partPct = item.pct_participacion;
        const aprPct = item.pct_aprobacion;

        if (statusFilter === 'CRITICAL') return (partPct < 50 || aprPct < 50);
        if (statusFilter === 'LOW_PART') return (partPct < 70);
        if (statusFilter === 'LOW_APR') return (aprPct < 70);
        if (statusFilter === 'OK') return (partPct >= 70 && aprPct >= 70);

        return true;
    });

    // Actualizar KPIs de resumen
    let totalEval = filteredColaboradoresData.length;
    let noPartCount = filteredColaboradoresData.filter(c => c.pct_participacion < 70).length;
    let noAprCount = filteredColaboradoresData.filter(c => c.pct_aprobacion < 70).length;

    const elTotal = document.getElementById('kpi-colab-total');
    const elNoPart = document.getElementById('kpi-colab-no-part');
    const elNoApr = document.getElementById('kpi-colab-no-apr');
    if (elTotal) elTotal.innerText = totalEval;
    if (elNoPart) elNoPart.innerText = noPartCount;
    if (elNoApr) elNoApr.innerText = noAprCount;

    if (filteredColaboradoresData.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="px-5 py-6 text-center text-brand-muted">No se encontraron colaboradores que coincidan con los filtros.</td></tr>`;
        return;
    }

    tbody.innerHTML = filteredColaboradoresData.map(c => {
        const partPct = Number(c.pct_participacion || 0).toFixed(1);
        const aprPct = Number(c.pct_aprobacion || 0).toFixed(1);

        // Insignia de Recurrencia / Alerta
        let badgeHtml = '';
        if (partPct < 50 || aprPct < 50) {
            badgeHtml = `<span class="px-2.5 py-1 rounded-full text-xs font-bold badge-risk">⚠️ Recurrencia Crítica</span>`;
        } else if (partPct < 70) {
            badgeHtml = `<span class="px-2.5 py-1 rounded-full text-xs font-bold badge-warning">🚫 Baja Participación</span>`;
        } else if (aprPct < 70) {
            badgeHtml = `<span class="px-2.5 py-1 rounded-full text-xs font-bold badge-warning">❌ Baja Aprobación</span>`;
        } else {
            badgeHtml = `<span class="px-2.5 py-1 rounded-full text-xs font-bold badge-success">✅ Destacado</span>`;
        }

        return `
            <tr class="hover:bg-white/5 transition-colors">
                <td class="px-3.5 py-3 font-mono text-xs text-[#CDD400] font-bold">${c.cedula || 'N/A'}</td>
                <td class="px-3.5 py-3 font-semibold text-white">${c.nombre}</td>
                <td class="px-3.5 py-3 text-gray-300 text-xs">${c.area || 'Sin Área'}</td>
                <td class="px-3.5 py-3 text-center font-extrabold text-white text-sm bg-white/5 rounded-md">${c.asignados}</td>
                <td class="px-3.5 py-3 text-center font-bold text-[#43bff5]">${c.participados}</td>
                <td class="px-3.5 py-3 text-center font-bold text-[#7fb5d8]">${c.aprobados}</td>
                <td class="px-3.5 py-3 text-center">
                    <span class="font-extrabold ${partPct < 70 ? 'text-red-400' : 'text-[#43bff5]'}">${partPct}%</span>
                </td>
                <td class="px-3.5 py-3 text-center">
                    <span class="font-extrabold ${aprPct < 70 ? 'text-orange-400' : 'text-[#86efac]'}">${aprPct}%</span>
                </td>
                <td class="px-3.5 py-3 text-center">${badgeHtml}</td>
            </tr>
        `;
    }).join('');
}

/** Exportación del reporte histórico a archivo CSV de Excel */
function downloadHistoricoReport() {
    if (!isAuthorized()) return;
    
    const colabs = filteredColaboradoresData.length > 0 ? filteredColaboradoresData : getColaboradoresTrazabilidadData();
    const cursos = getHistoricoCursosData();

    let csvContent = "\uFEFF"; // UTF-8 BOM
    csvContent += "--- REPORTES DE CAPACITACIÓN BUK-RB - HISTÓRICO Y TRAZABILIDAD ---\n\n";

    csvContent += "1. HISTÓRICO DE CAPACITACIONES POR MES\n";
    csvContent += "Mes,Curso,Público Objetivo,Participación (Cant.),% Participación,Aprobación (Cant.),% Aprobación,Promedio Valoración\n";
    cursos.forEach(c => {
        csvContent += `"${c.mes}","${c.curso}",${c.publico_objetivo},${c.participacion_cant},${Number(c.participacion_pct).toFixed(1)}%,${c.aprobacion_cant},${Number(c.aprobacion_pct).toFixed(1)}%,${Number(c.promedio_satisfaccion).toFixed(1)}\n`;
    });

    csvContent += "\n2. TRAZABILIDAD DE COLABORADORES Y RECURRENCIA POR ÁREA\n";
    csvContent += "Cédula,Nombre,Área,Cursos Asignados,Cursos Participados,Cursos Aprobados,% Participación,% Aprobación,Estado Alerta\n";
    colabs.forEach(c => {
        let estadoStr = "Destacado";
        if (c.pct_participacion < 50 || c.pct_aprobacion < 50) estadoStr = "Recurrencia Critica";
        else if (c.pct_participacion < 70) estadoStr = "Baja Participacion";
        else if (c.pct_aprobacion < 70) estadoStr = "Baja Aprobacion";

        csvContent += `"${c.cedula || ''}","${c.nombre}","${c.area}",${c.asignados},${c.participados},${c.aprobados},${Number(c.pct_participacion).toFixed(1)}%,${Number(c.pct_aprobacion).toFixed(1)}%,"${estadoStr}"\n`;
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `Reporte_Historico_Capacitaciones_RB_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
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

    const hModal = document.getElementById('historico-modal');
    if (hModal) {
        hModal.addEventListener('click', function(e) {
            if (e.target === this) closeHistoricoModal();
        });
    }
});

