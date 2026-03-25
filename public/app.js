document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('fileInput');
    const uploadSection = document.getElementById('uploadSection');
    const settingsPanel = document.getElementById('settingsPanel');
    const carsSection = document.getElementById('carsSection');
    
    const marginRange = document.getElementById('marginRange');
    const marginValue = document.getElementById('marginValue');
    const maxDamage = document.getElementById('maxDamage');
    
    const carTableBody = document.getElementById('carTableBody');
    const carCountBadge = document.getElementById('carCountBadge');
    const auctionNameLabel = document.getElementById('auctionName');
    const headerStats = document.getElementById('headerStats');
    const summaryCards = document.getElementById('summaryCards');
    
    // Stats elements
    const totalCarsEl = document.getElementById('totalCars');
    const analyzedCarsEl = document.getElementById('analyzedCars');
    const profitableCarsEl = document.getElementById('profitableCars');
    const summaryProfitable = document.getElementById('summaryProfitable');
    const summaryAvgProfit = document.getElementById('summaryAvgProfit');
    const summaryConfidence = document.getElementById('summaryConfidence');
    const summaryBest = document.getElementById('summaryBest');
    
    // Progress bar
    const progressContainer = document.getElementById('progressContainer');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const progressCount = document.getElementById('progressCount');
    
    // Buttons
    const btnNewFile = document.getElementById('btnNewFile');
    const btnAnalyzeAll = document.getElementById('btnAnalyzeAll');
    
    // Search and Filters
    const searchInput = document.getElementById('searchInput');
    const filterChips = document.querySelectorAll('.chip');
    
    // Modal
    const modalOverlay = document.getElementById('modalOverlay');
    const modalClose = document.getElementById('modalClose');
    const modalContent = document.getElementById('modalContent');
    const toastContainer = document.getElementById('toastContainer');

    // State
    let carsData = [];
    let currentFilter = 'all';
    let currentSearch = '';
    let sortConfig = { key: null, direction: 'asc' };

    // --- INITIALIZATION ---
    initParticles();

    // --- EVENT LISTENERS ---
    
    // Margin range slider
    marginRange.addEventListener('input', (e) => {
        marginValue.textContent = e.target.value + '%';
        if (carsData.length > 0) {
            // Re-calculate for already analyzed cars
            recalculateMargins();
        }
    });

    // Upload interactions
    uploadZone.addEventListener('click', () => fileInput.click());
    
    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });
    
    uploadZone.addEventListener('dragleave', () => {
        uploadZone.classList.remove('dragover');
    });
    
    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleFileUpload(e.dataTransfer.files[0]);
        }
    });
    
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFileUpload(e.target.files[0]);
        }
    });

    // Reset button
    btnNewFile.addEventListener('click', () => {
        carsData = [];
        uploadSection.style.display = 'flex';
        settingsPanel.style.display = 'none';
        carsSection.style.display = 'none';
        headerStats.style.display = 'none';
        summaryCards.style.display = 'none';
        progressContainer.style.display = 'none';
        fileInput.value = '';
    });

    // Analyze All button
    btnAnalyzeAll.addEventListener('click', () => {
        analyzeAllCars();
    });

    // Filtering
    filterChips.forEach(chip => {
        chip.addEventListener('click', () => {
            filterChips.forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            currentFilter = chip.dataset.filter;
            renderTable();
        });
    });

    // Search
    searchInput.addEventListener('input', (e) => {
        currentSearch = e.target.value.toLowerCase();
        renderTable();
    });

    // Sorting
    document.querySelectorAll('.car-table th.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const key = th.dataset.sort;
            if (sortConfig.key === key) {
                sortConfig.direction = sortConfig.direction === 'asc' ? 'desc' : 'asc';
            } else {
                sortConfig.key = key;
                sortConfig.direction = 'desc'; // Default to desc for most numbers
            }
            
            // Update UI
            document.querySelectorAll('.car-table th.sortable').forEach(h => {
                h.classList.remove('active');
                h.textContent = h.textContent.replace(' ↑', ' ↕').replace(' ↓', ' ↕');
            });
            
            th.classList.add('active');
            th.textContent = th.textContent.replace(' ↕', sortConfig.direction === 'asc' ? ' ↑' : ' ↓');
            
            renderTable();
        });
    });

    // Modal close
    modalClose.addEventListener('click', () => {
        modalOverlay.style.animation = 'fadeOut 0.3s forwards';
        setTimeout(() => {
            modalOverlay.style.display = 'none';
            modalOverlay.style.animation = 'modalFadeIn 0.3s forwards';
        }, 300);
    });

    // Close modal on click outside
    modalOverlay.addEventListener('click', (e) => {
        if (e.target === modalOverlay) modalClose.click();
    });

    // --- CORE FUNCTIONS ---

    async function handleFileUpload(file) {
        if (!file.name.endsWith('.xls') && !file.name.endsWith('.xlsx')) {
            showToast('Por favor, sube un archivo Excel válido (.xls, .xlsx)', 'error');
            return;
        }

        const formData = new FormData();
        formData.append('file', file);

        uploadZone.style.opacity = '0.5';
        uploadZone.textContent = 'Procesando...';

        try {
            const res = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });
            const data = await res.json();
            
            if (data.error) throw new Error(data.error);
            
            // Filter and prepare data
            const maxDmgAllowed = parseInt(maxDamage.value);
            carsData = data.cars
                .filter(car => car.damageLevel <= maxDmgAllowed)
                .map(car => ({
                    ...car,
                    status: 'pending', // pending, analyzing, complete, error
                    result: null
                }));

            // Update UI
            uploadSection.style.display = 'none';
            carsSection.style.display = 'flex';
            settingsPanel.style.display = 'block';
            headerStats.style.display = 'flex';
            
            auctionNameLabel.textContent = data.auctionName;
            
            updateStats();
            renderTable();
            showToast(`Archivo cargado: ${carsData.length} coches listos para analizar.`, 'success');
            
            // Move settings panel above table
            const tableContainer = document.querySelector('.table-container');
            carsSection.insertBefore(settingsPanel, tableContainer);

        } catch (err) {
            showToast('Error al procesar el archivo: ' + err.message, 'error');
            // Reset upload zone
            uploadZone.style.opacity = '1';
            uploadZone.innerHTML = `
                <div class="upload-zone-inner">
                    <h3 class="upload-title">Arrastra tu archivo Excel aquí</h3>
                    <p class="upload-formats">Formatos: .xls, .xlsx</p>
                </div>
            `;
        }
    }

    async function analyzeSingleCar(id) {
        const carIndex = carsData.findIndex(c => c.id === id);
        if (carIndex === -1) return;
        
        const car = carsData[carIndex];
        car.status = 'analyzing';
        renderTable(); // Update row to spinning state

        try {
            const res = await fetch('/api/analyze', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    car, 
                    targetMargin: parseInt(marginRange.value) / 100 
                })
            });
            
            const data = await res.json();
            if (data.error) throw new Error(data.error);
            
            carsData[carIndex].status = 'complete';
            carsData[carIndex].result = data;
            
        } catch (err) {
            carsData[carIndex].status = 'error';
            carsData[carIndex].error = err.message;
            showToast(`Error al analizar ${car.marca} ${car.modelo}: ${err.message}`, 'error');
        }
        
        updateStats();
        renderTable();
    }

    async function analyzeAllCars() {
        const pendingCars = carsData.filter(c => c.status === 'pending' || c.status === 'error');
        if (pendingCars.length === 0) return showToast('No hay coches pendientes', 'info');

        btnAnalyzeAll.disabled = true;
        progressContainer.style.display = 'flex';
        summaryCards.style.display = 'grid';
        
        // Start EventSource connection for batch processing
        try {
            const response = await fetch('/api/analyze-batch', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    cars: pendingCars,
                    targetMargin: parseInt(marginRange.value) / 100 
                })
            });

            if (!response.ok) throw new Error('Error en el servidor');

            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            
            while (true) {
                const { value, done } = await reader.read();
                if (done) break;
                
                const chunk = decoder.decode(value);
                const lines = chunk.split('\n\n');
                
                for (let line of lines) {
                    if (line.startsWith('data: ')) {
                        try {
                            const data = JSON.parse(line.substring(6));
                            
                            if (data.type === 'progress') {
                                progressText.textContent = `Analizando: ${data.car}`;
                                progressCount.textContent = `${data.current}/${data.total}`;
                                const percent = (data.current / data.total) * 100;
                                progressFill.style.width = `${percent}%`;
                                
                                // Mark current as analyzing
                                const carToAnalyze = pendingCars[data.current - 1];
                                const idx = carsData.findIndex(c => c.id === carToAnalyze.id);
                                if (idx > -1) {
                                    carsData[idx].status = 'analyzing';
                                    renderTable();
                                }
                                
                            } else if (data.type === 'result') {
                                const idx = carsData.findIndex(c => c.id === data.result.car.id);
                                if (idx > -1) {
                                    carsData[idx].status = data.result.error ? 'error' : 'complete';
                                    carsData[idx].result = data.result;
                                    carsData[idx].error = data.result.error;
                                    updateStats();
                                    renderTable(); // Re-render table on each result to see progress
                                }
                            } else if (data.type === 'complete') {
                                progressText.textContent = 'Análisis completado 👋';
                                btnAnalyzeAll.disabled = false;
                                showToast('Análisis masivo completado', 'success');
                                setTimeout(() => progressContainer.style.display = 'none', 3000);
                            } else if (data.type === 'error') {
                                throw new Error(data.message);
                            }
                        } catch (e) {
                            console.error('Error parsing SSE data:', e, line);
                        }
                    }
                }
            }
            
        } catch (err) {
            showToast('Error en análisis masivo: ' + err.message, 'error');
            btnAnalyzeAll.disabled = false;
        }
    }

    // --- RE-CALCULATE LOCAL ---
    function recalculateMargins() {
        const margin = parseInt(marginRange.value) / 100;
        
        carsData.forEach(car => {
            if (car.status === 'complete' && car.result && car.result.priceAnalysis.marketPrice > 0) {
                // Simplified client-side recalc
                const mrkPrice = car.result.priceAnalysis.marketPrice;
                const repairCost = car.result.buyCalculation.repairCost || 1200; // default Moderate
                const fixed = 350 + 300 + repairCost;
                
                const buyP = Math.round(mrkPrice / (1 + margin) - fixed);
                const estProf = mrkPrice - buyP - fixed;
                
                car.result.buyCalculation.buyPrice = Math.max(0, buyP);
                car.result.buyCalculation.estimatedProfit = Math.round(estProf);
                car.result.buyCalculation.actualMargin = Math.round((estProf / buyP) * 1000) / 10;
                car.result.buyCalculation.targetMargin = margin * 100;
            }
        });
        
        updateStats();
        renderTable();
    }

    // --- RENDER ---

    function updateStats() {
        const total = carsData.length;
        const analyzed = carsData.filter(c => c.status === 'complete').length;
        
        let profitable = 0;
        let totalProfit = 0;
        let totalConf = 0;
        let bestCar = null;
        let bestProfit = 0;

        carsData.forEach(car => {
            if (car.status === 'complete' && car.result && !car.result.error) {
                const profit = car.result.buyCalculation.estimatedProfit;
                const buyP = car.result.buyCalculation.buyPrice;
                const conf = car.result.priceAnalysis.confidence;
                
                if (profit > 0 && buyP > 0) {
                    profitable++;
                    totalProfit += profit;
                    totalConf += conf;
                    
                    if (profit > bestProfit && conf > 40) {
                        bestProfit = profit;
                        bestCar = car;
                    }
                }
            }
        });

        // Header
        totalCarsEl.textContent = total;
        carCountBadge.textContent = total;
        analyzedCarsEl.textContent = analyzed;
        profitableCarsEl.textContent = profitable;

        // Summary cards
        summaryProfitable.textContent = profitable;
        summaryAvgProfit.textContent = profitable > 0 ? formatMoney(Math.round(totalProfit / profitable)) : '0€';
        summaryConfidence.textContent = profitable > 0 ? Math.round(totalConf / profitable) + '%' : '0%';
        
        if (bestCar) {
            summaryBest.innerHTML = `<span style="font-size:0.6em;display:block;color:var(--text-secondary)">${bestCar.marca} ${bestCar.modelo}</span>${formatMoney(bestProfit)}`;
        } else {
            summaryBest.textContent = '-';
        }
    }

    function formatMoney(amount) {
        if (!amount && amount !== 0) return '-';
        return new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(amount);
    }
    
    function formatNumber(num) {
        if (!num && num !== 0) return '-';
        return new Intl.NumberFormat('es-ES').format(num);
    }

    function renderTable() {
        // Filter
        let filtered = carsData.filter(car => {
            // Search filter
            if (currentSearch) {
                const text = `${car.marca} ${car.modelo} ${car.matricula}`.toLowerCase();
                if (!text.includes(currentSearch)) return false;
            }
            
            // Tab filter
            if (currentFilter === 'profitable') {
                return car.status === 'complete' && car.result && car.result.buyCalculation.buyPrice > 0;
            } else if (currentFilter === 'analyzed') {
                return car.status === 'complete';
            } else if (currentFilter === 'pending') {
                return car.status === 'pending' || car.status === 'error';
            }
            
            return true;
        });

        // Sort
        if (!sortConfig.key) {
            sortConfig = { key: 'market', direction: 'asc' };
        }
        
        if (sortConfig.key) {
            filtered.sort((a, b) => {
                let valA, valB;
                
                if (sortConfig.key === 'marca') { valA = a.marca; valB = b.marca; }
                else if (sortConfig.key === 'year') { valA = a.year; valB = b.year; }
                else if (sortConfig.key === 'km') { valA = a.kilometraje; valB = b.kilometraje; }
                else if (sortConfig.key === 'market') { 
                    valA = a.result?.priceAnalysis.marketPrice || 0; 
                    valB = b.result?.priceAnalysis.marketPrice || 0; 
                }
                else if (sortConfig.key === 'buy') { 
                    valA = a.result?.buyCalculation.buyPrice || 0; 
                    valB = b.result?.buyCalculation.buyPrice || 0; 
                }
                else if (sortConfig.key === 'profit') { 
                    valA = a.result?.buyCalculation.estimatedProfit || 0; 
                    valB = b.result?.buyCalculation.estimatedProfit || 0; 
                }

                if (valA < valB) return sortConfig.direction === 'asc' ? -1 : 1;
                if (valA > valB) return sortConfig.direction === 'asc' ? 1 : -1;
                return 0;
            });
        }

        // Render rows
        carTableBody.innerHTML = '';
        
        filtered.forEach((car, index) => {
            const tr = document.createElement('tr');
            
            let statusDot = 'gray';
            if (car.status === 'analyzing') statusDot = 'yellow';
            else if (car.status === 'error') statusDot = 'red';
            else if (car.status === 'complete') {
                statusDot = (car.result && car.result.buyCalculation.buyPrice > 0) ? 'green' : 'gray';
            }

            const isProfitable = car.status === 'complete' && car.result && car.result.buyCalculation.buyPrice > 0;
            if (isProfitable) tr.classList.add('row-profitable');
            
            let marketHTML = '-';
            let buyHTML = '-';
            let profitHTML = '-';
            let confHTML = '-';
            let actionsHTML = '';

            if (car.status === 'pending') {
                actionsHTML = `<button class="btn-action" onclick="window.analyzeSingleCar(${car.id})">Analizar</button>`;
            } else if (car.status === 'analyzing') {
                actionsHTML = `<button class="btn-action analyzing" disabled>⏳</button>`;
            } else if (car.status === 'error') {
                marketHTML = '<span style="color:var(--danger);font-size:0.8em">Error</span>';
                actionsHTML = `<button class="btn-action" onclick="window.analyzeSingleCar(${car.id})" title="${car.error}">Reintentar</button>`;
            } else if (car.status === 'complete') {
                if (car.result.priceAnalysis.marketPrice > 0) {
                    marketHTML = formatMoney(car.result.priceAnalysis.marketPrice);
                    buyHTML = formatMoney(car.result.buyCalculation.buyPrice);
                    
                    const pClass = isProfitable ? 'positive' : 'negative';
                    profitHTML = `<span class="td-profit ${pClass}">${formatMoney(car.result.buyCalculation.estimatedProfit)}</span>`;
                    
                    const conf = car.result.priceAnalysis.confidence;
                    let confColor = conf > 70 ? 'linear-gradient(90deg, #10b981, #34d399)' : 
                                    (conf > 40 ? 'linear-gradient(90deg, #f59e0b, #fbbf24)' : 'linear-gradient(90deg, #ef4444, #f87171)');
                                    
                    confHTML = `
                        <div class="conf-wrapper" title="${car.result.priceAnalysis.method}">
                            <div class="conf-bar-bg">
                                <div class="conf-fill" style="width: ${conf}%; background: ${confColor}"></div>
                            </div>
                            <span class="conf-val">${conf}%</span>
                        </div>
                    `;
                } else {
                    marketHTML = '<span style="color:var(--text-secondary);font-size:0.8em">Sin datos</span>';
                }
                actionsHTML = `<button class="btn-action success" onclick="window.showCarDetails(${car.id})">Detalles</button>
                               <a href="${car.result.scrapeResult.url}" target="_blank" class="btn-action" style="text-decoration:none; background:#3b82f6; display:inline-block; margin-left:4px" title="Ver búsqueda en coches.net">🔗 Buscar</a>`;
            }

            // Damage badge class
            const dmgClass = car.damageLevel ? `dmg-badge dmg-${car.damageLevel}` : '';
            const dmgText = car.damageLevel > 0 ? `Lv ${car.damageLevel}` : (car.clasificacionDanos || '-');

            tr.innerHTML = `
                <td class="th-narrow">
                    <div title="${car.status}" class="status-dot ${statusDot}"></div>
                </td>
                <td class="text-mono">${car.matricula || '-'}</td>
                <td style="font-weight:600">${car.marca}</td>
                <td class="td-model" title="${car.modelo} ${car.derivado}">${car.modelo} <span style="color:var(--text-secondary);font-size:0.9em">${car.derivado || ''}</span></td>
                <td>${car.combustible}</td>
                <td>${car.year || '-'}</td>
                <td class="text-mono">${formatNumber(car.kilometraje)} km</td>
                <td>${car.cv} CV</td>
                <td><span class="${dmgClass}">${dmgText}</span></td>
                <td class="td-money">${marketHTML}</td>
                <td class="td-money">${buyHTML}</td>
                <td class="td-money">${profitHTML}</td>
                <td>${confHTML}</td>
                <td style="text-align:center">${actionsHTML}</td>
            `;
            
            carTableBody.appendChild(tr);
        });
    }

    // Modal Details
    window.showCarDetails = function(id) {
        const car = carsData.find(c => c.id === id);
        if (!car || !car.result) return;
        
        const res = car.result;
        const pa = res.priceAnalysis;
        const bc = res.buyCalculation;
        
        const bcaLink = car.bcaLink ? `<a href="https://${car.bcaLink}" target="_blank" class="btn btn-secondary" style="font-size:0.8rem">🔗 Ver en BCA</a>` : '';

        modalContent.innerHTML = `
            <div class="m-header">
                <div style="display:flex; justify-content:space-between; align-items:flex-start">
                    <div>
                        <div class="m-badge-list">
                            <span class="badge" style="background:var(--accent-glow);color:white">${car.matricula}</span>
                            <span class="badge">${car.combustible}</span>
                            <span class="badge">${car.year}</span>
                            <span class="badge">${formatNumber(car.kilometraje)} km</span>
                            <span class="badge" style="border-color:${car.damageLevel > 3 ? 'var(--danger)' : 'var(--warning)'}">Daño Nivel ${car.damageLevel}</span>
                        </div>
                        <h2 class="m-title">${car.marca} ${car.modelo} - ${car.motor}</h2>
                        <p style="color:var(--text-secondary)">${car.derivado}</p>
                    </div>
                    ${bcaLink}
                </div>
            </div>

            <div class="m-grid">
                <!-- PANEL 1: Mercado -->
                <div class="m-panel">
                    <h3>📈 Análisis de Mercado (Coches.net)</h3>
                    <div class="m-stats-list">
                        <div class="m-stat">
                            <span class="m-stat-label">Anuncios encontrados</span>
                            <span class="m-stat-val">${res.scrapeResult.pricesFound}</span>
                        </div>
                        <div class="m-stat">
                            <span class="m-stat-label">Método estadístico</span>
                            <span class="m-stat-val" style="color:#a78bfa">${pa.method}</span>
                        </div>
                        <div class="m-stat">
                            <span class="m-stat-label">Confianza del dato</span>
                            <span class="m-stat-val">${pa.confidence}%</span>
                        </div>
                        <div class="m-stat" style="margin-top:1rem">
                            <span class="m-stat-label">Precio Mínimo Scrap</span>
                            <span class="m-stat-val">${formatMoney(pa.stats.min)}</span>
                        </div>
                        <div class="m-stat">
                            <span class="m-stat-label">Precio Máximo Scrap</span>
                            <span class="m-stat-val">${formatMoney(pa.stats.max)}</span>
                        </div>
                        <div class="m-stat" style="background:rgba(96,165,250,0.1); padding:0.5rem; margin-top:0.5rem; border-radius:4px; border:1px solid rgba(96,165,250,0.3)">
                            <span class="m-stat-label" style="color:white; font-weight:600">PRECIO DE MERCADO ESTIMADO</span>
                            <span class="m-stat-val highlight-blue" style="font-size:1.2rem">${formatMoney(pa.marketPrice)}</span>
                        </div>
                        
                        <a href="${res.scrapeResult.url}" target="_blank" class="m-source-link">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6M15 3h6v6M10 14L21 3"/></svg>
                            Ver búsqueda en coches.net
                        </a>
                    </div>
                    ${pa.cheapestPrices && pa.cheapestPrices.length > 0 ? `
                        <div style="margin-top:1.5rem; padding-top:1.5rem; border-top:1px solid var(--border-color)">
                            <h4 style="color:var(--text-secondary);margin-bottom:0.8rem;font-size:0.8rem;letter-spacing:0.05rem">LOS 5 MÁS BARATOS ACTUALES (usados):</h4>
                            <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:0.5rem">
                                ${pa.cheapestPrices.map(p => `
                                    <div style="background:rgba(16,185,129,0.1); padding:0.4rem; border-radius:6px; border:1px solid rgba(16,185,129,0.3); text-align:center">
                                        <span style="font-weight:700; color:#10b981; font-size:0.9rem">${formatMoney(p)}</span>
                                    </div>
                                `).join('')}
                            </div>
                        </div>
                    ` : ''}
                </div>

                <!-- PANEL 2: Rentabilidad -->
                <div class="m-panel" style="border-color:${bc.estimatedProfit > 0 ? 'var(--success)' : 'var(--danger)'}">
                    <h3 style="color:${bc.estimatedProfit > 0 ? 'var(--success)' : 'var(--danger)'}">💰 Cálculo de Compra (Margen ${bc.targetMargin}%)</h3>
                    <div class="m-stats-list">
                        <div class="m-stat">
                            <span class="m-stat-label">Precio de Venta (Mercado)</span>
                            <span class="m-stat-val highlight-blue">${formatMoney(bc.sellingPrice)}</span>
                        </div>
                        <div class="m-stat">
                            <span class="m-stat-label">Coste Reparación Estimado</span>
                            <span class="m-stat-val" style="color:var(--danger)">- ${formatMoney(bc.repairCost)}</span>
                        </div>
                        <div class="m-stat">
                            <span class="m-stat-label">Otros Costes (Traspaso, ITV)</span>
                            <span class="m-stat-val" style="color:var(--danger)">- ${formatMoney(bc.totalFixedCosts - bc.repairCost)}</span>
                        </div>
                        <div class="m-stat" style="margin-top:1rem">
                            <span class="m-stat-label">BENEFICIO NETO OBJETIVO</span>
                            <span class="m-stat-val" style="color:${bc.estimatedProfit > 0 ? 'var(--success)' : 'var(--danger)'}">${formatMoney(bc.estimatedProfit)}</span>
                        </div>
                        <div class="m-stat" style="background:${bc.estimatedProfit > 0 ? 'rgba(16,185,129,0.1)' : 'rgba(239,68,68,0.1)'}; padding:1rem; border-radius:8px; border:1px solid ${bc.estimatedProfit > 0 ? 'var(--success)' : 'var(--danger)'}; margin-top:1rem; align-items:center;">
                            <span class="m-stat-label" style="color:white; font-size:1.1rem; font-weight:700">PRECIO DE PUJA MÁXIMO</span>
                            <span class="m-stat-val highlight-green" style="font-size:1.8rem; color:${bc.estimatedProfit > 0 ? 'var(--success)' : 'var(--danger)'}">${formatMoney(bc.buyPrice)}</span>
                        </div>
                    </div>
                    
                    ${bc.buyPrice <= 0 ? '<p style="color:var(--danger); font-size:0.875rem; margin-top:1rem; background:var(--danger-bg); padding:0.5rem; border-radius:4px">⚠️ Este coche no es rentable con el margen deseado o los costes de reparación son muy altos.</p>' : ''}
                </div>
            </div>
            
            <div style="margin-top:2rem; border-top:1px solid var(--border); padding-top:1rem">
                <h4 style="margin-bottom:0.5rem; color:var(--text-secondary)">Equipamiento Extra:</h4>
                <p style="font-size:0.875rem; color:#9ca3af; line-height:1.6">${car.equipamiento || 'No especificado'}</p>
            </div>
        `;
        
        modalOverlay.style.display = 'flex';
    };

    // Attach to window so onclick works
    window.analyzeSingleCar = analyzeSingleCar;

    // Toast functionality
    function showToast(msg, type='info') {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        
        let icon = 'ℹ️';
        if (type === 'success') icon = '✅';
        if (type === 'error') icon = '❌';
        if (type === 'warning') icon = '⚠️';
        
        toast.innerHTML = `<span style="font-size:1.25rem">${icon}</span> <span>${msg}</span>`;
        toastContainer.appendChild(toast);
        
        setTimeout(() => {
            toast.style.animation = 'fadeOut 0.3s forwards';
            setTimeout(() => toast.remove(), 300);
        }, 5000);
    }

    // Particles background
    function initParticles() {
        const canvas = document.getElementById('particles-canvas');
        const ctx = canvas.getContext('2d');
        canvas.width = window.innerWidth;
        canvas.height = window.innerHeight;
        
        const particles = [];
        for (let i = 0; i < 50; i++) {
            particles.push({
                x: Math.random() * canvas.width,
                y: Math.random() * canvas.height,
                radius: Math.random() * 2 + 0.5,
                vx: (Math.random() - 0.5) * 0.5,
                vy: (Math.random() - 0.5) * 0.5
            });
        }
        
        function animate() {
            requestAnimationFrame(animate);
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.fillStyle = 'rgba(99, 102, 241, 0.4)';
            
            particles.forEach(p => {
                p.x += p.vx;
                p.y += p.vy;
                
                if (p.x < 0 || p.x > canvas.width) p.vx *= -1;
                if (p.y < 0 || p.y > canvas.height) p.vy *= -1;
                
                ctx.beginPath();
                ctx.arc(p.x, p.y, p.radius, 0, Math.PI * 2);
                ctx.fill();
            });
        }
        animate();
        
        window.addEventListener('resize', () => {
            canvas.width = window.innerWidth;
            canvas.height = window.innerHeight;
        });
    }
});
