const LABOR_RATE = 22.00;
const STORAGE_KEY = 'sea_est_manual_records';

let allData = { records: [], summary: {} };
let manualRecords = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
let filteredRecords = [];
let dailyCostChart = null;
let dailyYieldChart = null;
let totalCostChart = null;
let currentSort = { field: 'date', dir: 'desc' };
let dailySort = { field: null, dir: 'asc' };
let dailyRowData = [];

// ---- INITIALIZATION ----
document.addEventListener('DOMContentLoaded', async () => {
    setupNavigation();
    setupEntryForm();
    setupDetailControls();
    await loadData();
    applyFilters();
});

async function loadData() {
    try {
        const resp = await fetch('data/production_data.json');
        allData = await resp.json();
    } catch (e) {
        allData = { records: [], summary: {}, labor_rate: LABOR_RATE };
    }
    const combined = [...allData.records, ...manualRecords];
    allData.records = combined;
    document.getElementById('data-timestamp').textContent = allData.generated_at
        ? new Date(allData.generated_at).toLocaleDateString()
        : 'N/A';
    populateFilters();
}

// ---- NAVIGATION ----
function setupNavigation() {
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById('page-' + btn.dataset.page).classList.add('active');
        });
    });
}

// ---- FILTERS ----
function populateFilters() {
    const classifications = [...new Set(allData.records.map(r => r.classification).filter(Boolean))].sort();
    const activities = [...new Set(allData.records.map(r => r.activity))].sort();
    const products = [...new Set(allData.records.map(r => r.product_format))].sort();
    const suppliers = [...new Set(allData.records.map(r => r.supplier).filter(Boolean))].sort();
    const dates = allData.records.map(r => r.date).sort();

    fillSelect('filter-classification', classifications);
    fillSelect('filter-activity', activities);
    buildProductMultiSelect(products);
    fillSelect('filter-supplier', suppliers);

    if (dates.length) {
        document.getElementById('filter-date-start').value = dates[0];
        document.getElementById('filter-date-end').value = dates[dates.length - 1];
    }

    ['filter-classification', 'filter-activity', 'filter-supplier', 'filter-date-start', 'filter-date-end'].forEach(id => {
        document.getElementById(id).addEventListener('change', applyFilters);
    });
    document.getElementById('btn-reset-filters').addEventListener('click', resetFilters);
}

function buildProductMultiSelect(products) {
    const wrapper = document.getElementById('filter-product-wrapper');
    const display = document.getElementById('filter-product-display');
    const dropdown = document.getElementById('filter-product-dropdown');

    dropdown.innerHTML = '';
    products.forEach(p => {
        const lbl = document.createElement('label');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = p;
        cb.addEventListener('change', () => {
            updateProductDisplay();
            applyFilters();
        });
        lbl.appendChild(cb);
        lbl.appendChild(document.createTextNode(p));
        dropdown.appendChild(lbl);
    });

    display.addEventListener('click', (e) => {
        e.stopPropagation();
        wrapper.classList.toggle('open');
    });
    document.addEventListener('click', (e) => {
        if (!wrapper.contains(e.target)) wrapper.classList.remove('open');
    });
}

function updateProductDisplay() {
    const checked = getSelectedProducts();
    const display = document.getElementById('filter-product-display');
    if (checked.length === 0) {
        display.textContent = 'All Formats';
    } else if (checked.length === 1) {
        display.textContent = checked[0];
    } else {
        display.textContent = checked.length + ' formats selected';
    }
}

function getSelectedProducts() {
    const checkboxes = document.querySelectorAll('#filter-product-dropdown input[type="checkbox"]');
    return [...checkboxes].filter(cb => cb.checked).map(cb => cb.value);
}

function fillSelect(id, values) {
    const sel = document.getElementById(id);
    const current = sel.value;
    while (sel.options.length > 1) sel.remove(1);
    values.forEach(v => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = v;
        sel.appendChild(opt);
    });
    sel.value = current;
}

function resetFilters() {
    document.getElementById('filter-classification').value = 'all';
    document.getElementById('filter-activity').value = 'all';
    document.querySelectorAll('#filter-product-dropdown input[type="checkbox"]').forEach(cb => cb.checked = false);
    updateProductDisplay();
    document.getElementById('filter-supplier').value = 'all';
    const dates = allData.records.map(r => r.date).sort();
    if (dates.length) {
        document.getElementById('filter-date-start').value = dates[0];
        document.getElementById('filter-date-end').value = dates[dates.length - 1];
    }
    applyFilters();
}

function applyFilters() {
    const classification = document.getElementById('filter-classification').value;
    const activity = document.getElementById('filter-activity').value;
    const selectedProducts = getSelectedProducts();
    const supplier = document.getElementById('filter-supplier').value;
    const dateStart = document.getElementById('filter-date-start').value;
    const dateEnd = document.getElementById('filter-date-end').value;

    filteredRecords = allData.records.filter(r => {
        if (classification !== 'all' && r.classification !== classification) return false;
        if (activity !== 'all' && r.activity !== activity) return false;
        if (selectedProducts.length > 0 && !selectedProducts.includes(r.product_format)) return false;
        if (supplier !== 'all' && r.supplier !== supplier) return false;
        if (dateStart && r.date < dateStart) return false;
        if (dateEnd && r.date > dateEnd) return false;
        return true;
    });

    updateKPIs();
    updateTotalCostChart();
    updateDailyCostChart();
    updateDailyYieldChart();
    populateDailyFilters();
    updateDailyBreakdownTable();
    updateDetailTable();
}

// ---- KPIs ----
function updateKPIs() {
    const laborCosts = filteredRecords.map(r => r.cost_per_finished_lb).filter(c => c && c > 0);
    const totalCosts = filteredRecords.map(r => r.total_cost_per_finished_lb).filter(c => c && c > 0);
    const yieldLossCosts = filteredRecords.map(r => r.yield_loss_cost_per_lb).filter(c => c && c > 0);
    const yields = filteredRecords.map(r => r.yield_pct).filter(y => y && y > 0);
    const totalLbs = filteredRecords.reduce((s, r) => s + (r.finished_lbs || 0), 0);

    document.getElementById('kpi-total-cost').textContent = totalCosts.length ? '$' + avg(totalCosts).toFixed(4) : '--';
    document.getElementById('kpi-yield-loss-cost').textContent = yieldLossCosts.length ? '$' + avg(yieldLossCosts).toFixed(4) : '--';
    document.getElementById('kpi-avg-cost').textContent = laborCosts.length ? '$' + avg(laborCosts).toFixed(4) : '--';
    document.getElementById('kpi-avg-yield').textContent = yields.length ? avg(yields).toFixed(1) + '%' : '--';
    document.getElementById('kpi-total-lbs').textContent = totalLbs > 0 ? numberFmt(totalLbs.toFixed(0)) : '--';
    document.getElementById('kpi-count').textContent = filteredRecords.length || '--';
}

// ---- TOTAL COST STACKED BAR CHART ----
function updateTotalCostChart() {
    const recsWithCost = filteredRecords.filter(r => r.total_cost_per_finished_lb && r.total_cost_per_finished_lb > 0);
    const grouped = groupBy(recsWithCost, 'date');
    const dates = Object.keys(grouped).sort();

    const rawProtein = dates.map(d => {
        const vals = grouped[d].map(r => r.raw_protein_cost_per_lb).filter(Boolean);
        return vals.length ? avg(vals) : 0;
    });
    const yieldLoss = dates.map(d => {
        const vals = grouped[d].map(r => r.yield_loss_cost_per_lb).filter(Boolean);
        return vals.length ? avg(vals) : 0;
    });
    const labor = dates.map(d => {
        const vals = grouped[d].map(r => r.cost_per_finished_lb).filter(Boolean);
        return vals.length ? avg(vals) : 0;
    });

    // Determine if we should show target cost line
    const selectedActivity = document.getElementById('filter-activity').value;
    const selProducts = getSelectedProducts();
    let targetLine = null;

    if (selectedActivity === 'Skinning' && selProducts.length === 1) {
        if (selProducts[0].includes('ABF')) {
            targetLine = { value: 7.84, label: 'Target Cost (ABF): $7.84' };
        } else if (selProducts[0].includes('Atlantic') && !selProducts[0].includes('ABF')) {
            targetLine = { value: 7.52, label: 'Target Cost: $7.52' };
        }
    }

    const annotations = {};
    if (targetLine) {
        annotations.targetCostLine = {
            type: 'line',
            yMin: targetLine.value,
            yMax: targetLine.value,
            borderColor: '#dc2626',
            borderWidth: 2,
            borderDash: [8, 4],
            label: {
                display: true,
                content: targetLine.label,
                position: 'start',
                backgroundColor: 'rgba(220, 38, 38, 0.85)',
                color: '#fff',
                font: { size: 12, weight: 'bold' },
                padding: 4
            }
        };
    }

    if (totalCostChart) totalCostChart.destroy();
    totalCostChart = new Chart(document.getElementById('chart-total-cost'), {
        type: 'bar',
        data: {
            labels: dates.map(d => formatDate(d)),
            datasets: [
                { label: 'Raw Protein', data: rawProtein, backgroundColor: '#1a56db', stack: 'total' },
                { label: 'Yield Loss', data: yieldLoss, backgroundColor: '#d97706', stack: 'total' },
                { label: 'Labor', data: labor, backgroundColor: '#059669', stack: 'total' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            aspectRatio: 2.5,
            interaction: { mode: 'index', intersect: false },
            scales: {
                y: {
                    stacked: true,
                    title: { display: true, text: '$ / Finished Lb' },
                    ticks: { callback: v => '$' + v.toFixed(2) }
                },
                x: { stacked: true, ticks: { maxRotation: 45 } }
            },
            plugins: {
                annotation: { annotations },
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.dataset.label + ': $' + (ctx.parsed.y?.toFixed(4) || '0'),
                        afterBody: (items) => {
                            const total = items.reduce((s, i) => s + (i.parsed.y || 0), 0);
                            return 'Total: $' + total.toFixed(4);
                        }
                    }
                }
            }
        }
    });
}

// ---- DAILY COST CHART ----
function updateDailyCostChart() {
    const grouped = groupBy(filteredRecords, 'date');
    const dates = Object.keys(grouped).sort();
    const avgCosts = dates.map(d => {
        const costs = grouped[d].map(r => r.cost_per_finished_lb).filter(c => c && c > 0);
        return costs.length ? avg(costs) : null;
    });

    const productGroups = {};
    filteredRecords.forEach(r => {
        if (!r.cost_per_finished_lb || r.cost_per_finished_lb <= 0) return;
        const key = r.product_format || 'Unknown';
        if (!productGroups[key]) productGroups[key] = {};
        if (!productGroups[key][r.date]) productGroups[key][r.date] = [];
        productGroups[key][r.date].push(r.cost_per_finished_lb);
    });

    const colors = ['#1a56db', '#059669', '#d97706', '#dc2626', '#7c3aed', '#0891b2', '#be185d'];
    const datasets = Object.keys(productGroups).sort().map((prod, i) => ({
        label: prod,
        data: dates.map(d => {
            const vals = productGroups[prod][d];
            return vals ? avg(vals) : null;
        }),
        backgroundColor: colors[i % colors.length] + '80',
        borderColor: colors[i % colors.length],
        borderWidth: 1,
        barPercentage: 0.8,
    }));

    if (dailyCostChart) dailyCostChart.destroy();
    dailyCostChart = new Chart(document.getElementById('chart-daily-cost'), {
        type: 'bar',
        data: { labels: dates.map(d => formatDate(d)), datasets },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            aspectRatio: 2,
            interaction: { mode: 'index', intersect: false },
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: '$ / Finished Lb' },
                    ticks: { callback: v => '$' + v.toFixed(3) }
                },
                x: { ticks: { maxRotation: 45 } }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.dataset.label + ': $' + (ctx.parsed.y?.toFixed(4) || 'N/A')
                    }
                },
                legend: { position: 'top' }
            }
        }
    });
}

// ---- DAILY YIELD CHART ----
function updateDailyYieldChart() {
    const grouped = groupBy(filteredRecords.filter(r => r.yield_pct && r.yield_pct > 0), 'date');
    const dates = Object.keys(grouped).sort();

    const activityGroups = {};
    filteredRecords.filter(r => r.yield_pct && r.yield_pct > 0).forEach(r => {
        if (!activityGroups[r.activity]) activityGroups[r.activity] = {};
        if (!activityGroups[r.activity][r.date]) activityGroups[r.activity][r.date] = [];
        activityGroups[r.activity][r.date].push(r.yield_pct);
    });

    const colors = ['#1a56db', '#059669', '#d97706', '#dc2626'];
    const datasets = Object.keys(activityGroups).sort().map((act, i) => ({
        label: act,
        data: dates.map(d => {
            const vals = activityGroups[act][d];
            return vals ? avg(vals) : null;
        }),
        borderColor: colors[i % colors.length],
        backgroundColor: colors[i % colors.length] + '20',
        fill: false,
        tension: 0.3,
        pointRadius: 3,
        spanGaps: true,
    }));

    if (dailyYieldChart) dailyYieldChart.destroy();
    dailyYieldChart = new Chart(document.getElementById('chart-daily-yield'), {
        type: 'line',
        data: { labels: dates.map(d => formatDate(d)), datasets },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            aspectRatio: 2,
            interaction: { mode: 'index', intersect: false },
            scales: {
                y: {
                    title: { display: true, text: 'Yield %' },
                    ticks: { callback: v => v.toFixed(1) + '%' },
                    suggestedMin: 80,
                    suggestedMax: 105
                },
                x: { ticks: { maxRotation: 45 } }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: ctx => ctx.dataset.label + ': ' + (ctx.parsed.y?.toFixed(2) || 'N/A') + '%'
                    }
                }
            }
        }
    });
}

// ---- DAILY BREAKDOWN TABLE (with own filters) ----
function populateDailyFilters() {
    const activities = [...new Set(filteredRecords.map(r => r.activity))].sort();
    const products = [...new Set(filteredRecords.map(r => r.product_format))].sort();
    const suppliers = [...new Set(filteredRecords.map(r => r.supplier).filter(Boolean))].sort();
    const dates = filteredRecords.map(r => r.date).sort();

    fillSelect('daily-filter-activity', activities);
    fillSelect('daily-filter-product', products);
    fillSelect('daily-filter-supplier', suppliers);

    const startEl = document.getElementById('daily-filter-date-start');
    const endEl = document.getElementById('daily-filter-date-end');
    if (dates.length && !startEl.value) startEl.value = dates[0];
    if (dates.length && !endEl.value) endEl.value = dates[dates.length - 1];

    // Only bind listeners once
    if (!startEl.dataset.bound) {
        ['daily-filter-activity', 'daily-filter-product', 'daily-filter-supplier', 'daily-filter-date-start', 'daily-filter-date-end'].forEach(id => {
            document.getElementById(id).addEventListener('change', updateDailyBreakdownTable);
        });
        document.getElementById('btn-reset-daily').addEventListener('click', () => {
            document.getElementById('daily-filter-activity').value = 'all';
            document.getElementById('daily-filter-product').value = 'all';
            document.getElementById('daily-filter-supplier').value = 'all';
            const d = filteredRecords.map(r => r.date).sort();
            if (d.length) {
                document.getElementById('daily-filter-date-start').value = d[0];
                document.getElementById('daily-filter-date-end').value = d[d.length - 1];
            }
            updateDailyBreakdownTable();
        });
        startEl.dataset.bound = '1';
    }
}

function updateDailyBreakdownTable() {
    const dActivity = document.getElementById('daily-filter-activity').value;
    const dProduct = document.getElementById('daily-filter-product').value;
    const dSupplier = document.getElementById('daily-filter-supplier').value;
    const dStart = document.getElementById('daily-filter-date-start').value;
    const dEnd = document.getElementById('daily-filter-date-end').value;

    const recs = filteredRecords.filter(r => {
        if (dActivity !== 'all' && r.activity !== dActivity) return false;
        if (dProduct !== 'all' && r.product_format !== dProduct) return false;
        if (dSupplier !== 'all' && r.supplier !== dSupplier) return false;
        if (dStart && r.date < dStart) return false;
        if (dEnd && r.date > dEnd) return false;
        return true;
    });

    const groups = {};
    recs.forEach(r => {
        const key = r.date + '|' + r.activity + '|' + r.product_format + '|' + (r.supplier || '--');
        if (!groups[key]) groups[key] = [];
        groups[key].push(r);
    });

    dailyRowData = [];
    Object.keys(groups).sort().forEach(key => {
        const [date, activity, product, supplier] = key.split('|');
        const grp = groups[key];
        const totalCosts = grp.map(r => r.total_cost_per_finished_lb).filter(c => c && c > 0);
        const spreads = grp.map(r => r.production_spread_per_lb).filter(s => s != null);
        const extSpreads = grp.map(r => r.extended_production_spread).filter(s => s != null);
        const kpi = grp.find(r => r.target_cost)?.target_cost;
        const yields = grp.map(r => r.yield_pct).filter(y => y && y > 0);
        const totalLbs = grp.reduce((s, r) => s + (r.finished_lbs || 0), 0);
        const totalExtSpread = extSpreads.reduce((s, v) => s + v, 0);

        if (!totalCosts.length) return;

        dailyRowData.push({
            date, activity, product, supplier,
            avgTotal: avg(totalCosts),
            kpi,
            spread: spreads.length ? avg(spreads) : null,
            yield: yields.length ? avg(yields) : null,
            lbs: totalLbs,
            ext_spread: extSpreads.length ? totalExtSpread : null
        });
    });

    renderDailyRows();
    setupDailySortHeaders();
    updateWeeklyTable();
}

function renderDailyRows() {
    let rows = [...dailyRowData];
    if (dailySort.field) {
        rows.sort((a, b) => {
            let va = a[dailySort.field], vb = b[dailySort.field];
            if (va == null) va = dailySort.dir === 'asc' ? Infinity : -Infinity;
            if (vb == null) vb = dailySort.dir === 'asc' ? Infinity : -Infinity;
            if (typeof va === 'string') return dailySort.dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
            return dailySort.dir === 'asc' ? va - vb : vb - va;
        });
    }

    const tbody = document.querySelector('#table-daily tbody');
    tbody.innerHTML = '';
    rows.forEach(r => {
        const spreadClass = r.spread != null && r.spread >= 0 ? 'cost-normal' : r.spread != null ? 'cost-high' : '';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${formatDate(r.date)}</td>
            <td>${r.activity}</td>
            <td>${r.product}</td>
            <td>${r.supplier}</td>
            <td class="text-right">$${r.avgTotal.toFixed(4)}</td>
            <td class="text-right">${r.kpi ? '$' + r.kpi.toFixed(2) : '--'}</td>
            <td class="text-right ${spreadClass}">${r.spread != null ? '$' + r.spread.toFixed(4) : '--'}</td>
            <td class="text-right">${r.yield != null ? r.yield.toFixed(1) + '%' : '--'}</td>
            <td class="text-right">${numberFmt(r.lbs.toFixed(0))}</td>
            <td class="text-right" style="font-weight:600">${r.ext_spread != null ? '$' + numberFmt(r.ext_spread.toFixed(0)) : '--'}</td>
        `;
        tbody.appendChild(tr);
    });
}

function setupDailySortHeaders() {
    document.querySelectorAll('#table-daily th[data-sort-daily]').forEach(th => {
        if (th.dataset.sortBound) return;
        th.dataset.sortBound = '1';
        // Add sort icon span if not already there
        if (!th.querySelector('.sort-icon')) {
            const span = document.createElement('span');
            span.className = 'sort-icon';
            span.textContent = ' ↕';
            th.appendChild(span);
        }
        th.addEventListener('click', () => {
            const field = th.dataset.sortDaily;
            if (dailySort.field === field) {
                dailySort.dir = dailySort.dir === 'asc' ? 'desc' : 'asc';
            } else {
                dailySort = { field, dir: 'asc' };
            }
            document.querySelectorAll('#table-daily th[data-sort-daily]').forEach(h => {
                h.classList.remove('sort-asc', 'sort-desc');
                const icon = h.querySelector('.sort-icon');
                if (icon) icon.textContent = ' ↕';
            });
            th.classList.add(dailySort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
            const icon = th.querySelector('.sort-icon');
            if (icon) icon.textContent = dailySort.dir === 'asc' ? ' ▲' : ' ▼';
            renderDailyRows();
        });
    });
}

// ---- WEEKLY TABLE ----
function updateWeeklyTable() {
    // Apply same filters as daily breakdown
    const dActivity = document.getElementById('daily-filter-activity')?.value || 'all';
    const dProduct = document.getElementById('daily-filter-product')?.value || 'all';
    const dSupplier = document.getElementById('daily-filter-supplier')?.value || 'all';
    const dStart = document.getElementById('daily-filter-date-start')?.value || '';
    const dEnd = document.getElementById('daily-filter-date-end')?.value || '';

    const weeklyRecs = filteredRecords.filter(r => {
        if (dActivity !== 'all' && r.activity !== dActivity) return false;
        if (dProduct !== 'all' && r.product_format !== dProduct) return false;
        if (dSupplier !== 'all' && r.supplier !== dSupplier) return false;
        if (dStart && r.date < dStart) return false;
        if (dEnd && r.date > dEnd) return false;
        return true;
    });

    const groups = {};
    weeklyRecs.forEach(r => {
        const key = r.week + '|' + r.activity + '|' + r.product_format;
        if (!groups[key]) groups[key] = [];
        groups[key].push(r);
    });

    const tbody = document.querySelector('#table-weekly tbody');
    tbody.innerHTML = '';

    Object.keys(groups).sort().forEach(key => {
        const [week, activity, product] = key.split('|');
        const recs = groups[key];
        const totalCosts = recs.map(r => r.total_cost_per_finished_lb).filter(c => c && c > 0);
        const inputCosts = recs.map(r => r.input_cost_per_lb || r.raw_protein_cost_per_lb).filter(Boolean);
        const spreads = recs.map(r => r.production_spread_per_lb).filter(s => s != null);
        const extSpreads = recs.map(r => r.extended_production_spread).filter(s => s != null);
        const kpi = recs.find(r => r.target_cost)?.target_cost;
        const yields = recs.map(r => r.yield_pct).filter(y => y && y > 0);
        const totalLbs = recs.reduce((s, r) => s + (r.finished_lbs || 0), 0);
        const totalExtSpread = extSpreads.reduce((s, v) => s + v, 0);

        if (!totalCosts.length) return;

        const spreadClass = spreads.length && avg(spreads) >= 0 ? 'cost-normal' : 'cost-high';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${week}</td>
            <td>${activity}</td>
            <td>${product}</td>
            <td class="text-right">${inputCosts.length ? '$' + avg(inputCosts).toFixed(2) : '--'}</td>
            <td class="text-right">${totalCosts.length ? '$' + avg(totalCosts).toFixed(4) : '--'}</td>
            <td class="text-right">${kpi ? '$' + kpi.toFixed(2) : '--'}</td>
            <td class="text-right ${spreadClass}">${spreads.length ? '$' + avg(spreads).toFixed(4) : '--'}</td>
            <td class="text-right">${yields.length ? avg(yields).toFixed(1) + '%' : '--'}</td>
            <td class="text-right">${numberFmt(totalLbs.toFixed(0))}</td>
            <td class="text-right" style="font-weight:600">${extSpreads.length ? '$' + numberFmt(totalExtSpread.toFixed(0)) : '--'}</td>
        `;
        tbody.appendChild(tr);
    });
}

// ---- DETAIL TABLE ----
function updateDetailTable() {
    const tbody = document.querySelector('#table-detail tbody');
    const search = (document.getElementById('detail-search').value || '').toLowerCase();

    let recs = filteredRecords;
    if (search) {
        recs = recs.filter(r =>
            (r.lot || '').toLowerCase().includes(search) ||
            (r.supplier || '').toLowerCase().includes(search) ||
            (r.product_format || '').toLowerCase().includes(search) ||
            (r.activity || '').toLowerCase().includes(search)
        );
    }

    recs = [...recs].sort((a, b) => {
        let va = a[currentSort.field], vb = b[currentSort.field];
        if (va == null) va = '';
        if (vb == null) vb = '';
        if (typeof va === 'number' && typeof vb === 'number') {
            return currentSort.dir === 'asc' ? va - vb : vb - va;
        }
        va = String(va); vb = String(vb);
        return currentSort.dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });

    tbody.innerHTML = '';
    recs.slice(0, 500).forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${formatDate(r.date)}</td>
            <td>${r.activity}</td>
            <td>${r.supplier || '--'}</td>
            <td>${r.lot || '--'}</td>
            <td>${r.product_format}</td>
            <td class="text-right">${r.incoming_lbs?.toFixed(1) || '--'}</td>
            <td class="text-right">${r.finished_lbs?.toFixed(1) || '--'}</td>
            <td class="text-right">${r.yield_pct?.toFixed(1) || '--'}%</td>
            <td class="text-right">${r.raw_protein_cost_per_lb ? '$' + r.raw_protein_cost_per_lb.toFixed(2) : '--'}</td>
            <td class="text-right">${r.yield_loss_cost_per_lb ? '$' + r.yield_loss_cost_per_lb.toFixed(4) : '--'}</td>
            <td class="text-right">$${r.cost_per_finished_lb?.toFixed(4) || '--'}</td>
            <td class="text-right" style="font-weight:600">${r.total_cost_per_finished_lb ? '$' + r.total_cost_per_finished_lb.toFixed(4) : '--'}</td>
        `;
        tbody.appendChild(tr);
    });
}

function setupDetailControls() {
    document.getElementById('detail-search').addEventListener('input', updateDetailTable);

    document.querySelectorAll('#table-detail th[data-sort]').forEach(th => {
        th.addEventListener('click', () => {
            const field = th.dataset.sort;
            if (currentSort.field === field) {
                currentSort.dir = currentSort.dir === 'asc' ? 'desc' : 'asc';
            } else {
                currentSort = { field, dir: 'asc' };
            }
            updateDetailTable();
        });
    });

    document.getElementById('btn-export-csv').addEventListener('click', exportCSV);
}

function exportCSV() {
    const headers = ['Date', 'Activity', 'Supplier', 'Lot', 'Product', 'Incoming Lbs', 'Finished Lbs', 'Yield %', 'People', 'Hours', 'Raw Protein $/Lb', 'Protein Cost/Finished Lb', 'Yield Loss $/Lb', 'Labor $/Lb', 'Total $/Lb'];
    const rows = filteredRecords.map(r => [
        r.date, r.activity, r.supplier || '', r.lot || '', r.product_format,
        r.incoming_lbs, r.finished_lbs, r.yield_pct, r.people, r.hours_worked,
        r.raw_protein_cost_per_lb, r.protein_cost_per_finished_lb, r.yield_loss_cost_per_lb,
        r.cost_per_finished_lb, r.total_cost_per_finished_lb
    ]);

    let csv = headers.join(',') + '\n';
    rows.forEach(row => {
        csv += row.map(v => {
            const s = String(v ?? '');
            return s.includes(',') ? '"' + s + '"' : s;
        }).join(',') + '\n';
    });

    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'activity_cost_detail_' + new Date().toISOString().slice(0, 10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
}

// ---- DATA ENTRY ----
function setupEntryForm() {
    const form = document.getElementById('entry-form');
    const fields = ['entry-incoming', 'entry-finished', 'entry-people', 'entry-time-start', 'entry-time-end'];
    fields.forEach(id => document.getElementById(id).addEventListener('input', updateEntryPreview));

    form.addEventListener('submit', e => {
        e.preventDefault();
        saveEntry();
    });
    form.addEventListener('reset', () => {
        setTimeout(() => {
            document.getElementById('entry-preview').style.display = 'none';
            document.getElementById('entry-date').value = new Date().toISOString().slice(0, 10);
        }, 10);
    });

    document.getElementById('entry-date').value = new Date().toISOString().slice(0, 10);
    renderRecentEntries();
}

function updateEntryPreview() {
    const incoming = parseFloat(document.getElementById('entry-incoming').value);
    const finished = parseFloat(document.getElementById('entry-finished').value);
    const people = parseInt(document.getElementById('entry-people').value);
    const startTime = document.getElementById('entry-time-start').value;
    const endTime = document.getElementById('entry-time-end').value;

    const preview = document.getElementById('entry-preview');
    if (!incoming || !finished || !people || !startTime || !endTime) {
        preview.style.display = 'none';
        return;
    }

    const hours = calcHoursFromTimeInputs(startTime, endTime);
    const yieldPct = (finished / incoming * 100);
    const laborCost = people * hours * LABOR_RATE;
    const costPerLb = laborCost / finished;

    preview.style.display = 'block';
    document.getElementById('preview-yield').textContent = yieldPct.toFixed(1) + '%';
    document.getElementById('preview-hours').textContent = hours.toFixed(2) + ' hrs';
    document.getElementById('preview-labor-cost').textContent = '$' + laborCost.toFixed(2);
    document.getElementById('preview-cost-per-lb').textContent = '$' + costPerLb.toFixed(4) + '/lb';
}

function calcHoursFromTimeInputs(start, end) {
    const [sh, sm] = start.split(':').map(Number);
    const [eh, em] = end.split(':').map(Number);
    let diff = (eh * 60 + em) - (sh * 60 + sm);
    if (diff < 0) diff += 24 * 60;
    return diff / 60;
}

function saveEntry() {
    const activity = document.getElementById('entry-activity').value;
    const date = document.getElementById('entry-date').value;
    const supplier = document.getElementById('entry-supplier').value || null;
    const product = document.getElementById('entry-product').value;
    const lot = document.getElementById('entry-lot').value || null;
    const pallet = document.getElementById('entry-pallet').value || null;
    const incoming = parseFloat(document.getElementById('entry-incoming').value);
    const finished = parseFloat(document.getElementById('entry-finished').value);
    const people = parseInt(document.getElementById('entry-people').value);
    const startTime = document.getElementById('entry-time-start').value;
    const endTime = document.getElementById('entry-time-end').value;

    const hours = calcHoursFromTimeInputs(startTime, endTime);
    const totalLaborHours = people * hours;
    const laborCost = totalLaborHours * LABOR_RATE;
    const costPerLb = laborCost / finished;
    const yieldPct = activity === 'Stripping' ? null : (finished / incoming * 100);

    const dt = new Date(date);
    const iso = dt.getUTCDay() === 0
        ? getISOWeek(new Date(dt.getTime() - 86400000))
        : getISOWeek(dt);

    const record = {
        activity,
        date,
        week: iso,
        supplier,
        lot,
        pallet,
        product_format: product,
        incoming_lbs: Math.round(incoming * 100) / 100,
        finished_lbs: Math.round(finished * 100) / 100,
        yield_pct: yieldPct ? Math.round(yieldPct * 100) / 100 : null,
        people,
        hours_worked: Math.round(hours * 10000) / 10000,
        total_labor_hours: Math.round(totalLaborHours * 10000) / 10000,
        labor_cost: Math.round(laborCost * 100) / 100,
        cost_per_finished_lb: Math.round(costPerLb * 10000) / 10000,
        _manual: true,
        _entered_at: new Date().toISOString()
    };

    manualRecords.push(record);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(manualRecords));
    allData.records.push(record);

    const msg = document.getElementById('entry-success');
    msg.style.display = 'block';
    setTimeout(() => msg.style.display = 'none', 3000);

    document.getElementById('entry-form').reset();
    document.getElementById('entry-date').value = new Date().toISOString().slice(0, 10);
    document.getElementById('entry-preview').style.display = 'none';

    renderRecentEntries();
    applyFilters();
}

function renderRecentEntries() {
    const tbody = document.querySelector('#table-recent tbody');
    tbody.innerHTML = '';
    const recent = [...manualRecords].reverse().slice(0, 20);

    recent.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${formatDate(r.date)}</td>
            <td>${r.activity}</td>
            <td>${r.product_format}</td>
            <td class="text-right">${r.incoming_lbs?.toFixed(1)}</td>
            <td class="text-right">${r.finished_lbs?.toFixed(1)}</td>
            <td class="text-right">${r.yield_pct?.toFixed(1) || '--'}%</td>
            <td class="text-right">$${r.cost_per_finished_lb?.toFixed(4)}</td>
            <td><button class="btn-delete" data-idx="${manualRecords.length - 1 - idx}" title="Delete">&#x2715;</button></td>
        `;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll('.btn-delete').forEach(btn => {
        btn.addEventListener('click', () => {
            const i = parseInt(btn.dataset.idx);
            const removed = manualRecords.splice(i, 1)[0];
            localStorage.setItem(STORAGE_KEY, JSON.stringify(manualRecords));
            const ri = allData.records.findIndex(r => r === removed || (r._entered_at === removed._entered_at && r._manual));
            if (ri >= 0) allData.records.splice(ri, 1);
            renderRecentEntries();
            applyFilters();
        });
    });
}

// ---- HELPERS ----
function avg(arr) { return arr.reduce((s, v) => s + v, 0) / arr.length; }
function median(arr) {
    const s = [...arr].sort((a, b) => a - b);
    const m = Math.floor(s.length / 2);
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}
function percentile(sorted, p) {
    const i = Math.floor(sorted.length * p / 100);
    return sorted[Math.min(i, sorted.length - 1)];
}
function groupBy(arr, key) {
    return arr.reduce((g, item) => { (g[item[key]] = g[item[key]] || []).push(item); return g; }, {});
}
function formatDate(d) {
    if (!d) return '--';
    const parts = d.split('-');
    return parts[1] + '/' + parts[2];
}
function numberFmt(n) {
    return Number(n).toLocaleString('en-US');
}
function getISOWeek(d) {
    const date = new Date(d.getTime());
    date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
    return date.getUTCFullYear() + '-W' + String(weekNo).padStart(2, '0');
}
