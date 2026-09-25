geotab.addin.fuelMonitor = function (outerApi, outerState) {

    let currentApi = outerApi; 
    let chartInstance = null;
    let reportDataForExport = [];

    // Referencias DOM
    const modeRadios  = document.getElementsByName('searchMode');
    const entitySelect= document.getElementById('entitySelect');
    const dateFrom    = document.getElementById('dateFrom');
    const dateTo      = document.getElementById('dateTo');
    const btnFetch    = document.getElementById('btn-fetch-data');
    const btnExport   = document.getElementById('btn-export-excel');
    const panel       = document.getElementById('results-panel');
    const chartWrap   = document.getElementById('chartWrapper');

    // ─── Direct API Calls (Bypassing Drive Proxy) ─────────────────────────────
    function directCall(method, params, successCallback, errorCallback) {
        currentApi.getSession(function(credentials, server) {
            var url = 'https://' + (server || 'my.geotab.com') + '/apiv1';
            fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ method: method, params: Object.assign({}, params, { credentials: credentials }) })
            })
            .then(res => res.json())
            .then(json => json.error ? (errorCallback && errorCallback(json.error)) : (successCallback && successCallback(json.result)))
            .catch(err => errorCallback && errorCallback(err));
        });
    }

    // ─── Carga Automática de Vehículos / Conductores ──────────────────────────
    function loadEntityList() {
        const mode = document.querySelector('input[name="searchMode"]:checked').value;
        entitySelect.innerHTML = '<option value="ALL">Cargando datos...</option>';
        entitySelect.disabled = true;

        var searchParam = mode === 'User' ? { isDriver: true } : {};

        directCall('Get', { typeName: mode, search: searchParam }, function(entities) {
            entitySelect.disabled = false;
            entitySelect.innerHTML = '';

            // Opción para ver TODOS
            var optAll = document.createElement('option');
            optAll.value = 'ALL';
            optAll.textContent = mode === 'Device' ? '-- TODOS LOS VEHÍCULOS --' : '-- TODOS LOS CONDUCTORES --';
            entitySelect.appendChild(optAll);

            // Ordenar alfabéticamente
            entities.sort((a, b) => (a.name || '').localeCompare(b.name || ''));

            entities.forEach(e => {
                if (e.name) {
                    var opt = document.createElement('option');
                    opt.value = e.id;
                    opt.textContent = e.name;
                    entitySelect.appendChild(opt);
                }
            });
        }, function(error) {
            console.error('Error al cargar lista:', error);
            entitySelect.innerHTML = '<option value="ALL">-- Error al cargar la lista --</option>';
            entitySelect.disabled = false;
        });
    }

    // ─── Carga de Informe (TODOS o INDIVIDUAL) ────────────────────────────────
    function loadReport() {
        if (!dateFrom.value || !dateTo.value) {
            panel.innerHTML = '<div class="card" style="color:red;">Por favor, selecciona las fechas Desde y Hasta.</div>';
            return;
        }

        const mode = document.querySelector('input[name="searchMode"]:checked').value;
        const selectedId = entitySelect.value;
        const fromDate = new Date(dateFrom.value + "T00:00:00Z").toISOString();
        const toDate = new Date(dateTo.value + "T23:59:59Z").toISOString();

        panel.innerHTML = '<div class="card"><p>⏳ Procesando datos de telemetría... Por favor espera.</p></div>';
        btnExport.style.display = 'none';
        chartWrap.style.display = 'none';
        reportDataForExport = [];

        if (mode === 'Device') {
            if (selectedId === 'ALL') {
                processAllVehicles(fromDate, toDate);
            } else {
                processSingleVehicle(selectedId, entitySelect.options[entitySelect.selectedIndex].text, fromDate, toDate);
            }
        } else {
            if (selectedId === 'ALL') {
                processAllDrivers(fromDate, toDate);
            } else {
                processSingleDriver(selectedId, entitySelect.options[entitySelect.selectedIndex].text, fromDate, toDate);
            }
        }
    }

    // ─── LÓGICA: TODOS LOS VEHÍCULOS ──────────────────────────────────────────
    function processAllVehicles(fromDate, toDate) {
        // Pedimos la lista de vehículos, odómetros y combustible global en paralelo
        directCall('Get', { typeName: 'Device' }, function(devices) {
            directCall('Get', {
                typeName: 'StatusData',
                search: { diagnosticSearch: { id: 'DiagnosticOdometerAdjustmentId' }, fromDate: fromDate, toDate: toDate }
            }, function(odoRecords) {
                directCall('Get', {
                    typeName: 'StatusData',
                    search: { diagnosticSearch: { id: 'DiagnosticTotalFuelUsedId' }, fromDate: fromDate, toDate: toDate }
                }, function(fuelRecords) {
                    
                    // Agrupar lecturas por ID de Vehículo
                    const odoMap = {}, fuelMap = {};
                    odoRecords.forEach(r => {
                        if (r.device && r.device.id) {
                            if (!odoMap[r.device.id]) odoMap[r.device.id] = [];
                            odoMap[r.device.id].push(r);
                        }
                    });
                    fuelRecords.forEach(r => {
                        if (r.device && r.device.id) {
                            if (!fuelMap[r.device.id]) fuelMap[r.device.id] = [];
                            fuelMap[r.device.id].push(r);
                        }
                    });

                    let tableHtml = `
                        <div class="card">
                            <h3>Resumen Global de Vehículos</h3>
                            <table>
                                <thead>
                                    <tr>
                                        <th>Vehículo</th>
                                        <th>Distancia (Km)</th>
                                        <th>Combustible (L)</th>
                                        <th>Consumo Medio (L/100km)</th>
                                    </tr>
                                </thead>
                                <tbody>`;

                    const chartLabels = [], chartData = [];

                    devices.forEach(dev => {
                        const devOdo = odoMap[dev.id] || [];
                        const devFuel = fuelMap[dev.id] || [];

                        if (devOdo.length >= 2 && devFuel.length >= 2) {
                            const distanceKm = (devOdo[devOdo.length - 1].data - devOdo[0].data) / 1000;
                            const fuelLiters = devFuel[devFuel.length - 1].data - devFuel[0].data;
                            const avg = distanceKm > 0 ? (fuelLiters / distanceKm) * 100 : 0;

                            if (distanceKm > 0 && fuelLiters >= 0) {
                                tableHtml += `
                                    <tr>
                                        <td><strong>${dev.name}</strong></td>
                                        <td>${distanceKm.toFixed(2)} km</td>
                                        <td>${fuelLiters.toFixed(2)} L</td>
                                        <td><strong style="color:${avg > 35 ? '#dc2626' : '#16a34a'}">${avg.toFixed(2)} L/100km</strong></td>
                                    </tr>`;

                                reportDataForExport.push({
                                    Vehiculo: dev.name,
                                    Distancia_Km: distanceKm.toFixed(2),
                                    Combustible_L: fuelLiters.toFixed(2),
                                    Consumo_Medio: avg.toFixed(2)
                                });

                                chartLabels.push(dev.name);
                                chartData.push(avg.toFixed(2));
                            }
                        }
                    });

                    tableHtml += `</tbody></table></div>`;
                    panel.innerHTML = tableHtml;

                    if (reportDataForExport.length > 0) {
                        btnExport.style.display = 'inline-block';
                        renderBarChart(chartLabels, chartData, 'Consumo Medio (L/100km) por Vehículo');
                    } else {
                        panel.innerHTML = '<div class="card"><p>No se encontraron datos suficientes en el periodo seleccionado.</p></div>';
                    }
                });
            });
        });
    }

    // ─── LÓGICA: VEHÍCULO INDIVIDUAL ──────────────────────────────────────────
    function processSingleVehicle(deviceId, deviceName, fromDate, toDate) {
        directCall('Get', {
            typeName: 'StatusData',
            search: { deviceSearch: { id: deviceId }, diagnosticSearch: { id: 'DiagnosticOdometerAdjustmentId' }, fromDate: fromDate, toDate: toDate }
        }, function(odoData) {
            directCall('Get', {
                typeName: 'StatusData',
                search: { deviceSearch: { id: deviceId }, diagnosticSearch: { id: 'DiagnosticTotalFuelUsedId' }, fromDate: fromDate, toDate: toDate }
            }, function(fuelData) {

                if (!odoData.length || !fuelData.length) {
                    panel.innerHTML = '<div class="card"><p>No hay datos de telemetría suficientes para este vehículo en las fechas seleccionadas.</p></div>';
                    return;
                }

                const distanceKm = (odoData[odoData.length - 1].data - odoData[0].data) / 1000;
                const fuelLiters = fuelData[fuelData.length - 1].data - fuelData[0].data;
                const avg = distanceKm > 0 ? (fuelLiters / distanceKm) * 100 : 0;

                panel.innerHTML = `
                    <div class="card">
                        <h3 style="color:#0284c7; margin-top:0;">Vehículo: ${deviceName}</h3>
                        <p><strong>Distancia Recorrida:</strong> ${distanceKm.toFixed(2)} km</p>
                        <p><strong>Combustible Consumido:</strong> ${fuelLiters.toFixed(2)} Litros</p>
                        <p><strong>Consumo Medio:</strong> <span style="font-size:18px; font-weight:bold; color:${avg > 35 ? '#dc2626' : '#16a34a'}">${avg.toFixed(2)} L/100km</span></p>
                    </div>`;

                reportDataForExport.push({
                    Vehiculo: deviceName,
                    Distancia_Km: distanceKm.toFixed(2),
                    Combustible_L: fuelLiters.toFixed(2),
                    Consumo_Medio: avg.toFixed(2)
                });

                btnExport.style.display = 'inline-block';
                renderLineChart(fuelData, 'Progresión de Consumo (Litros Acumulados)');
            });
        });
    }

    // ─── LÓGICA: TODOS LOS CONDUCTORES ────────────────────────────────────────
    function processAllDrivers(fromDate, toDate) {
        directCall('Get', { typeName: 'User', search: { isDriver: true } }, function(drivers) {
            directCall('Get', { typeName: 'Trip', search: { fromDate: fromDate, toDate: toDate } }, function(trips) {
                
                const driverStats = {};
                drivers.forEach(d => { driverStats[d.id] = { name: d.name, distance: 0, tripsCount: 0 }; });

                trips.forEach(t => {
                    var driverId = t.driver ? t.driver.id : null;
                    if (driverId && driverStats[driverId]) {
                        driverStats[driverId].distance += (t.distance || 0);
                        driverStats[driverId].tripsCount++;
                    }
                });

                let tableHtml = `
                    <div class="card">
                        <h3>Resumen Global de Conductores</h3>
                        <table>
                            <thead>
                                <tr>
                                    <th>Conductor</th>
                                    <th>Nº de Viajes</th>
                                    <th>Distancia Total (Km)</th>
                                </tr>
                            </thead>
                            <tbody>`;

                const chartLabels = [], chartData = [];

                Object.keys(driverStats).forEach(id => {
                    const d = driverStats[id];
                    if (d.tripsCount > 0) {
                        tableHtml += `
                            <tr>
                                <td><strong>${d.name}</strong></td>
                                <td>${d.tripsCount}</td>
                                <td>${d.distance.toFixed(2)} km</td>
                            </tr>`;

                        reportDataForExport.push({
                            Conductor: d.name,
                            Viajes: d.tripsCount,
                            Distancia_Km: d.distance.toFixed(2)
                        });

                        chartLabels.push(d.name);
                        chartData.push(d.distance.toFixed(2));
                    }
                });

                tableHtml += `</tbody></table></div>`;
                panel.innerHTML = tableHtml;

                if (reportDataForExport.length > 0) {
                    btnExport.style.display = 'inline-block';
                    renderBarChart(chartLabels, chartData, 'Distancia Recorrida (Km) por Conductor');
                } else {
                    panel.innerHTML = '<div class="card"><p>No se registraron viajes para los conductores en este periodo.</p></div>';
                }
            });
        });
    }

    // ─── LÓGICA: CONDUCTOR INDIVIDUAL ─────────────────────────────────────────
    function processSingleDriver(driverId, driverName, fromDate, toDate) {
        directCall('Get', {
            typeName: 'Trip',
            search: { userSearch: { id: driverId }, fromDate: fromDate, toDate: toDate }
        }, function(trips) {
            let totalDistance = 0;
            trips.forEach(t => { totalDistance += (t.distance || 0); });

            panel.innerHTML = `
                <div class="card">
                    <h3 style="color:#0284c7; margin-top:0;">Conductor: ${driverName}</h3>
                    <p><strong>Viajes Realizados:</strong> ${trips.length}</p>
                    <p><strong>Distancia Total Recorrida:</strong> ${totalDistance.toFixed(2)} km</p>
                </div>`;

            reportDataForExport.push({
                Conductor: driverName,
                Viajes: trips.length,
                Distancia_Km: totalDistance.toFixed(2)
            });

            btnExport.style.display = 'inline-block';
        });
    }

    // ─── GRÁFICOS (Chart.js) ──────────────────────────────────────────────────
    function renderBarChart(labels, data, title) {
        chartWrap.style.display = 'block';
        const ctx = document.getElementById('fuelChart').getContext('2d');
        if (chartInstance) chartInstance.destroy();

        chartInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: title,
                    data: data,
                    backgroundColor: '#2563eb',
                    borderRadius: 4
                }]
            },
            options: { responsive: true, maintainAspectRatio: false }
        });
    }

    function renderLineChart(fuelData, title) {
        chartWrap.style.display = 'block';
        const ctx = document.getElementById('fuelChart').getContext('2d');
        if (chartInstance) chartInstance.destroy();

        const dataset = fuelData.map(d => ({ x: new Date(d.dateTime), y: d.data }));

        chartInstance = new Chart(ctx, {
            type: 'line',
            data: {
                datasets: [{
                    label: title,
                    data: dataset,
                    borderColor: '#2563eb',
                    backgroundColor: 'rgba(37, 99, 235, 0.1)',
                    fill: true
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                scales: { x: { type: 'time' } }
            }
        });
    }

    // ─── Exportar a Excel (CSV con formato UTF-8) ─────────────────────────────
    function exportToCSV() {
        if (!reportDataForExport.length) return;
        const headers = Object.keys(reportDataForExport[0]);
        const csvRows = [headers.join(';')];

        for (const row of reportDataForExport) {
            csvRows.push(headers.map(h => row[h]).join(';'));
        }

        const blob = new Blob(["\uFEFF" + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `Informe_Combustible_${dateFrom.value}_al_${dateTo.value}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // ─── Ciclo de Vida del Add-in ─────────────────────────────────────────────
    return {
        initialize: function (api, state, callback) {
            currentApi = api;

            const today = new Date().toISOString().split('T')[0];
            dateFrom.value = today;
            dateTo.value = today;

            // Event Listeners
            modeRadios.forEach(r => r.addEventListener('change', loadEntityList));
            btnFetch.addEventListener('click', loadReport);
            btnExport.addEventListener('click', exportToCSV);

            callback();
        },
        focus: function (api, state) {
            // Cargar automáticamente la lista al entrar a la pantalla
            loadEntityList();
        },
        blur: function () { }
    };
};