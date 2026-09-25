geotab.addin.fuelMonitor = function (outerApi, outerState) {

    let currentApi = outerApi; 
    let chartInstance = null;
    let reportDataForExport = [];

    const modeRadios   = document.getElementsByName('searchMode');
    const entitySelect = document.getElementById('entitySelect');
    const dateFrom     = document.getElementById('dateFrom');
    const dateTo       = document.getElementById('dateTo');
    const btnFetch     = document.getElementById('btn-fetch-data');
    const btnExport    = document.getElementById('btn-export-excel');
    const panel        = document.getElementById('results-panel');
    const chartWrap    = document.getElementById('chartWrapper');

    // Helper para nombre y apellidos del conductor
    function getDriverDisplayName(user) {
        if (!user) return "Conductor Desconocido";
        
        let fullName = "";
        if (user.firstName || user.lastName) {
            fullName = `${user.firstName || ''} ${user.lastName || ''}`.trim();
        }
        
        if (!fullName) {
            fullName = user.name || user.id || "Sin Nombre";
        }

        // Limpiar identificadores tipo clave interna si existen como fallback único
        if (fullName.startsWith("#E_")) {
            fullName = user.name || "Conductor " + user.id;
        }

        return fullName;
    }

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

    function directMultiCall(callsArray, successCallback, errorCallback) {
        currentApi.getSession(function(credentials, server) {
            var url = 'https://' + (server || 'my.geotab.com') + '/apiv1';
            fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    method: "ExecuteMultiCall",
                    params: { calls: callsArray.map(call => ({ method: call[0], params: call[1] })), credentials: credentials }
                })
            })
            .then(res => res.json())
            .then(json => json.error ? (errorCallback && errorCallback(json.error)) : (successCallback && successCallback(json.result)))
            .catch(err => errorCallback && errorCallback(err));
        });
    }

    function loadEntityList() {
        const mode = document.querySelector('input[name="searchMode"]:checked').value;
        entitySelect.innerHTML = '<option value="ALL">Cargando...</option>';
        entitySelect.disabled = true;

        var searchParam = mode === 'User' ? { isDriver: true } : {};

        directCall('Get', { typeName: mode, search: searchParam }, function(entities) {
            entitySelect.disabled = false;
            entitySelect.innerHTML = '';

            var optAll = document.createElement('option');
            optAll.value = 'ALL';
            optAll.textContent = mode === 'Device' ? '-- TODOS LOS VEHÍCULOS --' : '-- TODOS LOS CONDUCTORES --';
            entitySelect.appendChild(optAll);

            var items = (entities || []).map(e => ({
                id: e.id,
                name: mode === 'User' ? getDriverDisplayName(e) : (e.name || e.id)
            }));

            items.sort((a, b) => a.name.localeCompare(b.name));

            items.forEach(item => {
                var opt = document.createElement('option');
                opt.value = item.id;
                opt.textContent = item.name;
                entitySelect.appendChild(opt);
            });
        }, function(error) {
            console.error('Error al cargar lista:', error);
            entitySelect.innerHTML = '<option value="ALL">-- Error al cargar --</option>';
            entitySelect.disabled = false;
        });
    }

    function loadReport() {
        if (!dateFrom.value || !dateTo.value) {
            panel.innerHTML = '<div class="card-result" style="color:#b91c1c;">Por favor selecciona las fechas "Desde" y "Hasta".</div>';
            return;
        }

        const mode = document.querySelector('input[name="searchMode"]:checked').value;
        const selectedId = entitySelect.value;
        const fromDate = new Date(dateFrom.value + "T00:00:00Z").toISOString();
        const toDate = new Date(dateTo.value + "T23:59:59Z").toISOString();

        panel.innerHTML = '<div class="card-result"><p>⏳ Obteniendo y procesando telemetría... Por favor espera.</p></div>';
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

    // ─── PROCESO: VEHÍCULOS ───────────────────────────────────────────────────
    function processAllVehicles(fromDate, toDate) {
        directCall('Get', { typeName: 'Device' }, function(devices) {
            if (!devices || devices.length === 0) {
                panel.innerHTML = '<div class="card-result"><p>No se encontraron vehículos en el sistema.</p></div>';
                return;
            }

            var calls = [];
            devices.forEach(function(dev) {
                calls.push(['Get', {
                    typeName: 'StatusData',
                    search: { deviceSearch: { id: dev.id }, diagnosticSearch: { id: 'DiagnosticOdometerAdjustmentId' }, fromDate: fromDate, toDate: toDate }
                }]);
                calls.push(['Get', {
                    typeName: 'StatusData',
                    search: { deviceSearch: { id: dev.id }, diagnosticSearch: { id: 'DiagnosticTotalFuelUsedId' }, fromDate: fromDate, toDate: toDate }
                }]);
            });

            directMultiCall(calls, function(results) {
                let tableHtml = `
                    <div class="card-result">
                        <h3>Resumen Global de Vehículos</h3>
                        <table>
                            <thead>
                                <tr>
                                    <th>Vehículo</th>
                                    <th>Distancia (Km)</th>
                                    <th>Combustible (L)</th>
                                    <th>Consumo Medio</th>
                                </tr>
                            </thead>
                            <tbody>`;

                const chartLabels = [], chartData = [];

                devices.forEach((dev, index) => {
                    const odoData = results[index * 2] || [];
                    const fuelData = results[index * 2 + 1] || [];

                    let distanceKm = 0;
                    let fuelLiters = 0;
                    let avg = 0;

                    if (odoData.length >= 2) {
                        distanceKm = (odoData[odoData.length - 1].data - odoData[0].data) / 1000;
                    }
                    if (fuelData.length >= 2) {
                        fuelLiters = fuelData[fuelData.length - 1].data - fuelData[0].data;
                    }

                    if (distanceKm > 0 && fuelLiters >= 0) {
                        avg = (fuelLiters / distanceKm) * 100;
                    }

                    if (distanceKm > 0 || fuelLiters > 0) {
                        const badgeClass = avg > 35 ? 'badge-high' : 'badge-good';
                        tableHtml += `
                            <tr>
                                <td><strong>${dev.name}</strong></td>
                                <td>${distanceKm.toFixed(2)} km</td>
                                <td>${fuelLiters.toFixed(2)} L</td>
                                <td><span class="badge-avg ${badgeClass}">${avg > 0 ? avg.toFixed(2) + ' L/100km' : 'N/D'}</span></td>
                            </tr>`;

                        reportDataForExport.push({
                            Vehiculo: dev.name,
                            Distancia_Km: distanceKm.toFixed(2),
                            Combustible_L: fuelLiters.toFixed(2),
                            Consumo_Medio: avg > 0 ? avg.toFixed(2) : 'N/D'
                        });

                        if (avg > 0) {
                            chartLabels.push(dev.name);
                            chartData.push(avg.toFixed(2));
                        }
                    }
                });

                tableHtml += `</tbody></table></div>`;

                if (reportDataForExport.length > 0) {
                    panel.innerHTML = tableHtml;
                    btnExport.style.display = 'inline-flex';
                    if (chartLabels.length > 0) {
                        renderBarChart(chartLabels, chartData, 'Consumo Medio (L/100km)');
                    }
                } else {
                    panel.innerHTML = '<div class="card-result"><p>No se encontraron datos de consumo u odómetro en el periodo seleccionado.</p></div>';
                }

            }, function(err) {
                console.error('Error MultiCall:', err);
                panel.innerHTML = '<div class="card-result" style="color:#b91c1c;"><p>Error de conexión al obtener telemetría de vehículos.</p></div>';
            });
        });
    }

    function processSingleVehicle(deviceId, deviceName, fromDate, toDate) {
        var calls = [
            ['Get', { typeName: 'StatusData', search: { deviceSearch: { id: deviceId }, diagnosticSearch: { id: 'DiagnosticOdometerAdjustmentId' }, fromDate: fromDate, toDate: toDate } }],
            ['Get', { typeName: 'StatusData', search: { deviceSearch: { id: deviceId }, diagnosticSearch: { id: 'DiagnosticTotalFuelUsedId' }, fromDate: fromDate, toDate: toDate } }]
        ];

        directMultiCall(calls, function(results) {
            const odoData = results[0] || [];
            const fuelData = results[1] || [];

            if (!odoData.length || !fuelData.length) {
                panel.innerHTML = '<div class="card-result"><p>No hay suficientes datos registrados para este vehículo en el rango de fechas seleccionadas.</p></div>';
                return;
            }

            const distanceKm = (odoData[odoData.length - 1].data - odoData[0].data) / 1000;
            const fuelLiters = fuelData[fuelData.length - 1].data - fuelData[0].data;
            const avg = distanceKm > 0 ? (fuelLiters / distanceKm) * 100 : 0;

            panel.innerHTML = `
                <div class="card-result">
                    <h3>Vehículo: ${deviceName}</h3>
                    <div style="display: flex; gap: 40px; margin-top: 10px;">
                        <div><span style="font-size:12px; color:#5a6a75; font-weight:700; text-transform:uppercase;">Distancia Recorrida</span><br><strong style="font-size:18px;">${distanceKm.toFixed(2)} km</strong></div>
                        <div><span style="font-size:12px; color:#5a6a75; font-weight:700; text-transform:uppercase;">Combustible Consumido</span><br><strong style="font-size:18px;">${fuelLiters.toFixed(2)} Litros</strong></div>
                        <div><span style="font-size:12px; color:#5a6a75; font-weight:700; text-transform:uppercase;">Consumo Medio</span><br><strong style="font-size:18px; color:${avg > 35 ? '#b91c1c' : '#15803d'};">${avg.toFixed(2)} L/100km</strong></div>
                    </div>
                </div>`;

            reportDataForExport.push({
                Vehiculo: deviceName,
                Distancia_Km: distanceKm.toFixed(2),
                Combustible_L: fuelLiters.toFixed(2),
                Consumo_Medio: avg.toFixed(2)
            });

            btnExport.style.display = 'inline-flex';
            renderLineChart(fuelData, 'Evolución de Consumo Acumulado (Litros)');
        });
    }

    // ─── PROCESO: CONDUCTORES (Con Litros y Consumo Medio) ─────────────────────
    function processAllDrivers(fromDate, toDate) {
        directCall('Get', { typeName: 'User', search: { isDriver: true } }, function(drivers) {
            directCall('Get', { typeName: 'Trip', search: { fromDate: fromDate, toDate: toDate } }, function(trips) {

                if (!trips || trips.length === 0) {
                    panel.innerHTML = '<div class="card-result"><p>No hay trayectos registrados en el periodo seleccionado.</p></div>';
                    return;
                }

                const driverMap = {};
                (drivers || []).forEach(d => {
                    driverMap[d.id] = {
                        name: getDriverDisplayName(d),
                        distance: 0,
                        tripsCount: 0,
                        devicesUsed: new Set()
                    };
                });

                // Agrupar viajes y dispositivos por conductor
                trips.forEach(t => {
                    var driverId = t.driver ? t.driver.id : null;
                    if (driverId && driverMap[driverId]) {
                        driverMap[driverId].distance += (t.distance || 0);
                        driverMap[driverId].tripsCount++;
                        if (t.device && t.device.id) {
                            driverMap[driverId].devicesUsed.add(t.device.id);
                        }
                    }
                });

                // Extraer dispositivos únicos implicados
                const uniqueDeviceIds = Array.from(new Set(trips.map(t => t.device ? t.device.id : null).filter(Boolean)));

                if (uniqueDeviceIds.length === 0) {
                    panel.innerHTML = '<div class="card-result"><p>No se encontraron datos de vehículos asociados a los conductores.</p></div>';
                    return;
                }

                // Consultar telemetría de combustible de los vehículos asociados
                var fuelCalls = uniqueDeviceIds.map(devId => ['Get', {
                    typeName: 'StatusData',
                    search: { deviceSearch: { id: devId }, diagnosticSearch: { id: 'DiagnosticTotalFuelUsedId' }, fromDate: fromDate, toDate: toDate }
                }]);

                directMultiCall(fuelCalls, function(fuelResults) {
                    const deviceFuelMap = {};
                    uniqueDeviceIds.forEach((devId, idx) => {
                        deviceFuelMap[devId] = fuelResults[idx] || [];
                    });

                    // Calcular combustible por conductor sumando el consumo durante sus trayectos
                    Object.keys(driverMap).forEach(dId => {
                        let totalFuel = 0;
                        const driverTrips = trips.filter(t => t.driver && t.driver.id === dId);

                        driverTrips.forEach(t => {
                            if (!t.device || !deviceFuelMap[t.device.id]) return;
                            const readings = deviceFuelMap[t.device.id];
                            if (readings.length < 2) return;

                            const tStart = new Date(t.start).getTime();
                            const tStop = new Date(t.stop).getTime();

                            let fuelStart = null, fuelStop = null;

                            for (let r of readings) {
                                const rTime = new Date(r.dateTime).getTime();
                                if (rTime >= tStart && fuelStart === null) fuelStart = r.data;
                                if (rTime <= tStop) fuelStop = r.data;
                            }

                            if (fuelStart !== null && fuelStop !== null && fuelStop >= fuelStart) {
                                totalFuel += (fuelStop - fuelStart);
                            }
                        });

                        driverMap[dId].fuel = totalFuel;
                    });

                    // Renderizado de tabla
                    let tableHtml = `
                        <div class="card-result">
                            <h3>Resumen Global de Conductores</h3>
                            <table>
                                <thead>
                                    <tr>
                                        <th>Conductor</th>
                                        <th>Nº Viajes</th>
                                        <th>Distancia (Km)</th>
                                        <th>Combustible (L)</th>
                                        <th>Consumo Medio</th>
                                    </tr>
                                </thead>
                                <tbody>`;

                    const chartLabels = [], chartData = [];

                    Object.keys(driverMap).forEach(id => {
                        const d = driverMap[id];
                        if (d.tripsCount > 0) {
                            const avg = d.distance > 0 ? (d.fuel / d.distance) * 100 : 0;
                            const badgeClass = avg > 35 ? 'badge-high' : 'badge-good';

                            tableHtml += `
                                <tr>
                                    <td><strong>${d.name}</strong></td>
                                    <td>${d.tripsCount}</td>
                                    <td>${d.distance.toFixed(2)} km</td>
                                    <td>${d.fuel > 0 ? d.fuel.toFixed(2) + ' L' : '0 L'}</td>
                                    <td><span class="badge-avg ${badgeClass}">${avg > 0 ? avg.toFixed(2) + ' L/100km' : 'N/D'}</span></td>
                                </tr>`;

                            reportDataForExport.push({
                                Conductor: d.name,
                                Viajes: d.tripsCount,
                                Distancia_Km: d.distance.toFixed(2),
                                Combustible_L: d.fuel.toFixed(2),
                                Consumo_Medio: avg > 0 ? avg.toFixed(2) : 'N/D'
                            });

                            if (avg > 0) {
                                chartLabels.push(d.name);
                                chartData.push(avg.toFixed(2));
                            }
                        }
                    });

                    tableHtml += `</tbody></table></div>`;

                    if (reportDataForExport.length > 0) {
                        panel.innerHTML = tableHtml;
                        btnExport.style.display = 'inline-flex';
                        if (chartLabels.length > 0) {
                            renderBarChart(chartLabels, chartData, 'Consumo Medio (L/100km) por Conductor');
                        }
                    } else {
                        panel.innerHTML = '<div class="card-result"><p>No se encontraron registros suficientes de trayectos o combustible para los conductores.</p></div>';
                    }
                });
            });
        });
    }

    function processSingleDriver(driverId, driverName, fromDate, toDate) {
        directCall('Get', {
            typeName: 'Trip',
            search: { userSearch: { id: driverId }, fromDate: fromDate, toDate: toDate }
        }, function(trips) {
            if (!trips || trips.length === 0) {
                panel.innerHTML = '<div class="card-result"><p>No hay viajes registrados para este conductor en las fechas elegidas.</p></div>';
                return;
            }

            let totalDistance = 0;
            const uniqueDeviceIds = new Set();

            trips.forEach(t => {
                totalDistance += (t.distance || 0);
                if (t.device && t.device.id) uniqueDeviceIds.add(t.device.id);
            });

            var fuelCalls = Array.from(uniqueDeviceIds).map(devId => ['Get', {
                typeName: 'StatusData',
                search: { deviceSearch: { id: devId }, diagnosticSearch: { id: 'DiagnosticTotalFuelUsedId' }, fromDate: fromDate, toDate: toDate }
            }]);

            directMultiCall(fuelCalls, function(fuelResults) {
                const deviceFuelMap = {};
                Array.from(uniqueDeviceIds).forEach((devId, idx) => {
                    deviceFuelMap[devId] = fuelResults[idx] || [];
                });

                let totalFuel = 0;
                trips.forEach(t => {
                    if (!t.device || !deviceFuelMap[t.device.id]) return;
                    const readings = deviceFuelMap[t.device.id];
                    if (readings.length < 2) return;

                    const tStart = new Date(t.start).getTime();
                    const tStop = new Date(t.stop).getTime();

                    let fuelStart = null, fuelStop = null;
                    for (let r of readings) {
                        const rTime = new Date(r.dateTime).getTime();
                        if (rTime >= tStart && fuelStart === null) fuelStart = r.data;
                        if (rTime <= tStop) fuelStop = r.data;
                    }

                    if (fuelStart !== null && fuelStop !== null && fuelStop >= fuelStart) {
                        totalFuel += (fuelStop - fuelStart);
                    }
                });

                const avg = totalDistance > 0 ? (totalFuel / totalDistance) * 100 : 0;

                panel.innerHTML = `
                    <div class="card-result">
                        <h3>Conductor: ${driverName}</h3>
                        <div style="display: flex; gap: 30px; margin-top: 10px;">
                            <div><span style="font-size:12px; color:#5a6a75; font-weight:700; text-transform:uppercase;">Viajes</span><br><strong style="font-size:18px;">${trips.length}</strong></div>
                            <div><span style="font-size:12px; color:#5a6a75; font-weight:700; text-transform:uppercase;">Distancia Total</span><br><strong style="font-size:18px;">${totalDistance.toFixed(2)} km</strong></div>
                            <div><span style="font-size:12px; color:#5a6a75; font-weight:700; text-transform:uppercase;">Combustible Consumido</span><br><strong style="font-size:18px;">${totalFuel.toFixed(2)} Litros</strong></div>
                            <div><span style="font-size:12px; color:#5a6a75; font-weight:700; text-transform:uppercase;">Consumo Medio</span><br><strong style="font-size:18px; color:${avg > 35 ? '#b91c1c' : '#15803d'};">${avg.toFixed(2)} L/100km</strong></div>
                        </div>
                    </div>`;

                reportDataForExport.push({
                    Conductor: driverName,
                    Viajes: trips.length,
                    Distancia_Km: totalDistance.toFixed(2),
                    Combustible_L: totalFuel.toFixed(2),
                    Consumo_Medio: avg.toFixed(2)
                });

                btnExport.style.display = 'inline-flex';
            });
        });
    }

    // ─── GRÁFICOS Y EXPORTACIÓN ───────────────────────────────────────────────
    function renderBarChart(labels, data, title) {
        chartWrap.style.display = 'block';
        const ctx = document.getElementById('fuelChart').getContext('2d');
        if (chartInstance) chartInstance.destroy();

        chartInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{ label: title, data: data, backgroundColor: '#2b537d', borderRadius: 4 }]
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
                datasets: [{ label: title, data: dataset, borderColor: '#2b537d', backgroundColor: 'rgba(43, 83, 125, 0.08)', fill: true }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                scales: { x: { type: 'time' } }
            }
        });
    }

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
        link.download = `Informe_Consumo_${dateFrom.value}_al_${dateTo.value}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    return {
        initialize: function (api, state, callback) {
            currentApi = api;

            const today = new Date();
            const lastWeek = new Date(today.getTime() - (7 * 24 * 60 * 60 * 1000));
            dateFrom.value = lastWeek.toISOString().split('T')[0];
            dateTo.value = today.toISOString().split('T')[0];

            modeRadios.forEach(r => r.addEventListener('change', loadEntityList));
            btnFetch.addEventListener('click', loadReport);
            btnExport.addEventListener('click', exportToCSV);

            callback();
        },
        focus: function (api, state) {
            loadEntityList();
        },
        blur: function () { }
    };
};