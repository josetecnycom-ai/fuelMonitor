/**
 * Add-in de Monitorización de Consumo de Combustible para MyGeotab
 * Archivo: app.js
 */

geotab.addin.fuelMonitor = function (api, state) {
    let chartInstance = null;
    let currentReportData = [];

    // Identificador del diagnóstico de combustible acumulado en Geotab
    const DIAGNOSTICS = {
        FUEL_USED: "DiagnosticTotalFuelUsedId"
    };

    // Umbral mínimo de distancia para procesar trayectos
    const MIN_TRIP_DISTANCE_KM = 0;

    return {
        initialize: function (api, state, callback) {
            initDefaultDates();
            bindEventListeners(api);
            populateEntitySelect(api, "Device");
            callback();
        },
        focus: function (api, state) {},
        blur: function (api, state) {}
    };

    // =========================================================================
    // 1. INICIALIZACIÓN Y EVENTOS
    // =========================================================================

    function initDefaultDates() {
        const today = new Date();
        const firstDayOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

        document.getElementById("dateFrom").value = firstDayOfMonth.toISOString().split("T")[0];
        document.getElementById("dateTo").value = today.toISOString().split("T")[0];
    }

    function bindEventListeners(api) {
        document.querySelectorAll('input[name="searchMode"]').forEach(radio => {
            radio.addEventListener("change", (e) => {
                populateEntitySelect(api, e.target.value);
            });
        });

        document.getElementById("btn-fetch-data").addEventListener("click", () => {
            generateReport(api);
        });

        document.getElementById("btn-export-excel").addEventListener("click", () => {
            exportToCSV();
        });
    }

    function populateEntitySelect(api, mode) {
        const select = document.getElementById("entitySelect");
        select.innerHTML = '<option value="ALL">-- TODOS --</option>';

        if (mode === "Device") {
            api.call("Get", { typeName: "Device" }, function (devices) {
                devices.sort((a, b) => a.name.localeCompare(b.name));
                devices.forEach(d => {
                    const opt = document.createElement("option");
                    opt.value = d.id;
                    opt.textContent = d.name;
                    select.appendChild(opt);
                });
            }, showError);
        } else {
            api.call("Get", { typeName: "User", search: { isDriver: true } }, function (drivers) {
                drivers.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
                drivers.forEach(u => {
                    const opt = document.createElement("option");
                    opt.value = u.id;
                    opt.textContent = `${u.firstName || ''} ${u.lastName || ''}`.trim() || u.name;
                    select.appendChild(opt);
                });
            }, showError);
        }
    }

    // =========================================================================
    // 2. ORQUESTACIÓN DE INFORMES
    // =========================================================================

    function generateReport(api) {
        const mode = document.querySelector('input[name="searchMode"]:checked').value;
        const selectedEntityId = document.getElementById("entitySelect").value;
        const dateFromVal = document.getElementById("dateFrom").value;
        const dateToVal = document.getElementById("dateTo").value;

        if (!dateFromVal || !dateToVal) {
            alert("Por favor, selecciona un rango de fechas válido.");
            return;
        }

        const fromDate = new Date(dateFromVal + "T00:00:00.000Z").toISOString();
        const toDate = new Date(dateToVal + "T23:59:59.999Z").toISOString();

        showLoading();

        if (mode === "Device") {
            processVehiclesReport(api, selectedEntityId, fromDate, toDate);
        } else {
            processDriversReport(api, selectedEntityId, fromDate, toDate);
        }
    }

    // =========================================================================
    // 3. CÁLCULO OPTIMIZADO POR VEHÍCULO
    // =========================================================================

    function processVehiclesReport(api, selectedId, fromDate, toDate) {
        const deviceSearch = selectedId === "ALL" ? {} : { id: selectedId };

        const expandedFrom = new Date(new Date(fromDate).getTime() - 24 * 60 * 60 * 1000).toISOString();
        const expandedTo = new Date(new Date(toDate).getTime() + 24 * 60 * 60 * 1000).toISOString();

        const fromDateTs = new Date(fromDate).getTime();
        const toDateTs = new Date(toDate).getTime();

        api.call("Get", { typeName: "Device", search: deviceSearch }, function (devices) {
            if (!devices || devices.length === 0) {
                renderEmptyResults("No se encontraron vehículos.");
                return;
            }

            const tripSearch = { fromDate: fromDate, toDate: toDate };
            if (selectedId !== "ALL") tripSearch.deviceId = selectedId;

            api.call("Get", { typeName: "Trip", search: tripSearch }, function (trips) {
                const deviceDistanceMap = {};
                (trips || []).forEach(t => {
                    const devId = t.device.id;
                    if (!deviceDistanceMap[devId]) deviceDistanceMap[devId] = 0;
                    deviceDistanceMap[devId] += (t.distance || 0);
                });

                let calls = devices.map(d => ["Get", {
                    typeName: "StatusData",
                    search: {
                        deviceId: d.id,
                        diagnosticSearch: { id: DIAGNOSTICS.FUEL_USED },
                        fromDate: expandedFrom,
                        toDate: expandedTo
                    }
                }]);

                api.multiCall(calls, function (fuelResults) {
                    currentReportData = [];

                    for (let i = 0; i < devices.length; i++) {
                        const device = devices[i];
                        const distKm = deviceDistanceMap[device.id] || 0;
                        const cleanFuelList = prepareFuelData(fuelResults[i] || []);

                        const startFuel = getInterpolatedFuelValue(cleanFuelList, fromDateTs);
                        const endFuel = getInterpolatedFuelValue(cleanFuelList, toDateTs);

                        let fuelLiters = 0;
                        let hasCanBus = false;

                        if (startFuel !== null && endFuel !== null && endFuel >= startFuel) {
                            fuelLiters = endFuel - startFuel;
                            hasCanBus = true;
                        }

                        let avgConsumption = (distKm > 0 && hasCanBus) ? (fuelLiters / distKm) * 100.0 : 0;

                        currentReportData.push({
                            id: device.id,
                            name: device.name,
                            distanceKm: parseFloat(distKm.toFixed(2)),
                            fuelLiters: parseFloat(fuelLiters.toFixed(2)),
                            avgConsumption: parseFloat(avgConsumption.toFixed(2)),
                            hasCanBus: hasCanBus,
                            tripsCount: "N/A"
                        });
                    }

                    renderTable("Device");
                    renderChart();
                }, showError);

            }, showError);

        }, showError);
    }

    // =========================================================================
    // 4. CÁLCULO OPTIMIZADO POR CONDUCTOR (DRIVERCHANGE + PREPROCESAMIENTO)
    // =========================================================================

    function processDriversReport(api, selectedId, fromDate, toDate) {
        const expandedFrom = new Date(new Date(fromDate).getTime() - 24 * 60 * 60 * 1000).toISOString();
        const expandedTo = new Date(new Date(toDate).getTime() + 24 * 60 * 60 * 1000).toISOString();

        const multiCallRequests = [
            ["Get", { typeName: "Trip", search: { fromDate: fromDate, toDate: toDate } }],
            ["Get", { typeName: "DriverChange", search: { fromDate: expandedFrom, toDate: expandedTo } }],
            ["Get", { typeName: "User", search: { isDriver: true } }]
        ];

        api.multiCall(multiCallRequests, function (results) {
            const rawTrips = results[0] || [];
            const rawDriverChanges = results[1] || [];
            const rawUsers = results[2] || [];

            if (!rawTrips || rawTrips.length === 0) {
                renderEmptyResults("No se registraron viajes para el período seleccionado.");
                return;
            }

            // 1. Mapa de nombres de usuario
            const userMap = {};
            rawUsers.forEach(u => {
                const fullName = `${u.firstName || ''} ${u.lastName || ''}`.trim() || u.name;
                userMap[u.id] = fullName;
            });

            // 2. Preprocesar y ordenar DriverChange por timestamps numéricos
            const driverChangesByDevice = {};
            rawDriverChanges.forEach(dc => {
                if (dc.device && dc.device.id) {
                    const devId = dc.device.id;
                    if (!driverChangesByDevice[devId]) driverChangesByDevice[devId] = [];
                    driverChangesByDevice[devId].push({
                        driverId: dc.driver ? dc.driver.id : null,
                        _ts: new Date(dc.dateTime).getTime()
                    });
                }
            });

            Object.keys(driverChangesByDevice).forEach(devId => {
                driverChangesByDevice[devId].sort((a, b) => a._ts - b._ts);
            });

            // 3. Preprocesar viajes con timestamps numéricos
            const validTrips = [];
            rawTrips.forEach(t => {
                const dist = t.distance || 0;
                if (dist >= MIN_TRIP_DISTANCE_KM) {
                    validTrips.push({
                        raw: t,
                        deviceId: t.device ? t.device.id : null,
                        distance: dist,
                        startTs: new Date(t.start).getTime(),
                        stopTs: new Date(t.stop).getTime(),
                        driverId: t.driver ? t.driver.id : null
                    });
                }
            });

            if (validTrips.length === 0) {
                renderEmptyResults("No hay viajes válidos en el rango seleccionado.");
                return;
            }

            // Asignación rápida de conductor por viaje
            function resolveDriverForTrip(trip) {
                const devId = trip.deviceId;
                const tStart = trip.startTs;

                if (devId && driverChangesByDevice[devId]) {
                    const changes = driverChangesByDevice[devId];
                    let matchedDriverId = null;

                    for (let i = 0; i < changes.length; i++) {
                        if (changes[i]._ts <= tStart) {
                            matchedDriverId = changes[i].driverId;
                        } else {
                            break;
                        }
                    }

                    if (matchedDriverId && matchedDriverId !== "NoDriverId" && matchedDriverId !== "UnknownDriverId") {
                        return matchedDriverId;
                    }
                }

                if (trip.driverId && trip.driverId !== "NoDriverId" && trip.driverId !== "UnknownDriverId") {
                    return trip.driverId;
                }

                return "UNKNOWN";
            }

            // 4. Agrupar viajes por conductor y registrar activos
            const driverMap = {};
            const activeDeviceIds = new Set();

            validTrips.forEach(t => {
                const resolvedDriverId = resolveDriverForTrip(t);

                if (selectedId !== "ALL" && resolvedDriverId !== selectedId) {
                    return;
                }

                const driverName = resolvedDriverId === "UNKNOWN" 
                    ? "Conductor No Identificado" 
                    : (userMap[resolvedDriverId] || resolvedDriverId);

                if (!driverMap[resolvedDriverId]) {
                    driverMap[resolvedDriverId] = {
                        id: resolvedDriverId,
                        name: driverName,
                        trips: [],
                        totalDistanceKm: 0
                    };
                }

                driverMap[resolvedDriverId].trips.push(t);
                driverMap[resolvedDriverId].totalDistanceKm += t.distance;

                if (t.deviceId) {
                    activeDeviceIds.add(t.deviceId);
                }
            });

            const driverKeys = Object.keys(driverMap);
            if (driverKeys.length === 0) {
                renderEmptyResults("No existen datos de viajes atribuidos al conductor seleccionado.");
                return;
            }

            // 5. Consultar telemetría únicamente para los vehículos involucrados
            const deviceArray = Array.from(activeDeviceIds);
            let fuelCalls = deviceArray.map(devId => ["Get", {
                typeName: "StatusData",
                search: {
                    deviceId: devId,
                    diagnosticSearch: { id: DIAGNOSTICS.FUEL_USED },
                    fromDate: expandedFrom,
                    toDate: expandedTo
                }
            }]);

            api.multiCall(fuelCalls, function (fuelResults) {
                const deviceFuelDataMap = {};
                deviceArray.forEach((devId, idx) => {
                    deviceFuelDataMap[devId] = prepareFuelData(fuelResults[idx] || []);
                });

                currentReportData = [];

                driverKeys.forEach(dKey => {
                    const driverGroup = driverMap[dKey];
                    let totalFuelLiters = 0;
                    let validTripsWithCan = 0;

                    driverGroup.trips.forEach(trip => {
                        const cleanFuelList = deviceFuelDataMap[trip.deviceId] || [];
                        if (cleanFuelList.length >= 2) {
                            const tripFuel = getFuelForTimeRange(cleanFuelList, trip.startTs, trip.stopTs);
                            if (tripFuel > 0) {
                                totalFuelLiters += tripFuel;
                                validTripsWithCan++;
                            }
                        }
                    });

                    const dist = driverGroup.totalDistanceKm;
                    const avg = dist > 0 ? (totalFuelLiters / dist) * 100.0 : 0;

                    currentReportData.push({
                        id: driverGroup.id,
                        name: driverGroup.name,
                        distanceKm: parseFloat(dist.toFixed(2)),
                        fuelLiters: parseFloat(totalFuelLiters.toFixed(2)),
                        avgConsumption: parseFloat(avg.toFixed(2)),
                        hasCanBus: validTripsWithCan > 0 || totalFuelLiters > 0,
                        tripsCount: driverGroup.trips.length
                    });
                });

                renderTable("User");
                renderChart();
            }, showError);

        }, showError);
    }

    // =========================================================================
    // 5. INTERPOLACIÓN Y CÁLCULO MEDIANTE BÚSQUEDA BINARIA O(log N)
    // =========================================================================

    function prepareFuelData(rawList) {
        const cleanList = [];
        if (!rawList || !Array.isArray(rawList)) return cleanList;

        for (let i = 0; i < rawList.length; i++) {
            const pt = rawList[i];
            if (pt && pt.dateTime && typeof pt.data === 'number') {
                cleanList.push({
                    _ts: new Date(pt.dateTime).getTime(),
                    data: pt.data
                });
            }
        }
        cleanList.sort((a, b) => a._ts - b._ts);
        return cleanList;
    }

    function getInterpolatedFuelValue(cleanFuelList, targetTs) {
        if (!cleanFuelList || cleanFuelList.length === 0) return null;

        let low = 0;
        let high = cleanFuelList.length - 1;
        let before = null;
        let after = null;

        // Búsqueda binaria para 'before'
        while (low <= high) {
            const mid = (low + high) >> 1;
            if (cleanFuelList[mid]._ts <= targetTs) {
                before = cleanFuelList[mid];
                low = mid + 1;
            } else {
                high = mid - 1;
            }
        }

        // Búsqueda binaria para 'after'
        low = 0;
        high = cleanFuelList.length - 1;
        while (low <= high) {
            const mid = (low + high) >> 1;
            if (cleanFuelList[mid]._ts >= targetTs) {
                after = cleanFuelList[mid];
                high = mid - 1;
            } else {
                low = mid + 1;
            }
        }

        if (!before || !after) return null;

        const t1 = before._ts;
        const t2 = after._ts;

        if (t1 === t2) return before.data;

        const ratio = (targetTs - t1) / (t2 - t1);
        return before.data + (after.data - before.data) * ratio;
    }

    function getFuelForTimeRange(cleanFuelList, startTs, stopTs) {
        const fuelStart = getInterpolatedFuelValue(cleanFuelList, startTs);
        const fuelStop = getInterpolatedFuelValue(cleanFuelList, stopTs);

        if (fuelStart === null || fuelStop === null) {
            return 0;
        }

        const diff = fuelStop - fuelStart;
        return diff > 0 ? diff : 0;
    }

    // =========================================================================
    // 6. RENDERIZADO (TABLA Y GRÁFICO)
    // =========================================================================

    function renderTable(mode) {
        const threshold = parseFloat(document.getElementById("thresholdLimit").value) || 30.0;
        const panel = document.getElementById("results-panel");
        document.getElementById("btn-export-excel").style.display = "inline-block";

        if (currentReportData.length === 0) {
            renderEmptyResults("No se obtuvieron registros.");
            return;
        }

        let html = `
            <table class="fuel-table">
                <thead>
                    <tr>
                        <th>${mode === "Device" ? "Vehículo / Activo" : "Conductor"}</th>
                        ${mode === "User" ? "<th>Viajes Analizados</th>" : ""}
                        <th style="text-align: right;">Distancia (km)</th>
                        <th style="text-align: right;">Combustible (L)</th>
                        <th style="text-align: right;">Consumo Medio (L/100km)</th>
                        <th style="text-align: center;">Estado CAN / Alerta</th>
                    </tr>
                </thead>
                <tbody>
        `;

        currentReportData.forEach(row => {
            const isAlert = row.avgConsumption > threshold;
            const rowStyle = isAlert ? 'style="background-color: #fee2e2;"' : '';
            const textAlertStyle = isAlert ? 'color: #b91c1c; font-weight: 700;' : '';

            let statusBadge = '';
            if (!row.hasCanBus) {
                statusBadge = '<span style="color: #d97706; font-weight: 600;">Sin Datos CAN</span>';
            } else if (isAlert) {
                statusBadge = `<span style="color: #b91c1c; font-weight: 700;">Excede Umbral (&gt;${threshold})</span>`;
            } else {
                statusBadge = '<span style="color: #16a34a; font-weight: 600;">Normal</span>';
            }

            html += `
                <tr ${rowStyle}>
                    <td style="font-weight: 600;">${escapeHtml(row.name)}</td>
                    ${mode === "User" ? `<td>${row.tripsCount}</td>` : ""}
                    <td style="text-align: right;">${row.distanceKm.toLocaleString('es-ES')}</td>
                    <td style="text-align: right;">${row.fuelLiters.toLocaleString('es-ES')}</td>
                    <td style="text-align: right; ${textAlertStyle}">${row.avgConsumption.toLocaleString('es-ES')}</td>
                    <td style="text-align: center;">${statusBadge}</td>
                </tr>
            `;
        });

        html += `</tbody></table>`;
        panel.innerHTML = html;
    }

    function renderChart() {
        const chartWrapper = document.getElementById("chartWrapper");
        const ctx = document.getElementById("fuelChart").getContext("2d");
        const threshold = parseFloat(document.getElementById("thresholdLimit").value) || 30.0;

        chartWrapper.style.display = "block";

        if (chartInstance) {
            chartInstance.destroy();
        }

        const labels = currentReportData.map(d => d.name);
        const dataValues = currentReportData.map(d => d.avgConsumption);
        const backgroundColors = dataValues.map(v => v > threshold ? 'rgba(220, 38, 38, 0.75)' : 'rgba(43, 83, 125, 0.75)');
        const borderColors = dataValues.map(v => v > threshold ? '#b91c1c' : '#1f3e5e');

        chartInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Consumo Medio (L/100km)',
                    data: dataValues,
                    backgroundColor: backgroundColors,
                    borderColor: borderColors,
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: true, position: 'top' },
                    tooltip: {
                        callbacks: {
                            label: function (context) {
                                return ` Consumo: ${context.parsed.y} L/100km`;
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: { display: true, text: 'L/100km' }
                    },
                    x: {
                        ticks: { maxRotation: 45, minRotation: 0 }
                    }
                }
            }
        });
    }

    // =========================================================================
    // 7. EXPORTACIÓN CSV (UTF-8 BOM)
    // =========================================================================

    function exportToCSV() {
        if (currentReportData.length === 0) return;

        const mode = document.querySelector('input[name="searchMode"]:checked').value;
        const threshold = parseFloat(document.getElementById("thresholdLimit").value) || 30.0;

        let csvContent = "\uFEFF";
        csvContent += mode === "Device" 
            ? "Vehículo;Distancia (km);Combustible (L);Consumo Medio (L/100km);Estado\n"
            : "Conductor;Viajes Analizados;Distancia (km);Combustible (L);Consumo Medio (L/100km);Estado\n";

        currentReportData.forEach(row => {
            const status = !row.hasCanBus ? "Sin Datos CAN" : (row.avgConsumption > threshold ? "Excede Umbral" : "Normal");
            const col2 = mode === "User" ? row.tripsCount : "";
            
            const distStr = row.distanceKm.toString().replace('.', ',');
            const fuelStr = row.fuelLiters.toString().replace('.', ',');
            const avgStr = row.avgConsumption.toString().replace('.', ',');

            csvContent += `"${row.name.replace(/"/g, '""')}";"${col2}";${distStr};${fuelStr};${avgStr};"${status}"\n`;
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", `Informe_Consumo_${mode}_${new Date().toISOString().split('T')[0]}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // =========================================================================
    // 8. UTILIDADES
    // =========================================================================

    function showLoading() {
        document.getElementById("btn-export-excel").style.display = "none";
        document.getElementById("chartWrapper").style.display = "none";
        document.getElementById("results-panel").innerHTML = `
            <div style="padding: 40px; text-align: center; color: #2b537d; font-weight: 600;">
                <svg width="32" height="32" viewBox="0 0 24 24" style="animation: spin 1s linear infinite;" fill="none" stroke="currentColor" stroke-width="2">
                    <circle cx="12" cy="12" r="10" stroke-opacity="0.25"></circle>
                    <path d="M12 2 a10 10 0 0 1 10 10" stroke-linecap="round"></path>
                </svg>
                <style>@keyframes spin { 100% { transform: rotate(360deg); } }</style>
                <p style="margin-top: 12px; font-size: 14px;">Consultando telemetría y procesando datos CAN-bus...</p>
            </div>
        `;
    }

    function renderEmptyResults(msg) {
        document.getElementById("results-panel").innerHTML = `
            <div style="padding: 30px; text-align: center; color: #5a6a75; background: #ffffff; border: 1px solid #dce2e6; border-radius: 6px;">
                <p style="margin: 0; font-size: 14px; font-weight: 500;">${escapeHtml(msg)}</p>
            </div>
        `;
        document.getElementById("chartWrapper").style.display = "none";
        document.getElementById("btn-export-excel").style.display = "none";
    }

    function showError(error) {
        console.error("Geotab Fuel Monitor Error:", error);
        document.getElementById("results-panel").innerHTML = `
            <div style="padding: 16px; background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; color: #991b1b; font-size: 13px;">
                <strong>Error al consultar la API de Geotab:</strong> ${escapeHtml(error.message || JSON.stringify(error))}
            </div>
        `;
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }
};
