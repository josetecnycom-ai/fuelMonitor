/**
 * Add-in de Monitorización de Consumo de Combustible para MyGeotab
 * Archivo: app.js (Optimizado para Cuotas de API y Telemetría CAN-bus)
 */

geotab.addin.fuelMonitor = function (api, state) {
    let chartInstance = null;
    let currentReportData = [];

    // Identificador del diagnóstico de combustible acumulado en Geotab
    const DIAGNOSTICS = {
        FUEL_USED: "DiagnosticTotalFuelUsedId"
    };

    // Umbral mínimo de distancia para filtrar micro-movimientos en conductores
    const MIN_TRIP_DISTANCE_KM = 0.5;

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

        api.call("Get", { typeName: "Device", search: deviceSearch }, function (devices) {
            if (!devices || devices.length === 0) {
                renderEmptyResults("No se encontraron vehículos.");
                return;
            }

            // Petición de Trips para obtener la distancia real exacta del periodo
            const tripSearch = { fromDate: fromDate, toDate: toDate };
            if (selectedId !== "ALL") tripSearch.deviceId = selectedId;

            api.call("Get", { typeName: "Trip", search: tripSearch }, function (trips) {
                const deviceDistanceMap = {};
                (trips || []).forEach(t => {
                    const devId = t.device.id;
                    if (!deviceDistanceMap[devId]) deviceDistanceMap[devId] = 0;
                    deviceDistanceMap[devId] += (t.distance || 0);
                });

                // Consulta de combustible acumulado (1 llamada por vehículo)
                let calls = devices.map(d => ["Get", {
                    typeName: "StatusData",
                    search: {
                        deviceId: d.id,
                        diagnosticSearch: { id: DIAGNOSTICS.FUEL_USED },
                        fromDate: fromDate,
                        toDate: toDate
                    }
                }]);

                api.multiCall(calls, function (fuelResults) {
                    currentReportData = [];

                    for (let i = 0; i < devices.length; i++) {
                        const device = devices[i];
                        const distKm = deviceDistanceMap[device.id] || 0;
                        const fuelData = fuelResults[i] || [];

                        let fuelLiters = 0;
                        let hasCanBus = false;

                        if (fuelData.length >= 2) {
                            const startFuel = fuelData[0].data;
                            const endFuel = fuelData[fuelData.length - 1].data;
                            const diff = endFuel - startFuel;
                            if (diff >= 0) {
                                fuelLiters = diff;
                                hasCanBus = true;
                            }
                        }

                        let avgConsumption = (distKm > 0 && hasCanBus) ? (fuelLiters / distKm) * 100.0 : 0;

                        currentReportData.push({
                            id: device.id,
                            name: device.name,
                            serialNumber: device.serialNumber || "N/D",
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
    // 4. CÁLCULO OPTIMIZADO POR CONDUCTOR (SIN EXCEDER CUOTA)
    // =========================================================================

    function processDriversReport(api, selectedId, fromDate, toDate) {
        const tripSearch = { fromDate: fromDate, toDate: toDate };
        if (selectedId !== "ALL") {
            tripSearch.driverSearch = { id: selectedId };
        }

        // Obtener todos los viajes del rango en UNA sola llamada
        api.call("Get", { typeName: "Trip", search: tripSearch }, function (trips) {
            if (!trips || trips.length === 0) {
                renderEmptyResults("No se registraron viajes para los criterios seleccionados.");
                return;
            }

            const validTrips = trips.filter(t => (t.distance || 0) >= MIN_TRIP_DISTANCE_KM);

            if (validTrips.length === 0) {
                renderEmptyResults(`No hay viajes registrados superiores a ${MIN_TRIP_DISTANCE_KM} km.`);
                return;
            }

            // Agrupar por conductor y recopilar vehículos involucrados
            const driverMap = {};
            const activeDeviceIds = new Set();

            validTrips.forEach(t => {
                const driverId = (t.driver && t.driver.id !== "NoDriverId") ? t.driver.id : "UNKNOWN";
                const driverName = (t.driver && (t.driver.name || t.driver.firstName)) 
                    ? `${t.driver.firstName || ''} ${t.driver.lastName || ''}`.trim() || t.driver.name
                    : "Sin Identificar";

                if (!driverMap[driverId]) {
                    driverMap[driverId] = {
                        id: driverId,
                        name: driverName,
                        trips: [],
                        totalDistanceKm: 0
                    };
                }
                driverMap[driverId].trips.push(t);
                driverMap[driverId].totalDistanceKm += (t.distance || 0);

                if (t.device && t.device.id) {
                    activeDeviceIds.add(t.device.id);
                }
            });

            // Pedir la telemetría solo para los vehículos involucrados (1 llamada por vehículo activo)
            const deviceArray = Array.from(activeDeviceIds);
            let fuelCalls = deviceArray.map(devId => ["Get", {
                typeName: "StatusData",
                search: {
                    deviceId: devId,
                    diagnosticSearch: { id: DIAGNOSTICS.FUEL_USED },
                    fromDate: fromDate,
                    toDate: toDate
                }
            }]);

            api.multiCall(fuelCalls, function (fuelResults) {
                // Mapa de datos de combustible por vehículo
                const deviceFuelDataMap = {};
                deviceArray.forEach((devId, idx) => {
                    deviceFuelDataMap[devId] = fuelResults[idx] || [];
                });

                currentReportData = [];

                // Correlación local en JS entre viaje y telemetría de combustible
                Object.keys(driverMap).forEach(dKey => {
                    const driverGroup = driverMap[dKey];
                    let totalFuelLiters = 0;
                    let validTripsWithCan = 0;

                    driverGroup.trips.forEach(trip => {
                        const devFuelList = deviceFuelDataMap[trip.device.id] || [];
                        if (devFuelList.length >= 2) {
                            const tripFuel = getFuelForTimeRange(devFuelList, trip.start, trip.stop);
                            if (tripFuel >= 0) {
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
                        serialNumber: "N/A",
                        distanceKm: parseFloat(dist.toFixed(2)),
                        fuelLiters: parseFloat(totalFuelLiters.toFixed(2)),
                        avgConsumption: parseFloat(avg.toFixed(2)),
                        hasCanBus: validTripsWithCan > 0,
                        tripsCount: driverGroup.trips.length
                    });
                });

                renderTable("User");
                renderChart();
            }, showError);

        }, showError);
    }

    // Calcula el combustible gastado en la ventana temporal de un viaje
    function getFuelForTimeRange(fuelDataList, startTimeStr, stopTimeStr) {
        const tStart = new Date(startTimeStr).getTime();
        const tStop = new Date(stopTimeStr).getTime();

        const pointsInTrip = fuelDataList.filter(p => {
            const ptTime = new Date(p.dateTime).getTime();
            return ptTime >= tStart && ptTime <= tStop;
        });

        if (pointsInTrip.length >= 2) {
            return pointsInTrip[pointsInTrip.length - 1].data - pointsInTrip[0].data;
        } else if (fuelDataList.length >= 2) {
            const startPt = fuelDataList.find(p => new Date(p.dateTime).getTime() >= tStart) || fuelDataList[0];
            const endPt = [...fuelDataList].reverse().find(p => new Date(p.dateTime).getTime() <= tStop) || fuelDataList[fuelDataList.length - 1];
            const diff = endPt.data - startPt.data;
            return diff > 0 ? diff : 0;
        }
        return 0;
    }

    // =========================================================================
    // 5. RENDERIZADO (TABLA Y GRÁFICO)
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
                        ${mode === "Device" ? "<th>Nº Serie / VIN</th>" : "<th>Viajes Analizados</th>"}
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
                    <td>${mode === "Device" ? escapeHtml(row.serialNumber) : row.tripsCount}</td>
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
    // 6. EXPORTACIÓN CSV (UTF-8 BOM)
    // =========================================================================

    function exportToCSV() {
        if (currentReportData.length === 0) return;

        const mode = document.querySelector('input[name="searchMode"]:checked').value;
        const threshold = parseFloat(document.getElementById("thresholdLimit").value) || 30.0;

        let csvContent = "\uFEFF";
        csvContent += mode === "Device" 
            ? "Vehículo;Nº Serie;Distancia (km);Combustible (L);Consumo Medio (L/100km);Estado\n"
            : "Conductor;Viajes Analizados;Distancia (km);Combustible (L);Consumo Medio (L/100km);Estado\n";

        currentReportData.forEach(row => {
            const status = !row.hasCanBus ? "Sin Datos CAN" : (row.avgConsumption > threshold ? "Excede Umbral" : "Normal");
            const col2 = mode === "Device" ? row.serialNumber : row.tripsCount;
            
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
    // 7. UTILIDADES
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
