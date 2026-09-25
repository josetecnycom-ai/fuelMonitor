geotab.addin.fuelMonitor = function (outerApi, outerState) {

    let currentApi = outerApi; 
    let chartInstance = null;
    let entityMap = {}; // Guardará IDs de vehículos o conductores
    let reportDataForExport = []; // Guardaremos los datos para pasarlos al Excel

    // Referencias DOM
    const modeRadios  = document.getElementsByName('searchMode');
    const inputSearch = document.getElementById('entitySearch');
    const dataList    = document.getElementById('entityList');
    const dateFrom    = document.getElementById('dateFrom');
    const dateTo      = document.getElementById('dateTo');
    const btnFetch    = document.getElementById('btn-fetch-data');
    const btnExport   = document.getElementById('btn-export-excel');
    const panel       = document.getElementById('results-panel');
    const chartWrap   = document.getElementById('chartWrapper');

    // ─── Bypassing Drive Proxy (Mantenido de tu versión original) ─────────────
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

    // ─── Carga de Lista (Vehículos o Conductores) ─────────────────────────────
    function loadEntityList() {
        const mode = document.querySelector('input[name="searchMode"]:checked').value; // "Device" o "User"
        inputSearch.placeholder = 'Cargando...';
        inputSearch.disabled = true;

        directCall('Get', { typeName: mode }, function(entities) {
            inputSearch.disabled = false;
            entities.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            
            entityMap = {};
            dataList.innerHTML = '';
            
            entities.forEach(e => {
                // Filtramos usuarios genéricos si es necesario
                if (e.name && e.name.indexOf('*') === -1) {
                    var option = document.createElement('option');
                    option.value = e.name;
                    dataList.appendChild(option);
                    entityMap[e.name.trim().toUpperCase()] = e.id;
                }
            });
            inputSearch.placeholder = mode === 'Device' ? 'Selecciona un vehículo...' : 'Selecciona un conductor...';
        }, function(error) {
            console.error('Error cargando lista', error);
            inputSearch.placeholder = 'Error cargando datos. Escribe el ID exacto.';
            inputSearch.disabled = false;
        });
    }

    // ─── Procesamiento de Datos de Combustible ────────────────────────────────
    function loadFuelData(entityId, entityName, mode) {
        if (!dateFrom.value || !dateTo.value) {
            panel.innerHTML = '<p style="color:red;">Por favor, selecciona las fechas Desde y Hasta.</p>';
            return;
        }

        panel.innerHTML = '<p>Analizando datos de consumo. Por favor espera...</p>';
        btnExport.style.display = 'none';
        chartWrap.style.display = 'none';
        reportDataForExport = [];

        // Aseguramos formato ISO para las fechas
        const fromDate = new Date(dateFrom.value + "T00:00:00Z").toISOString();
        const toDate = new Date(dateTo.value + "T23:59:59Z").toISOString();

        if (mode === 'Device') {
            // Para Vehículos: Calculamos la diferencia del odómetro y el combustible total usado en el periodo
            var calls = [
                ['Get', { typeName: 'StatusData', search: { deviceSearch: { id: entityId }, diagnosticSearch: { id: 'DiagnosticOdometerAdjustmentId' }, fromDate: fromDate, toDate: toDate } }],
                ['Get', { typeName: 'StatusData', search: { deviceSearch: { id: entityId }, diagnosticSearch: { id: 'DiagnosticTotalFuelUsedId' }, fromDate: fromDate, toDate: toDate } }]
            ];

            directMultiCall(calls, function(results) {
                const odoData = results[0];
                const fuelData = results[1];

                if (!odoData.length || !fuelData.length) {
                    panel.innerHTML = '<p>No hay suficientes datos de telemetría en este periodo para calcular medias.</p>';
                    return;
                }

                // Cálculo de deltas
                const distanceKm = (odoData[odoData.length - 1].data - odoData[0].data) / 1000; // Geotab devuelve metros
                const fuelLiters = fuelData[fuelData.length - 1].data - fuelData[0].data;
                const avgConsumption = (distanceKm > 0) ? (fuelLiters / distanceKm) * 100 : 0;

                renderResults(entityName, distanceKm, fuelLiters, avgConsumption);
                renderChart(fuelData); // Mostramos progresión de consumo

                // Preparamos datos para Excel
                reportDataForExport.push({
                    Activo: entityName,
                    Distancia_Km: distanceKm.toFixed(2),
                    Combustible_Litros: fuelLiters.toFixed(2),
                    Media_L_100Km: avgConsumption.toFixed(2)
                });
                btnExport.style.display = 'block';

            }, function(err) {
                panel.innerHTML = '<p style="color:red;">Error al obtener la telemetría.</p>';
            });

        } else {
            // Para Conductores: Buscamos sus Viajes (Trips) en el periodo.
            // Nota: Para obtener el combustible exacto por conductor, lo ideal es sumar la distancia de sus Trips. 
            // (El combustible avanzado requeriría cruzar cada Trip con StatusData, aquí haremos una aproximación con Trips).
            directCall('Get', {
                typeName: 'Trip',
                search: { userSearch: { id: entityId }, fromDate: fromDate, toDate: toDate }
            }, function(trips) {
                if (!trips || trips.length === 0) {
                    panel.innerHTML = '<p>No se registraron viajes para este conductor en las fechas seleccionadas.</p>';
                    return;
                }

                let totalDistance = 0;
                // En flotas avanzadas, Geotab puede calcular la media si hay reglas de IFTA asignadas.
                // Aquí sumamos los kilómetros.
                trips.forEach(t => { totalDistance += (t.distance || 0); });
                
                const distanceKm = totalDistance;
                
                panel.innerHTML = `
                    <div style="padding: 15px; background: #f1f5f9; border-radius: 8px; margin-top: 15px;">
                        <h3 style="margin-top:0;">Resumen del Conductor: ${entityName}</h3>
                        <p><strong>Distancia total recorrida:</strong> ${distanceKm.toFixed(2)} Km</p>
                        <p><em>*Nota: Para calcular los litros exactos por conductor, es necesario cruzar el historial de los vehículos que ha conducido.</em></p>
                    </div>
                `;

                reportDataForExport.push({ Conductor: entityName, Distancia_Km: distanceKm.toFixed(2) });
                btnExport.style.display = 'block';
            });
        }
    }

    // ─── Renderizado de Resultados y Gráficas ─────────────────────────────────
    function renderResults(name, distance, fuel, avg) {
        panel.innerHTML = `
            <div style="padding: 15px; background: #e0f2fe; border-radius: 8px; margin-top: 15px; border: 1px solid #bae6fd;">
                <h3 style="margin-top:0; color: #0369a1;">Resumen de Vehículo: ${name}</h3>
                <ul style="font-size: 16px; line-height: 1.8;">
                    <li><strong>Distancia Recorrida:</strong> ${distance.toFixed(2)} Km</li>
                    <li><strong>Combustible Consumido:</strong> ${fuel.toFixed(2)} Litros</li>
                    <li><strong>Consumo Medio:</strong> <span style="color: ${avg > 35 ? 'red' : 'green'}; font-weight: bold;">${avg.toFixed(2)} L/100km</span></li>
                </ul>
            </div>
        `;
    }

    function renderChart(fuelData) {
        chartWrap.style.display = 'block';
        const ctx = document.getElementById('fuelChart').getContext('2d');
        
        // Mapeamos los datos de StatusData a formato para Chart.js
        const dataset = fuelData.map(d => ({ x: new Date(d.dateTime), y: d.data }));

        if (chartInstance) chartInstance.destroy();

        chartInstance = new Chart(ctx, {
            type: 'line',
            data: {
                datasets: [{
                    label: 'Consumo Acumulado (Litros)',
                    data: dataset,
                    borderColor: '#2563eb',
                    backgroundColor: 'rgba(37, 99, 235, 0.1)',
                    fill: true,
                    tension: 0.1
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                scales: {
                    x: { type: 'time', title: { display: true, text: 'Fecha' } },
                    y: { title: { display: true, text: 'Litros Totales Acumulados' } }
                }
            }
        });
    }

    // ─── Exportar a Excel (CSV) ───────────────────────────────────────────────
    function exportToCSV() {
        if (!reportDataForExport.length) return;
        
        const headers = Object.keys(reportDataForExport[0]);
        const csvRows = [];
        
        // Cabeceras
        csvRows.push(headers.join(';'));
        
        // Filas
        for (const row of reportDataForExport) {
            const values = headers.map(header => row[header]);
            csvRows.push(values.join(';'));
        }
        
        const blob = new Blob(["\uFEFF" + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", "Informe_Combustible.csv");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // ─── Ciclo de vida del Add-In ─────────────────────────────────────────────
    return {
        initialize: function (api, state, callback) {
            currentApi = api;

            // Por defecto ponemos la fecha de hoy
            const today = new Date().toISOString().split('T')[0];
            dateFrom.value = today;
            dateTo.value = today;

            // Listeners
            modeRadios.forEach(radio => {
                radio.addEventListener('change', () => {
                    inputSearch.value = '';
                    loadEntityList();
                });
            });

            inputSearch.addEventListener('focus', loadEntityList);

            btnFetch.addEventListener('click', () => {
                const name = inputSearch.value.trim().toUpperCase();
                const mode = document.querySelector('input[name="searchMode"]:checked').value;
                const id = entityMap[name];

                if (id) {
                    loadFuelData(id, inputSearch.value, mode);
                } else {
                    panel.innerHTML = '<p style="color:red;">Por favor, selecciona un vehículo o conductor válido de la lista.</p>';
                }
            });

            btnExport.addEventListener('click', exportToCSV);

            callback();
        },
        focus: function (api, state) { },
        blur: function () { }
    };
};