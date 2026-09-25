/**
 * Obtiene el consumo exacto de combustible de un conductor e interpola
 * las lecturas acumuladas del diagnóstico (DiagnosticTotalFuelUsedId)
 * en los instantes exactos de inicio y fin de su actividad.
 *
 * @param {Object} api - Instancia del API de Geotab
 * @param {string} driverId - ID del usuario/conductor
 * @param {string} fromDate - Fecha de inicio (ISO 8601)
 * @param {string} toDate - Fecha de fin (ISO 8601)
 */
async function getDriverFuelConsumptionExact(api, driverId, fromDate, toDate) {
    const DIAGNOSTIC_FUEL_ID = "DiagnosticTotalFuelUsedId";

    // 1. Obtener los viajes (Trips) del conductor en el periodo
    const trips = await api.callAsync("Get", {
        typeName: "Trip",
        search: {
            userSearch: { id: driverId },
            fromDate: fromDate,
            toDate: toDate
        }
    });

    if (!trips || trips.length === 0) {
        return { totalDistanceKm: 0, totalFuelLiters: 0, avgConsumption: 0 };
    }

    // Ordenar viajes cronológicamente
    trips.sort((a, b) => new Date(a.start) - new Date(b.start));

    // 2. AGRUPACIÓN: Fusionar viajes consecutivos del mismo vehículo en un bloque continuo
    const deviceBlocks = [];
    let currentBlock = null;

    for (const trip of trips) {
        const deviceId = trip.device ? trip.device.id : null;
        if (!deviceId) continue;

        const tripStart = new Date(trip.start);
        const tripStop = new Date(trip.stop);

        if (!currentBlock || currentBlock.deviceId !== deviceId) {
            if (currentBlock) deviceBlocks.push(currentBlock);
            currentBlock = {
                deviceId: deviceId,
                start: tripStart,
                stop: tripStop,
                distanceMeters: trip.distance || 0
            };
        } else {
            // Si el conductor continúa en el mismo camión, extendemos la ventana de fin
            currentBlock.stop = tripStop;
            currentBlock.distanceMeters += (trip.distance || 0);
        }
    }
    if (currentBlock) deviceBlocks.push(currentBlock);

    let totalFuelLiters = 0;
    let totalDistanceKm = 0;

    // 3. INTERPOLACIÓN: Consultar contadores ampliando el margen de búsqueda
    for (const block of deviceBlocks) {
        totalDistanceKm += (block.distanceMeters / 1000);

        // Ventana de búsqueda extendida (1 hora antes y 1 hora después) para capturar lecturas fuera del viaje
        const marginMs = 60 * 60 * 1000;
        const searchFrom = new Date(block.start.getTime() - marginMs).toISOString();
        const searchTo = new Date(block.stop.getTime() + marginMs).toISOString();

        const statusData = await api.callAsync("Get", {
            typeName: "StatusData",
            search: {
                diagnosticSearch: { id: DIAGNOSTIC_FUEL_ID },
                deviceSearch: { id: block.deviceId },
                fromDate: searchFrom,
                toDate: searchTo
            }
        });

        if (!statusData || statusData.length < 2) continue;

        // Ordenar lecturas de combustible
        statusData.sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));

        // Obtener el valor interpolado en el segundo exacto de inicio y fin
        const fuelAtStart = interpolateFuelValue(statusData, block.start);
        const fuelAtStop = interpolateFuelValue(statusData, block.stop);

        if (fuelAtStart !== null && fuelAtStop !== null && fuelAtStop >= fuelAtStart) {
            totalFuelLiters += (fuelAtStop - fuelAtStart);
        }
    }

    const avgConsumption = totalDistanceKm > 0 ? (totalFuelLiters / totalDistanceKm) * 100 : 0;

    return {
        totalDistanceKm: Number(totalDistanceKm.toFixed(2)),
        totalFuelLiters: Number(totalFuelLiters.toFixed(2)),
        avgConsumption: Number(avgConsumption.toFixed(2))
    };
}

/**
 * Realiza una interpolación lineal para estimar el acumulado de combustible en un timestamp objetivo
 */
function interpolateFuelValue(statusDataList, targetDate) {
    const targetMs = targetDate.getTime();
    let prev = null;
    let next = null;

    for (let i = 0; i < statusDataList.length; i++) {
        const itemTime = new Date(statusDataList[i].dateTime).getTime();
        const itemVal = statusDataList[i].data;

        if (itemTime === targetMs) return itemVal;

        if (itemTime < targetMs) {
            prev = { time: itemTime, val: itemVal };
        } else if (itemTime > targetMs) {
            next = { time: itemTime, val: itemVal };
            break; // Primer punto posterior encontrado
        }
    }

    // Interpolación lineal entre punto previo y posterior
    if (prev && next) {
        const timeFraction = (targetMs - prev.time) / (next.time - prev.time);
        return prev.val + timeFraction * (next.val - prev.val);
    }

    // Si solo hay punto anterior o posterior más cercano en el margen extendido
    if (prev) return prev.val;
    if (next) return next.val;

    return null;
}