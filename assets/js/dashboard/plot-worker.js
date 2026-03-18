// Web Worker for heavy plot/statistics computations
// Runs in a separate thread; does not access DOM
console.log('[plot-worker] initialized');

function calculatePercentileFromSorted(sorted, p) {
    if (!Array.isArray(sorted) || sorted.length === 0) return null;
    const index = (sorted.length - 1) * p;
    const lower = Math.floor(index);
    const upper = Math.ceil(index);
    if (upper === lower) return sorted[lower];
    const weight = index - lower;
    return sorted[lower] + (sorted[upper] - sorted[lower]) * weight;
}

function calculateMeanAndCI(arr) {
    if (!arr || arr.length === 0) return { mean: null, lowerCI: null, upperCI: null };
    const mean = arr.reduce((a, b) => a + b, 0) / arr.length;
    const sumSqDiff = arr.reduce((acc, val) => acc + (val - mean) ** 2, 0);
    const stdErr = Math.sqrt(sumSqDiff / (arr.length > 1 ? arr.length - 1 : 1)) / Math.sqrt(arr.length);
    // Use normal approx (z=1.645 for ~90%) as lightweight t-table replacement in worker
    const z = 1.645;
    const me = z * stdErr;
    return { mean, lowerCI: mean - me, upperCI: mean + me };
}

self.onmessage = function(e) {
    const data = e.data;
    if (!data || data.action !== 'computeStats' || !Array.isArray(data.concsAtTime)) return;
    console.log('[plot-worker] received message', { id: data.id, trialIndex: data.trialIndex, frames: data.concsAtTime.length });
    const concsAtTime = data.concsAtTime;
    const stats = { means: [], medians: [], p00: [], p100: [], p025: [], p975: [], p05: [], p95: [], p25: [], p75: [], lowerCI: [], upperCI: [] };

    for (let i = 0; i < concsAtTime.length; i++) {
        const concs = concsAtTime[i];
        if (concs && concs.length > 0) {
            // Filter finite values
            const vals = concs.filter(v => Number.isFinite(v));
            if (vals.length === 0) {
                Object.keys(stats).forEach(k => stats[k].push(null));
                continue;
            }
            const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
            const sorted = vals.slice().sort((a, b) => a - b);

            stats.means.push(mean);
            stats.medians.push(calculatePercentileFromSorted(sorted, 0.5));
            stats.p00.push(sorted[0]);
            stats.p100.push(sorted[sorted.length - 1]);
            stats.p025.push(calculatePercentileFromSorted(sorted, 0.025));
            stats.p975.push(calculatePercentileFromSorted(sorted, 0.975));
            stats.p05.push(calculatePercentileFromSorted(sorted, 0.05));
            stats.p95.push(calculatePercentileFromSorted(sorted, 0.95));
            stats.p25.push(calculatePercentileFromSorted(sorted, 0.25));
            stats.p75.push(calculatePercentileFromSorted(sorted, 0.75));

            const ci = calculateMeanAndCI(vals);
            stats.lowerCI.push(ci.lowerCI);
            stats.upperCI.push(ci.upperCI);
        } else {
            Object.keys(stats).forEach(k => stats[k].push(null));
        }
    }

    // Post result back to main thread
    console.log('[plot-worker] posting result', { id: data.id, trialIndex: data.trialIndex });
    self.postMessage({ id: data.id, trialIndex: data.trialIndex, stats });
};
