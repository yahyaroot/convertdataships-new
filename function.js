/* ================= GLOBAL ================= */
let filteredData = [];


const ORDER = ['Cargo', 'Fishing', 'Other', 'Passenger', 'Tanker', 'Tug'];

const APPS_SCRIPT_URL =
    "https://script.google.com/macros/s/AKfycbxdTzKV7-tLkdpJPKTURK97mc_e5gxsw2me9homDMq7FnjQAkVOXmUHet1KwxPwWdQE/exec";

/* ================= LOADING UI ================= */
function setLoading(btn, loading = true, text = "Processing...") {
    if (!btn) return;

    if (loading) {
        btn.dataset.original = btn.innerHTML;
        btn.innerHTML = `⏳ ${text}`;
        btn.disabled = true;
        btn.style.opacity = "0.7";
    } else {
        btn.innerHTML = btn.dataset.original;
        btn.disabled = false;
        btn.style.opacity = "1";
    }
}

/* ================= UTIL ================= */
function normalizeType(t) {
    t = (t || "").toLowerCase();
    if (t.includes("cargo")) return "Cargo";
    if (t.includes("fish")) return "Fishing";
    if (t.includes("pass")) return "Passenger";
    if (t.includes("tank")) return "Tanker";
    if (t.includes("tug")) return "Tug";
    return "Other";
}

function splitLat(v) {
    if (!v) return ["", "", "", ""];
    v = parseFloat(v);
    const dir = v < 0 ? "S" : "N";
    v = Math.abs(v);
    const d = Math.floor(v);
    const mF = (v - d) * 60;
    const m = Math.floor(mF);
    const s = ((mF - m) * 60).toFixed(1);
    return [d, m, s, dir];
}

function splitLon(v) {
    if (!v) return ["", "", ""];
    v = Math.abs(parseFloat(v));
    const d = Math.floor(v);
    const mF = (v - d) * 60;
    const m = Math.floor(mF);
    const s = ((mF - m) * 60).toFixed(1);
    return [d, m, s];
}

/* ================= PARSER ================= */
function parseCSV(text) {
    const lines = text.trim().split(/\r?\n/);
    const headers = lines[0].split(",").map(h => h.trim().toLowerCase());

    return lines.slice(1).map(line => {
        const cols = line.split(",");
        const o = {};
        headers.forEach((h, i) => o[h] = cols[i] || "");
        return o;
    });
}

function parseXLSX(buffer) {
    const wb = XLSX.read(buffer, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

    return json.map(row => {
        const o = {};
        Object.keys(row).forEach(k => o[k.toLowerCase()] = row[k]);
        return o;
    });
}

/* ================= UI ================= */
function renderTable() {
    document.getElementById("thead").innerHTML = `
    <tr>
      <th>No</th><th>Name</th><th>Callsign</th><th>MMSI</th>
      <th>Course</th><th>Speed</th><th>IMO</th><th>Flag</th>
      <th>Type</th><th>Destination</th>
      <th>Latitude</th><th>Longitude</th><th>Status</th>
    </tr>`;

    document.getElementById("tbody").innerHTML =
        filteredData.map((r, i) => `
      <tr>
        <td>${i + 1}</td>
        <td>${r.name}</td>
        <td>${r.callsign}</td>
        <td>${r.mmsi}</td>
        <td>${r.course}</td>
        <td>${r.speed}</td>
        <td>${r.imo}</td>
        <td>${r.flag}</td>
        <td>${r.vessel_type}</td>
        <td>${r.destination}</td>
        <td>${r.latitude1}° ${r.latitude2}' ${r.latitude3}" ${r.latitude4}</td>
        <td>${r.longitude1}° ${r.longitude2}' ${r.longitude3}"</td>
        <td>${r.status}</td>
      </tr>`).join("");
}

function renderSummary() {
    const counts = {};
    ORDER.forEach(o => counts[o] = 0);
    filteredData.forEach(r => counts[r.vessel_type]++);
    document.getElementById("summary").innerHTML =
        ORDER.map(o => `<span class="pill"><b>${o}</b>: ${counts[o]}</span>`).join("");
}

/* ================= IMPORT ================= */
document.getElementById("fileInput").addEventListener("change", e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    const isExcel = /\.(xls|xlsx)$/i.test(file.name);

    reader.onload = () => {
        const raw = isExcel ? parseXLSX(reader.result) : parseCSV(reader.result);

        const unique = new Map();
        raw.forEach(r => {
            const key = r.mmsi || r.name;
            if (!unique.has(key)) unique.set(key, r);
        });

        filteredData = Array.from(unique.values()).map(r => {
            const [la1, la2, la3, la4] = splitLat(r.latitude);
            const [lo1, lo2, lo3] = splitLon(r.longitude);

            return {
                name: r.name || "",
                callsign: r.callsign || "",
                mmsi: r.mmsi || "",
                course: parseFloat((r.course || "").replace(/[^\d.]/g, "")) || "",
                speed: parseFloat((r.speed || "").replace(/[^\d.]/g, "")) || "",
                imo: r.imo || "",
                flag: r.flag || "",
                vessel_type: normalizeType(r.vessel_type),
                destination: r.destination || "",
                latitude1: la1,
                latitude2: la2,
                latitude3: la3,
                latitude4: la4,
                longitude1: lo1,
                longitude2: lo2,
                longitude3: lo3,
                status: (parseFloat(r.speed) || 0) >= 3 ? "MELINTAS" : "LEGO"
            };
        });

        filteredData.sort(
            (a, b) => ORDER.indexOf(a.vessel_type) - ORDER.indexOf(b.vessel_type)
        );

        renderTable();
        renderSummary();
    };

    isExcel ? reader.readAsArrayBuffer(file) : reader.readAsText(file);
});

/* ================= EXPORT HELPERS ================= */
function exportFetch(btn, url, filename, data = null) {
    const exportData = data || filteredData;

    if (!exportData.length) {
        alert("Tidak ada data");
        return;
    }

    setLoading(btn, true, "Exporting...");

    fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(exportData)
        })
        .then(res => {
            if (!res.ok) throw new Error("Export gagal");
            return res.blob();
        })
        .then(blob => {
            const a = document.createElement("a");
            a.href = URL.createObjectURL(blob);
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(a.href);
        })
        .catch(() => alert("Gagal export file"))
        .finally(() => setLoading(btn, false));
}

/* ================= EXPORT FUNCTIONS ================= */
function exportViaTemplatePHPHarian(btn) {
    exportFetch(btn, "export_laporan_harian.php", "rekap_harian_ais.xlsx");
}

function exportViaTemplatePHP(btn) {
    exportFetch(btn, "export_template.php", "laporan_harian_ais.xlsx");
}

function exportViaTemplateKapalAsing(btn) {
    if (!filteredData.length) {
        alert("Tidak ada data");
        return;
    }

    // ✅ Filter hanya kapal asing (flag bukan Indonesia)
    const foreignShips = filteredData.filter(r =>
        r.flag && r.flag.toLowerCase() !== 'indonesia'
    );

    if (!foreignShips.length) {
        alert("Tidak ada data kapal asing");
        return;
    }

    // Kirim data yang sudah difilter
    exportFetch(btn, "export_laporan_kapal_asing.php", "laporan_kapal_asing.xlsx", foreignShips);
}

function exportViaTemplateKapalIkan(btn) {
    if (!filteredData.length) {
        alert("Tidak ada data");
        return;
    }

    // ✅ Filter hanya kapal ikan
    const fishingShips = filteredData.filter(r => r.vessel_type === 'Fishing');

    if (!fishingShips.length) {
        alert("Tidak ada data kapal ikan");
        return;
    }

    // Kirim data yang sudah difilter
    exportFetch(btn, "export_laporan_kapal_ikan.php", "laporan_kapal_ikan.xlsx", fishingShips);
}