let dataList = [];

document.addEventListener('DOMContentLoaded', () => {
    loadData();
});

async function loadData() {
    try {
        const res = await fetch('scholarships.xlsx?t=' + new Date().getTime());
        const buf = await res.arrayBuffer();
        
        const wb = XLSX.read(buf, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(1);

        dataList = rows.map(r => {
            if (!r || r.length < 4) return null;
            return {
                deg: String(r[0] || '').trim(),
                title: String(r[1] || '').trim(),
                country: String(r[2] || '').trim(),
                link: String(r[3] || '').trim()
            };
        }).filter(item => item && item.title);

        fillCountries(dataList);
        showTable(dataList);
    } catch (err) {
        console.error('Error loading file:', err);
    }
}

function fillCountries(list) {
    const sel = document.getElementById('country');
    const countries = [...new Set(list.map(x => x.country))].sort();

    countries.forEach(c => {
        if (c) {
            const opt = document.createElement('option');
            opt.value = c.toLowerCase();
            opt.textContent = c;
            sel.appendChild(opt);
        }
    });
}

function showTable(list) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (list.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">No scholarships found matching your search.</td></tr>';
        return;
    }

    list.forEach(item => {
        const tr = document.createElement('tr');
        const badgeClass = item.deg.toLowerCase();

        tr.innerHTML = `
            <td><span class="badge ${badgeClass}">${item.deg}</span></td>
            <td>${item.title}</td>
            <td>${item.country}</td>
            <td><a href="${item.link}" target="_blank" class="apply-btn">Apply Now</a></td>
        `;
        tbody.appendChild(tr);
    });
}

function runFilter() {
    const k = document.getElementById('kw').value.toLowerCase().trim();
    const d = document.getElementById('degree').value.toLowerCase();
    const c = document.getElementById('country').value.toLowerCase();

    const result = dataList.filter(item => {
        const matchKw = (k === '') || item.title.toLowerCase().includes(k);
        const matchDeg = (d === 'all') || item.deg.toLowerCase().includes(d);
        const matchCountry = (c === 'all') || item.country.toLowerCase() === c;

        return matchKw && matchDeg && matchCountry;
    });

    showTable(result);
}