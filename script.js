let scholarshipsData = [];

document.addEventListener('DOMContentLoaded', () => {
    fetchScholarships();
});

async function fetchScholarships() {
    try {
        const response = await fetch('scholarships.csv');
        const data = await response.text();
        
        const rows = data.split('\n').slice(1);
        
        scholarshipsData = rows.map(row => {
            const cols = row.split(/[,;]/); 
            if (cols.length >= 4) {
                return {
                    degree: cols[0].replace(/"/g, '').trim(),
                    title: cols[1].replace(/"/g, '').trim(),
                    country: cols[2].replace(/"/g, '').trim(),
                    link: cols[3].replace(/"/g, '').trim()
                };
            }
        }).filter(item => item && item.title);

        renderTable(scholarshipsData);
    } catch (error) {
        console.error('Error loading scholarships data:', error);
    }
}

function renderTable(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">No scholarships found.</td></tr>';
        return;
    }

    data.forEach(item => {
        const tr = document.createElement('tr');
        const badgeClass = item.degree.toLowerCase();

        tr.innerHTML = `
            <td><span class="badge ${badgeClass}">${item.degree}</span></td>
            <td>${item.title}</td>
            <td>${item.country}</td>
            <td><a href="${item.link}" target="_blank" class="apply-btn">Apply Now</a></td>
        `;
        tbody.appendChild(tr);
    });
}

function filterScholarships() {
    const selectedDegree = document.getElementById('degreeFilter').value.toLowerCase();

    if (selectedDegree === 'all') {
        renderTable(scholarshipsData);
    } else {
        const filtered = scholarshipsData.filter(item => 
            item.degree.toLowerCase().includes(selectedDegree)
        );
        renderTable(filtered);
    }
}