let scholarshipsData = [];

document.addEventListener('DOMContentLoaded', () => {
    fetchScholarships();
});

async function fetchScholarships() {
    try {
        const response = await fetch('scholarships.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        scholarshipsData = jsonData.map(row => ({
            degree: row.Degree || '',
            title: row.Title || '',
            country: row.Country || '',
            link: row.Link || ''
        }));

        renderTable(scholarshipsData);
    } catch (error) {
        console.error('Error loading Excel file:', error);
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