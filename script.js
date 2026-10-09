let dataList = [];
let currentLang = 'en';

// متغيرا التقسيم (Pagination)
let currentPage = 1;
const itemsPerPage = 10;
let filteredData = [];

const translations = {
    en: {
        title: "Scholarship Directory",
        subtitle: "Find fully funded bachelor, master, and PhD opportunities worldwide.",
        by: "By",
        searchLabel: "Search Title:",
        ph: "e.g. Miami, Yale, Harvard...",
        degreeLabel: "Degree:",
        allDegrees: "All Degrees",
        bachelor: "Bachelor",
        master: "Master",
        phd: "PhD",
        countryLabel: "Country:",
        allCountries: "All Countries",
        searchBtn: "Search",
        thDegree: "Degree",
        thTitle: "Scholarship Title",
        thCountry: "Country",
        thDeadline: "Deadline",
        thAction: "Action",
        applyBtn: "Apply Now",
        noData: "No scholarships found matching your search.",
        footerText: "Developed by Habib Amer",
        lastUpdated: "Last updated",
        prevBtn: "Previous",
        nextBtn: "Next"
    },
    ar: {
        title: "دليل المنح الدراسية",
        subtitle: "ابحث عن فرص دراسية كاملة التمويل لمرحلة البكالوريوس والماجستير والدكتوراه حول العالم.",
        by: "بواسطة",
        searchLabel: "عنوان البحث:",
        ph: "مثال: ميامي، ييل، هارفارد...",
        degreeLabel: "الدرجة العلمية:",
        allDegrees: "جميع الدرجات",
        bachelor: "بكالوريوس",
        master: "ماجستير",
        phd: "دكتوراه",
        countryLabel: "الدولة:",
        allCountries: "جميع البلدان",
        searchBtn: "بحث",
        thDegree: "الدرجة",
        thTitle: "عنوان المنحة الدراسية",
        thCountry: "الدولة",
        thDeadline: "الموعد النهائي",
        thAction: "الإجراء",
        applyBtn: "قدّم الآن",
        noData: "لم يتم العثور على منح تطابق بحثك.",
        footerText: "تطوير حبيب عامر",
        lastUpdated: "آخر تحديث",
        prevBtn: "السابق",
        nextBtn: "التالي"
    },
    ru: {
        title: "Каталог Стипендий",
        subtitle: "Найдите полностью финансируемые программы бакалавриата, магистратуры и аспирантуры по всему миру.",
        by: "Автор",
        searchLabel: "Поиск по названию:",
        ph: "например, Майами, Йель, Гарвард...",
        degreeLabel: "Степень:",
        allDegrees: "Все степени",
        bachelor: "Бакалавриат",
        master: "Магистратура",
        phd: "Аспирантура",
        countryLabel: "Страна:",
        allCountries: "Все страны",
        searchBtn: "Искать",
        thDegree: "Степень",
        thTitle: "Название стипендии",
        thCountry: "Страна",
        thDeadline: "Крайний срок",
        thAction: "Действие",
        applyBtn: "Подать заявку",
        noData: "Стипендии, соответствующие вашему запросу, не найдены.",
        footerText: "Разработано Хабибом Амером",
        lastUpdated: "Последнее обновление",
        prevBtn: "Назад",
        nextBtn: "Вперед"
    },
    es: {
        title: "Directorio de Becas",
        subtitle: "Encuentra oportunidades totalmente financiadas para licenciatura, maestría y doctorado en todo el mundo.",
        by: "Por",
        searchLabel: "Buscar por título:",
        ph: "ej. Miami, Yale, Harvard...",
        degreeLabel: "Grado:",
        allDegrees: "Todos los grados",
        bachelor: "Licenciatura",
        master: "Maestría",
        phd: "Doctorado",
        countryLabel: "País:",
        allCountries: "Todos los países",
        searchBtn: "Buscar",
        thDegree: "Grado",
        thTitle: "Título de la beca",
        thCountry: "País",
        thDeadline: "Fecha límite",
        thAction: "Acción",
        applyBtn: "Postular ahora",
        noData: "No se encontraron becas que coincidan con tu búsqueda.",
        footerText: "Desarrollado por Habib Amer",
        lastUpdated: "Última actualización",
        prevBtn: "Anterior",
        nextBtn: "Siguiente"
    }
};

document.addEventListener('DOMContentLoaded', () => {
    loadData();
    loadLastUpdated();
});

function changeLanguage(lang) {
    currentLang = lang;
    document.body.dir = lang === 'ar' ? 'rtl' : 'ltr';

    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (translations[lang][key]) {
            el.textContent = translations[lang][key];
        }
    });

    document.querySelectorAll('[data-i18n-ph]').forEach(el => {
        const key = el.getAttribute('data-i18n-ph');
        if (translations[lang][key]) {
            el.placeholder = translations[lang][key];
        }
    });

    runFilter();
}

async function loadData() {
    try {
        const res = await fetch('scholarships.xlsx?t=' + new Date().getTime());
        if (!res.ok) throw new Error("Scholarships file not found");
        
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(1);

        dataList = rows.map(r => {
            if (!r || r.length === 0) return null;
            
            let rawDeadline = String(r[3] || '').trim();
            let rawLink = String(r[4] || '').trim();

            if (rawDeadline.startsWith('http')) {
                rawLink = rawDeadline;
                rawDeadline = 'N/A';
            }

            return {
                deg: String(r[0] || '').trim(),
                title: String(r[1] || '').trim(),
                country: String(r[2] || '').trim(),
                deadline: rawDeadline,
                link: rawLink
            };
        }).filter(item => item && item.title);

        fillCountries(dataList);
        filteredData = [...dataList];
        currentPage = 1;
        showTable();
    } catch (err) {
        console.error('Error loading file:', err);
    }
}

async function loadLastUpdated() {
    try {
        const r = await fetch('last_updated.json?t=' + Date.now());
        if (!r.ok) return;
        const d = await r.json();
        document.getElementById('lastUpdated').textContent =
            new Date(d.updated).toLocaleDateString('en-GB');
    } catch (e) {
        console.error('Error loading last updated date:', e);
    }
}

function fillCountries(list) {
    const sel = document.getElementById('country');
    sel.innerHTML = `<option value="all" data-i18n="allCountries">${translations[currentLang].allCountries}</option>`;
    
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

function showTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (filteredData.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;">${translations[currentLang].noData}</td></tr>`;
        renderPaginationControls(0);
        return;
    }

    // اقتصاص 10 منح فقط بناءً على رقم الصفحة الحالية
    const startIdx = (currentPage - 1) * itemsPerPage;
    const endIdx = startIdx + itemsPerPage;
    const pageItems = filteredData.slice(startIdx, endIdx);

    pageItems.forEach(item => {
        const tr = document.createElement('tr');
        const badgeClass = item.deg.toLowerCase();

        const deadlineHtml = (item.deadline && item.deadline !== 'N/A') 
            ? `<span class="deadline-tag">${item.deadline}</span>` 
            : `<span style="color:#94a3b8; font-size:0.85rem;">-</span>`;

        tr.innerHTML = `
            <td><span class="badge ${badgeClass}">${item.deg}</span></td>
            <td>${item.title}</td>
            <td>${item.country}</td>
            <td>${deadlineHtml}</td>
            <td><a href="${item.link}" target="_blank" class="apply-btn">${translations[currentLang].applyBtn}</a></td>
        `;
        tbody.appendChild(tr);
    });

    renderPaginationControls(filteredData.length);
}

// دالة لإنشاء أزرار التنقل بين الصفحات أسفل الجدول
function renderPaginationControls(totalItems) {
    let container = document.getElementById('paginationControls');
    
    // في حال عدم وجود حاوية الأزرار في ملف HTML، يتم إنشاؤها تلقائياً تحت الجدول
    if (!container) {
        container = document.createElement('div');
        container.id = 'paginationControls';
        container.className = 'pagination-container';
        const table = document.querySelector('table');
        if (table && table.parentNode) {
            table.parentNode.insertBefore(container, table.nextSibling);
        }
    }

    container.innerHTML = '';
    const totalPages = Math.ceil(totalItems / itemsPerPage);

    if (totalPages <= 1) return;

    // زر السابق
    const prevBtn = document.createElement('button');
    prevBtn.textContent = translations[currentLang].prevBtn;
    prevBtn.disabled = currentPage === 1;
    prevBtn.onclick = () => {
        if (currentPage > 1) {
            currentPage--;
            showTable();
        }
    };
    container.appendChild(prevBtn);

    // أزرار الأرقام (1, 2, 3...)
    for (let i = 1; i <= totalPages; i++) {
        const pageBtn = document.createElement('button');
        pageBtn.textContent = i;
        if (i === currentPage) {
            pageBtn.classList.add('active');
        }
        pageBtn.onclick = () => {
            currentPage = i;
            showTable();
        };
        container.appendChild(pageBtn);
    }

    // زر التالي
    const nextBtn = document.createElement('button');
    nextBtn.textContent = translations[currentLang].nextBtn;
    nextBtn.disabled = currentPage === totalPages;
    nextBtn.onclick = () => {
        if (currentPage < totalPages) {
            currentPage++;
            showTable();
        }
    };
    container.appendChild(nextBtn);
}

function runFilter() {
    const k = document.getElementById('kw').value.toLowerCase().trim();
    const d = document.getElementById('degree').value.toLowerCase();
    const c = document.getElementById('country').value.toLowerCase();

    filteredData = dataList.filter(item => {
        const matchKw = (k === '') || item.title.toLowerCase().includes(k);
        const matchDeg = (d === 'all') || item.deg.toLowerCase().includes(d);
        const matchCountry = (c === 'all') || item.country.toLowerCase() === c;

        return matchKw && matchDeg && matchCountry;
    });

    currentPage = 1; 
    showTable();
}