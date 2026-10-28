let teacherData = [];

// Excel файлды оқу
document.getElementById('excelFile').addEventListener('change', handleFile, false);

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        
        // Бірінші парақты алу
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // JSON форматына айналдыру
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // Массивке сақтау (бастапқы ескерту деңгейі = 0)
        teacherData = jsonData.map((row, index) => ({
            id: index + 1,
            name: row['Аты-жөні'] || row['ФИО'] || 'Көрсетілмеген',
            subject: row['Пәні'] || row['Предмет'] || 'Көрсетілмеген',
            grade: row['Сынып'] || row['Класс'] || '-',
            time: row['Сабақ уақыты'] || row['Время'] || '-',
            warningLevel: 0 // 0: Жасыл, 1: Сары, 2: Қызғылт сары, 3: Қызыл
        }));

        renderTable();
    };
    reader.readAsArrayBuffer(file);
}

// Кестені экранға шығару
function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (teacherData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;">Excel файлын жүктеңіз немесе мәлімет табылмады.</td></tr>';
        return;
    }

    teacherData.forEach((teacher) => {
        const tr = document.createElement('tr');
        tr.className = `status-${teacher.warningLevel}`;

        let statusText = '';
        if (teacher.warningLevel === 0) statusText = '🟢 Қалыпты';
        else if (teacher.warningLevel === 1) statusText = '🟡 1-ші ескерту (Сабақ жоспары жоқ)';
        else if (teacher.warningLevel === 2) statusText = '🟠 2-ші ескерту (Себепсіз кешікті / Жоспар дайын емес)';
        else if (teacher.warningLevel === 3) statusText = '🔴 3-ші ескерту (Сабаққа себепсіз кірмеді)';

        tr.innerHTML = `
            <td>${teacher.id}</td>
            <td><strong>${teacher.name}</strong></td>
            <td>${teacher.subject}</td>
            <td>${teacher.grade}</td>
            <td>${teacher.time}</td>
            <td><strong>${statusText}</strong></td>
            <td>
                <button class="btn btn-green" onclick="setWarning(${teacher.id}, 0)">🟢</button>
                <button class="btn btn-yellow" onclick="setWarning(${teacher.id}, 1)">🟡</button>
                <button class="btn btn-orange" onclick="setWarning(${teacher.id}, 2)">🟠</button>
                <button class="btn btn-red" onclick="setWarning(${teacher.id}, 3)">🔴</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// Ескерту мәртебесін ауыстыру
function setWarning(id, level) {
    const teacher = teacherData.find(t => t.id === id);
    if (teacher) {
        teacher.warningLevel = level;
        renderTable();
    }
}