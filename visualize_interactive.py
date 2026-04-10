"""
Интерактивная визуализация парсера с выбором файла и групп
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import glob
import json

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_cell_text(cell):
    texts = cell.findall('.//w:t', NS)
    return ''.join([t.text for t in texts if t.text])

def get_colspan(cell):
    gridSpan = cell.find('.//w:gridSpan', NS)
    if gridSpan is not None:
        colspan_val = gridSpan.get(f'{{{NS["w"]}}}val')
        if colspan_val:
            return int(colspan_val)
    return 1

def has_vmerge(cell):
    vMerge = cell.find('.//w:vMerge', NS)
    return vMerge is not None

def extract_group_name(text):
    pattern = r'\b\d{1,2}[А-ЯA-Z]{2,4}\d{2}\b'
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(0) if match else None

def parse_file(file_path):
    """Парсит файл и возвращает данные для визуализации"""
    try:
        with zipfile.ZipFile(file_path, 'r') as docx:
            xml_content = docx.read('word/document.xml')
            root = ET.fromstring(xml_content)
            
            tables = root.findall('.//w:tbl', NS)
            if not tables:
                return None
            
            # Берем самую большую таблицу (с наибольшим количеством строк)
            table = max(tables, key=lambda t: len(t.findall('.//w:tr', NS)))
            rows = table.findall('.//w:tr', NS)
            
            if len(rows) < 3:
                return None
            
            # Определяем группы из заголовка
            header_row = rows[0]
            header_cells = header_row.findall('.//w:tc', NS)
            
            groups = []
            logical_col_to_group = {}
            logical_col = 0
            
            for cell in header_cells:
                text = get_cell_text(cell)
                group = extract_group_name(text)
                colspan = get_colspan(cell)
                
                if group:
                    groups.append(group)
                    for i in range(colspan):
                        logical_col_to_group[logical_col + i] = group
                
                logical_col += colspan
            
            # Собираем данные строк
            rows_data = []
            for row_idx in range(min(100, len(rows))):
                row = rows[row_idx]
                cells = row.findall('.//w:tc', NS)
                
                cells_data = []
                logical_col = 0
                for cell_idx, cell in enumerate(cells):
                    text = get_cell_text(cell)
                    colspan = get_colspan(cell)
                    vmerge = has_vmerge(cell)
                    
                    group = logical_col_to_group.get(logical_col)
                    
                    cells_data.append({
                        'text': text,
                        'colspan': colspan,
                        'vmerge': vmerge,
                        'cell_idx': cell_idx,
                        'logical_col': logical_col,
                        'group': group
                    })
                    
                    logical_col += colspan
                
                rows_data.append(cells_data)
            
            return {
                'groups': groups,
                'rows': rows_data,
                'logical_col_to_group': logical_col_to_group
            }
    except Exception as e:
        print(f"Ошибка при парсинге {file_path}: {e}")
        return None

# Получаем список всех DOCX файлов
docx_files = glob.glob('schedules/*.docx')
docx_files.sort()

print(f"Найдено {len(docx_files)} DOCX файлов")

# Парсим все файлы
files_data = {}
for file_path in docx_files:
    file_name = file_path.replace('schedules/', '').replace('schedules\\', '')
    data = parse_file(file_path)
    if data:
        files_data[file_name] = data
        print(f"✅ {file_name}: {len(data['groups'])} групп")

# Генерируем HTML
html = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Интерактивная визуализация парсера</title>
    <style>
        body { 
            font-family: Arial; 
            font-size: 12px; 
            margin: 0;
            padding: 20px;
        }
        .controls {
            position: sticky;
            top: 0;
            background: white;
            padding: 20px;
            border-bottom: 2px solid #333;
            z-index: 100;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .control-group {
            margin-bottom: 15px;
        }
        label {
            font-weight: bold;
            margin-right: 10px;
        }
        select, input {
            padding: 5px;
            font-size: 14px;
        }
        .groups-checkboxes {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
            margin-top: 10px;
        }
        .group-checkbox {
            display: flex;
            align-items: center;
            gap: 5px;
        }
        .color-picker {
            width: 40px;
            height: 25px;
            border: 1px solid #ccc;
            cursor: pointer;
        }
        table { 
            border-collapse: collapse; 
            margin: 20px 0;
            width: 100%;
        }
        td { 
            border: 1px solid #ccc; 
            padding: 5px; 
            min-width: 80px;
            max-width: 300px;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .header { background: #e0e0e0; font-weight: bold; }
        .day { background: #ffeb3b; }
        .number { background: #4caf50; color: white; }
        .week-chet { background: #2196f3; color: white; }
        .week-nech { background: #f44336; color: white; }
        .week-empty { background: #ffc107; }
        .time { background: #9c27b0; color: white; }
        .lesson { background: #fff; }
        .vmerge { border: 2px dashed red; }
        .colspan { border: 3px solid green; }
        .legend {
            margin: 20px 0;
            padding: 15px;
            background: #f5f5f5;
            border-radius: 5px;
        }
        .legend h3 { margin-top: 0; }
        .legend p { margin: 5px 0; }
    </style>
</head>
<body>
    <div class="controls">
        <h1>Интерактивная визуализация парсера</h1>
        
        <div class="control-group">
            <label>Выберите файл:</label>
            <select id="fileSelect" onchange="loadFile()">
                <option value="">-- Выберите файл --</option>
""" + ''.join([f'<option value="{name}">{name}</option>' for name in files_data.keys()]) + """
            </select>
        </div>
        
        <div class="control-group">
            <label>Выберите группы для подсветки:</label>
            <div id="groupsContainer" class="groups-checkboxes"></div>
        </div>
        
        <div class="control-group">
            <label>
                <input type="checkbox" id="showNormalized" onchange="toggleNormalized()">
                Показать нормализованные данные
            </label>
        </div>
    </div>
    
    <div class="legend">
        <h3>Легенда:</h3>
        <p><span style="background: #ffeb3b; padding: 5px;">День недели</span></p>
        <p><span style="background: #4caf50; color: white; padding: 5px;">Номер пары</span></p>
        <p><span style="background: #2196f3; color: white; padding: 5px;">Чет неделя</span></p>
        <p><span style="background: #f44336; color: white; padding: 5px;">Неч неделя</span></p>
        <p><span style="background: #ffc107; padding: 5px;">Пустая неделя</span></p>
        <p><span style="background: #9c27b0; color: white; padding: 5px;">Время</span></p>
        <p><span style="border: 2px dashed red; padding: 5px;">vMerge (объединение по вертикали)</span></p>
        <p><span style="border: 3px solid green; padding: 5px;">colspan > 1</span></p>
    </div>
    
    <div id="tableContainer"></div>
    
    <script>
        const filesData = """ + json.dumps(files_data, ensure_ascii=False) + """;
        let currentFile = null;
        let groupColors = {};
        
        function loadFile() {
            const fileName = document.getElementById('fileSelect').value;
            if (!fileName) return;
            
            currentFile = filesData[fileName];
            
            // Создаем чекбоксы для групп
            const container = document.getElementById('groupsContainer');
            container.innerHTML = '';
            
            const colors = ['#e1f5fe', '#f3e5f5', '#fff9c4', '#ffccbc', '#c8e6c9', '#ffe0b2'];
            
            currentFile.groups.forEach((group, idx) => {
                const div = document.createElement('div');
                div.className = 'group-checkbox';
                
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.id = 'group_' + group;
                checkbox.checked = true;
                checkbox.onchange = renderTable;
                
                const colorPicker = document.createElement('input');
                colorPicker.type = 'color';
                colorPicker.className = 'color-picker';
                colorPicker.value = colors[idx % colors.length].replace('#', '#');
                colorPicker.onchange = () => {
                    groupColors[group] = colorPicker.value;
                    renderTable();
                };
                
                groupColors[group] = colors[idx % colors.length];
                
                const label = document.createElement('label');
                label.htmlFor = 'group_' + group;
                label.textContent = group;
                
                div.appendChild(checkbox);
                div.appendChild(colorPicker);
                div.appendChild(label);
                container.appendChild(div);
            });
            
            renderTable();
        }
        
        function renderTable() {
            if (!currentFile) return;
            
            const showNormalized = document.getElementById('showNormalized').checked;
            const selectedGroups = currentFile.groups.filter(g => 
                document.getElementById('group_' + g)?.checked
            );
            
            let html = '<table>';
            
            // Нормализация
            let lastDay = null;
            let lastLesson = null;
            let lastWeek = null;
            const days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
            let currentDayIdx = -1;
            
            const timeMap = {
                '1': '8:00-9:30', '2': '9:40-11:10', '3': '11:45-13:15',
                '4': '13:50-15:20', '5': '15:30-17:00', '6': '17:10-18:40', '7': '18:50-20:20'
            };
            
            currentFile.rows.forEach((row, rowIdx) => {
                html += '<tr>';
                
                row.forEach((cell, cellIdx) => {
                    let cssClass = '';
                    let style = '';
                    let text = cell.text || '&nbsp;';
                    
                    if (rowIdx < 2) {
                        cssClass = 'header';
                    } else if (cellIdx === 0) {
                        // ДЕНЬ
                        cssClass = 'day';
                        if (cell.text) {
                            // Есть текст - определяем день
                            const lower = cell.text.toLowerCase();
                            for (let i = 0; i < days.length; i++) {
                                if (lower.includes(days[i].toLowerCase())) {
                                    currentDayIdx = i;
                                    lastDay = days[i];
                                    break;
                                }
                            }
                        } else if (showNormalized) {
                            // Пусто - проверяем не новый ли день
                            const currentLesson = row[1]?.text;
                            if (currentLesson && lastLesson && parseInt(currentLesson) < parseInt(lastLesson)) {
                                // Номер пары уменьшился - новый день
                                currentDayIdx++;
                                if (currentDayIdx < days.length) {
                                    lastDay = days[currentDayIdx];
                                }
                            }
                            text = lastDay || '?';
                        }
                    } else if (cellIdx === 1) {
                        // НОМЕР ПАРЫ
                        cssClass = 'number';
                        if (showNormalized && !cell.text) {
                            text = lastLesson || '?';
                        } else if (cell.text) {
                            lastLesson = cell.text;
                        }
                    } else if (cellIdx === 2) {
                        // НЕДЕЛЯ
                        const lower = cell.text.toLowerCase();
                        if (lower.includes('чет')) {
                            cssClass = 'week-chet';
                            lastWeek = 'Чет';
                        } else if (lower.includes('неч')) {
                            cssClass = 'week-nech';
                            lastWeek = 'Неч';
                        } else {
                            cssClass = 'week-empty';
                            if (showNormalized) {
                                if (lastWeek === 'Чет') {
                                    lastWeek = 'Неч';
                                    text = 'Неч';
                                    cssClass = 'week-nech';
                                } else if (lastWeek === 'Неч') {
                                    lastWeek = 'Чет';
                                    text = 'Чет';
                                    cssClass = 'week-chet';
                                } else {
                                    lastWeek = 'Чет';
                                    text = 'Чет';
                                    cssClass = 'week-chet';
                                }
                            }
                        }
                    } else if (cellIdx === 3) {
                        // ВРЕМЯ
                        cssClass = 'time';
                        if (showNormalized && lastLesson) {
                            text = timeMap[lastLesson] || text;
                        }
                    } else {
                        cssClass = 'lesson';
                        if (cell.group && selectedGroups.includes(cell.group)) {
                            style = `background: ${groupColors[cell.group]};`;
                        }
                    }
                    
                    if (cell.vmerge) cssClass += ' vmerge';
                    if (cell.colspan > 1) cssClass += ' colspan';
                    
                    const attrs = `class="${cssClass}" ${style ? `style="${style}"` : ''} ${cell.colspan > 1 ? `colspan="${cell.colspan}"` : ''}`;
                    html += `<td ${attrs} title="Row ${rowIdx}, Cell ${cellIdx}, LogCol ${cell.logical_col}">${text}</td>`;
                });
                
                html += '</tr>';
            });
            
            html += '</table>';
            document.getElementById('tableContainer').innerHTML = html;
        }
        
        function toggleNormalized() {
            renderTable();
        }
    </script>
</body>
</html>
"""

with open('parser_interactive.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("\n✅ Интерактивная визуализация создана: parser_interactive.html")
print("Открой файл в браузере!")
