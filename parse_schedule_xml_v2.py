"""
Новый XML парсер расписания - работает с физическими ячейками
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import json
from pathlib import Path
from typing import List, Dict, Optional, Tuple


class XMLScheduleParserV2:
    """Парсер расписания через XML - версия 2"""
    
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    DAYS_MAP = {
        'понедельник': 1,
        'вторник': 2,
        'среда': 3,
        'четверг': 4,
        'пятница': 5,
        'суббота': 6,
        'воскресенье': 7
    }
    
    TIME_BY_LESSON = {
        1: ('8:00:00', '9:30:00'),
        2: ('9:40:00', '11:10:00'),
        3: ('11:45:00', '13:15:00'),
        4: ('13:50:00', '15:20:00'),
        5: ('15:30:00', '17:00:00'),
        6: ('17:10:00', '18:40:00'),
        7: ('18:50:00', '20:20:00')
    }
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.lessons = []
        
    def parse(self) -> List[Dict]:
        """Основной метод парсинга"""
        try:
            with zipfile.ZipFile(self.file_path) as z:
                xml_content = z.read('word/document.xml')
            
            root = ET.fromstring(xml_content)
            tables = root.findall('.//w:tbl', self.NS)
            
            if not tables:
                return []
            
            # Берем самую большую таблицу (с наибольшим количеством строк)
            # Первая таблица часто содержит только заголовок документа
            table = max(tables, key=lambda t: len(t.findall('.//w:tr', self.NS)))
            rows = table.findall('.//w:tr', self.NS)
            
            if len(rows) < 3:
                return []
            
            # Определяем структуру заголовка
            # Если в первой строке есть "Дата", "Номер", "Неделя" - это заголовок в 1 строку
            # Иначе - заголовок в 2 строки (группы + подгруппы)
            first_row = rows[0]
            first_cells = first_row.findall('.//w:tc', self.NS)
            first_row_text = ' '.join([self._get_cell_text(c) for c in first_cells]).lower()
            
            if 'дата' in first_row_text and 'номер' in first_row_text:
                # Заголовок в 1 строку
                group_mapping = self._parse_header_physical([rows[0]])
                data_start = 1
            else:
                # Заголовок в 2 строки
                group_mapping = self._parse_header_physical(rows[:2])
                data_start = 2
            
            # Парсим данные
            self._parse_data_physical(rows[data_start:], group_mapping)
            
            return self.lessons
            
        except Exception as e:
            print(f"❌ Ошибка при парсинге {self.file_path.name}: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _parse_header_physical(self, header_rows: List[ET.Element]) -> Dict:
        """
        Парсит заголовок и возвращает маппинг:
        {
            физ_ячейка_номер: {
                'group': 'название группы',
                'subgroup': номер подгруппы или 0
            }
        }
        """
        mapping = {}
        
        # Первая строка - группы
        first_row = header_rows[0]
        first_cells = first_row.findall('.//w:tc', self.NS)
        
        # Находим где начинаются группы (после служебных колонок)
        service_cols = 4  # Дата, Номер, Недели, Время
        
        # Парсим группы
        group_cells = []
        for i in range(service_cols, len(first_cells)):
            cell = first_cells[i]
            text = self._get_cell_text(cell)
            group_name = self._extract_group_name(text)
            
            if group_name:
                colspan = self._get_colspan(cell)
                group_cells.append({
                    'physical_index': i,
                    'name': group_name,
                    'colspan': colspan
                })
        
        # Вторая строка - подгруппы (если есть)
        if len(header_rows) > 1:
            second_row = header_rows[1]
            second_cells = second_row.findall('.//w:tc', self.NS)
            
            # Для каждой группы находим её подгруппы
            subgroup_cell_idx = service_cols
            
            for group_info in group_cells:
                group_name = group_info['name']
                group_colspan = group_info['colspan']
                
                # Собираем подгруппы для этой группы
                subgroups_found = []
                cells_processed = 0
                
                while cells_processed < group_colspan and subgroup_cell_idx < len(second_cells):
                    cell = second_cells[subgroup_cell_idx]
                    text = self._get_cell_text(cell).lower()
                    colspan = self._get_colspan(cell)
                    
                    subgroup_num = 0
                    if 'подгруппа' in text:
                        match = re.search(r'(\d+)', text)
                        if match:
                            subgroup_num = int(match.group(1))
                    
                    subgroups_found.append({
                        'physical_index': subgroup_cell_idx,
                        'subgroup': subgroup_num,
                        'colspan': colspan
                    })
                    
                    cells_processed += colspan
                    subgroup_cell_idx += 1
                
                # Если подгрупп нет - вся группа
                if not subgroups_found or all(s['subgroup'] == 0 for s in subgroups_found):
                    mapping[group_info['physical_index']] = {
                        'group': group_name,
                        'subgroup': 0
                    }
                else:
                    # Добавляем маппинг для каждой подгруппы
                    for sg in subgroups_found:
                        mapping[sg['physical_index']] = {
                            'group': group_name,
                            'subgroup': sg['subgroup']
                        }
        else:
            # Нет подгрупп - заголовок в 1 строку
            for group_info in group_cells:
                mapping[group_info['physical_index']] = {
                    'group': group_info['name'],
                    'subgroup': 0
                }
        
        return mapping
    
    def _parse_data_physical(self, data_rows: List[ET.Element], group_mapping: Dict):
        """Парсит строки данных используя физические ячейки"""
        last_day = None
        last_lesson_number = None
        last_week_type = None
        
        # Создаем обратный маппинг: логическая колонка -> группа
        logical_col_to_groups = {}
        for phys_idx, info in group_mapping.items():
            # Нужно понять какие логические колонки занимает эта физическая ячейка
            # Для этого пересчитаем из заголовка
            pass
        
        # Лучше создать маппинг из заголовка заново
        # Получаем маппинг: логическая колонка -> список (group, subgroup)
        logical_col_to_group_info = self._build_logical_column_mapping(group_mapping)
        
        for row in data_rows:
            cells = row.findall('.//w:tc', self.NS)
            
            if len(cells) < 4:
                continue
            
            # Служебные колонки
            day_text = self._get_cell_text(cells[0]).strip()
            lesson_number_text = self._get_cell_text(cells[1]).strip()
            week_type_text = self._get_cell_text(cells[2]).strip()
            
            # ДЕНЬ
            if day_text:
                for day_name, day_num in self.DAYS_MAP.items():
                    if day_name in day_text.lower():
                        last_day = day_num
                        break
            current_day = last_day
            
            if not current_day:
                continue
            
            # НОМЕР ПАРЫ
            if lesson_number_text.isdigit():
                last_lesson_number = int(lesson_number_text)
            lesson_number = last_lesson_number
            
            if not lesson_number:
                continue
            
            # ТИП НЕДЕЛИ
            if week_type_text:
                week_text_lower = week_type_text.lower()
                if 'неч' in week_text_lower:
                    last_week_type = 'Неч'
                elif 'чет' in week_text_lower:
                    last_week_type = 'Чет'
                elif 'об' in week_text_lower:
                    last_week_type = 'Обе'
            else:
                # Чередование только если пусто
                if last_week_type == 'Чет':
                    last_week_type = 'Неч'
                elif last_week_type == 'Неч':
                    last_week_type = 'Чет'
                else:
                    last_week_type = 'Чет'
            
            week_type = last_week_type
            
            # Время
            time_start, time_end = self.TIME_BY_LESSON.get(lesson_number, ('00:00:00', '00:00:00'))
            
            # Парсим занятия - идем по физическим ячейкам начиная с 4
            logical_col = 4  # Первые 4 колонки служебные
            
            for cell_idx in range(4, len(cells)):
                cell = cells[cell_idx]
                lesson_text = self._get_cell_text(cell).strip()
                colspan = self._get_colspan(cell)
                
                if not lesson_text:
                    logical_col += colspan
                    continue
                
                # Определяем какие группы покрывает эта ячейка
                covered_groups = set()
                for i in range(colspan):
                    col = logical_col + i
                    if col in logical_col_to_group_info:
                        info = logical_col_to_group_info[col]
                        covered_groups.add((info['group'], info['subgroup']))
                
                lesson_data = self._parse_lesson_text(lesson_text)
                
                if lesson_data:
                    # Создаем занятие для каждой уникальной группы
                    for group_name, subgroup in covered_groups:
                        self.lessons.append({
                            'group_name': group_name,
                            'subgroup': subgroup,
                            'day_of_week': current_day,
                            'lesson_number': lesson_number,
                            'time_start': time_start,
                            'time_end': time_end,
                            'week_type': week_type,
                            'subject': lesson_data['subject'],
                            'type': lesson_data['type'],
                            'teacher': lesson_data['teacher'],
                            'room': lesson_data['room'],
                            'semester': 'Весенний',
                            'start_date': lesson_data.get('start_date'),
                            'end_date': lesson_data.get('end_date')
                        })
                
                logical_col += colspan
    
    def _build_logical_column_mapping(self, group_mapping: Dict) -> Dict:
        """
        Строит маппинг: логическая колонка -> {'group': ..., 'subgroup': ...}
        На основе физического маппинга и информации о colspan из заголовка
        """
        try:
            with zipfile.ZipFile(self.file_path) as z:
                xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)
            tables = root.findall('.//w:tbl', self.NS)
            
            # Берем самую большую таблицу
            table = max(tables, key=lambda t: len(t.findall('.//w:tr', self.NS)))
            rows = table.findall('.//w:tr', self.NS)
            
            # Определяем какая строка с группами
            first_row = rows[0]
            first_cells = first_row.findall('.//w:tc', self.NS)
            first_row_text = ' '.join([self._get_cell_text(c) for c in first_cells]).lower()
            
            if 'дата' in first_row_text and 'номер' in first_row_text:
                # Заголовок в 1 строку - группы в первой строке
                header_row = rows[0]
                subgroup_row = None
            else:
                # Заголовок в 2 строки
                header_row = rows[0]
                subgroup_row = rows[1] if len(rows) > 1 else None
            
            header_cells = header_row.findall('.//w:tc', self.NS)
            
            # Строим маппинг логических колонок
            result = {}
            
            if subgroup_row is not None:
                # Есть подгруппы - идем по второй строке
                subgroup_cells = subgroup_row.findall('.//w:tc', self.NS)
                
                logical_col = 0
                for cell_idx, cell in enumerate(subgroup_cells):
                    colspan = self._get_colspan(cell)
                    
                    # Определяем к какой группе относится эта подгруппа
                    if cell_idx in group_mapping:
                        group_info = group_mapping[cell_idx]
                        # Все логические колонки этой ячейки принадлежат этой группе/подгруппе
                        for i in range(colspan):
                            result[logical_col + i] = {
                                'group': group_info['group'],
                                'subgroup': group_info['subgroup']
                            }
                    
                    logical_col += colspan
            else:
                # Нет подгрупп - идем по первой строке
                logical_col = 0
                for cell_idx, cell in enumerate(header_cells):
                    colspan = self._get_colspan(cell)
                    
                    if cell_idx in group_mapping:
                        group_info = group_mapping[cell_idx]
                        for i in range(colspan):
                            result[logical_col + i] = {
                                'group': group_info['group'],
                                'subgroup': group_info['subgroup']
                            }
                    
                    logical_col += colspan
            
            return result
            
        except Exception as e:
            print(f"Ошибка при построении маппинга: {e}")
            import traceback
            traceback.print_exc()
            return {}
    
    def _get_cell_text(self, cell: ET.Element) -> str:
        texts = cell.findall('.//w:t', self.NS)
        return ''.join([t.text for t in texts if t.text])
    
    def _get_colspan(self, cell: ET.Element) -> int:
        gridSpan = cell.find('.//w:gridSpan', self.NS)
        if gridSpan is not None:
            colspan_val = gridSpan.get(f'{{{self.NS["w"]}}}val')
            if colspan_val:
                return int(colspan_val)
        return 1
    
    def _extract_group_name(self, text: str) -> Optional[str]:
        pattern = r'\b\d{1,2}[А-ЯA-Z]{2,4}\d{2}\b'
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(0) if match else None
    
    def _parse_lesson_text(self, text: str) -> Optional[Dict]:
        if not text:
            return None
        
        start_date = None
        end_date = None
        
        # Даты
        start_date_match = re.search(r'[Сс]\s+(\d{2})\.(\d{2})\.(\d{2})\.?', text)
        if start_date_match:
            day, month, year = start_date_match.groups()
            start_date = f"20{year}-{month}-{day}"
            text = text[:start_date_match.start()] + text[start_date_match.end():]
        else:
            start_date_match = re.search(r'[Сс]\s+(\d{2})\.(\d{2})', text)
            if start_date_match:
                day, month = start_date_match.groups()
                start_date = f"2026-{month}-{day}"
                text = text[:start_date_match.start()] + text[start_date_match.end():]
        
        end_date_match = re.search(r'[Пп]о\s+(\d{2})\.(\d{2})\.(\d{2})\.?', text)
        if end_date_match:
            day, month, year = end_date_match.groups()
            end_date = f"20{year}-{month}-{day}"
            text = text[:end_date_match.start()] + text[end_date_match.end():]
        else:
            end_date_match = re.search(r'[Пп]о\s+(\d{2})\.(\d{2})', text)
            if end_date_match:
                day, month = end_date_match.groups()
                end_date = f"2026-{month}-{day}"
                text = text[:end_date_match.start()] + text[end_date_match.end():]
        
        subject = text.strip()
        subject = re.sub(r'^\s*\.\s*', '', subject)
        lesson_type = 'Практические'
        
        type_match = re.search(r'\((.*?)\)', text)
        if type_match:
            type_text = type_match.group(1).lower()
            if 'лекц' in type_text:
                lesson_type = 'Лекции'
            elif 'практ' in type_text:
                lesson_type = 'Практические'
            elif 'лаб' in type_text:
                lesson_type = 'Лабораторные'
            
            subject = text[:type_match.start()].strip()
        
        teacher = None
        # Улучшенный паттерн: звание должно быть после закрывающей скобки типа занятия
        # Формат: ) + любой текст + доц./проф./асс./ст.преп./ст.пр. + Фамилия + И. + О.
        # Ищем первого преподавателя (может быть список через запятую)
        teacher_pattern = r'\).*?(?:доц\.?|проф\.?|асс\.?|ст\.?\s*(?:пр(?:еп)?|преп)\.?)\s*([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?)?)'
        teacher_match = re.search(teacher_pattern, text, re.IGNORECASE)
        if teacher_match:
            teacher = teacher_match.group(1).strip()
        
        room = None
        # Улучшенный паттерн для кабинета
        # Сначала ищем специальные кабинеты (СК, НОЦ) после закрывающей скобки
        special_room_pattern = r'\)\s*([СC]К\s+[«"]?[А-Яа-я]+[»"]?|НОЦ\s+[А-Яа-я]+)'
        special_match = re.search(special_room_pattern, text)
        if special_match:
            room = special_match.group(1).strip()
        else:
            # Обычные кабинеты: цифры-цифры с буквой или просто цифры
            room_pattern = r'(\d{1,2}-\d{2,4}[а-яА-Яa-zA-Z]*|\d{1,4}[а-яА-Я]?)'
            room_matches = re.findall(room_pattern, text)
            if room_matches:
                room = room_matches[-1] if room_matches else None
        
        return {
            'subject': subject,
            'type': lesson_type,
            'teacher': teacher,
            'room': room,
            'start_date': start_date,
            'end_date': end_date
        }


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        group_filter = sys.argv[2] if len(sys.argv) > 2 else None
        
        parser = XMLScheduleParserV2(file_path)
        lessons = parser.parse()
        
        if group_filter:
            lessons = [l for l in lessons if l['group_name'] == group_filter]
        
        print(f"✅ Спарсено {len(lessons)} занятий")
        
        with open('parsed_v2.json', 'w', encoding='utf-8') as f:
            json.dump(lessons, f, ensure_ascii=False, indent=2)
