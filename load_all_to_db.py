"""
Загрузка всех расписаний в базу данных Supabase
"""

import os
import glob
from parse_schedule_xml_v2 import XMLScheduleParserV2
from supabase import create_client
from dotenv import load_dotenv

# Загружаем переменные окружения
env_path = 'тест кгасу/classmate-connect/.env'
load_dotenv(env_path)

SUPABASE_URL = os.getenv('VITE_SUPABASE_URL')
SUPABASE_KEY = os.getenv('VITE_SUPABASE_SERVICE_ROLE_KEY')

if not SUPABASE_URL or not SUPABASE_KEY:
    print("❌ Не найдены переменные окружения VITE_SUPABASE_URL или VITE_SUPABASE_SERVICE_ROLE_KEY")
    print(f"Проверьте файл: {env_path}")
    exit(1)

print(f"✅ Подключение к Supabase: {SUPABASE_URL}")

# Создаем клиент
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# Получаем все DOCX файлы
docx_files = glob.glob('schedules/*.docx')
docx_files.sort()

print(f"\n📁 Найдено {len(docx_files)} DOCX файлов")

# Сначала очищаем таблицу lessons_test
print("\n🗑️  Очищаем таблицу lessons_test...")
try:
    # Удаляем все записи
    result = supabase.table('lessons_test').delete().neq('id', 0).execute()
    print(f"✅ Таблица очищена")
except Exception as e:
    print(f"⚠️  Ошибка при очистке: {e}")

# Парсим все файлы
all_lessons = []
files_parsed = 0
files_skipped = 0

print("\n📝 Парсинг файлов...")

for file_path in docx_files:
    file_name = file_path.replace('schedules/', '').replace('schedules\\', '')
    
    parser = XMLScheduleParserV2(file_path)
    lessons = parser.parse()
    
    if lessons:
        files_parsed += 1
        all_lessons.extend(lessons)
        print(f"  ✅ {file_name}: {len(lessons)} занятий")
    else:
        files_skipped += 1
        print(f"  ⚠️  {file_name}: пропущен (0 занятий)")

print(f"\n📊 Статистика парсинга:")
print(f"  Файлов обработано: {files_parsed}")
print(f"  Файлов пропущено: {files_skipped}")
print(f"  Всего занятий: {len(all_lessons)}")

# Группируем по группам для статистики
from collections import defaultdict
by_group = defaultdict(int)
for lesson in all_lessons:
    by_group[lesson['group_name']] += 1

print(f"  Уникальных групп: {len(by_group)}")

# Загружаем в базу данных батчами
print(f"\n💾 Загрузка в базу данных...")

BATCH_SIZE = 100
total_inserted = 0
errors = 0

for i in range(0, len(all_lessons), BATCH_SIZE):
    batch = all_lessons[i:i+BATCH_SIZE]
    
    try:
        # Подготавливаем данные для вставки
        # Нужно получить group_id из таблицы groups
        batch_with_ids = []
        
        for lesson in batch:
            # Получаем group_id
            group_result = supabase.table('groups').select('id').eq('name', lesson['group_name']).execute()
            
            if not group_result.data:
                print(f"⚠️  Группа {lesson['group_name']} не найдена в базе")
                continue
            
            group_id = group_result.data[0]['id']
            
            batch_with_ids.append({
                'group_id': group_id,
                'subgroup': lesson['subgroup'],
                'subject': lesson['subject'],
                'type': lesson['type'],
                'teacher': lesson['teacher'],
                'room': lesson['room'],
                'day_of_week': lesson['day_of_week'],
                'lesson_number': lesson['lesson_number'],
                'time_start': lesson['time_start'],
                'time_end': lesson['time_end'],
                'week_type': lesson['week_type'],
                'semester': lesson['semester'],
                'start_date': lesson.get('start_date'),
                'end_date': lesson.get('end_date')
            })
        
        if batch_with_ids:
            result = supabase.table('lessons_test').insert(batch_with_ids).execute()
            total_inserted += len(batch_with_ids)
            print(f"  ✅ Загружено {total_inserted}/{len(all_lessons)} занятий")
    
    except Exception as e:
        errors += 1
        print(f"  ❌ Ошибка при загрузке батча {i//BATCH_SIZE + 1}: {e}")

print(f"\n{'='*60}")
print(f"✅ ЗАГРУЗКА ЗАВЕРШЕНА")
print(f"  Всего загружено: {total_inserted} занятий")
print(f"  Ошибок: {errors}")
print(f"{'='*60}")
