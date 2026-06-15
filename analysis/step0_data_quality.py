"""
0-BOSQICH: Ma'lumot sifati tahlili va chang bo'roni kunlarini aniqlash
=====================================================================
Bu skript barcha stansiyalar bo'yicha:
1. Ma'lumot sifatini tekshiradi (bo'sh qiymatlar, formatlar)
2. Chang bo'roni kunlarini aniqlaydi (V, shamol, namlik asosida)
3. Stansiyalar bo'yicha statistika chiqaradi
4. Natijalarni CSV formatida saqlaydi

Ishlatish: python analysis/step0_data_quality.py
"""

import sys
import os
import csv
from collections import defaultdict
import xlrd

# ============================================================
# SOZLAMALAR — yo'llarni o'z kompyuteringizga moslang!
# ============================================================

# Loyiha papkasi (step0_data_quality.py yoki notebook joylashgan papkadan bir qadam yuqori)
PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Ma'lumotlar papkasi
DATA_DIR = os.path.join(PROJECT_DIR, "data", "все станции за 10 лет")

# Natijalar saqlanadigan papka
OUTPUT_DIR = os.path.join(PROJECT_DIR, "analysis", "results")

# ─────────────────────────────────────────────────────────────
# AGAR YUQORIDAGI YO'LLAR ISHLAMASA, quyidagi qatorlardan birini
# izohdan chiqarib, o'z yo'lingizni yozing:
#
# DATA_DIR = "C:/Users/SizningNom/dust_storm_all/data/все станции за 10 лет"
# DATA_DIR = "/home/user/dust_storm_all/data/все станции за 10 лет"
#
# OUTPUT_DIR = "C:/Users/SizningNom/dust_storm_all/analysis/results"
# OUTPUT_DIR = "/home/user/dust_storm_all/analysis/results"
# ─────────────────────────────────────────────────────────────

# Chang bo'roni aniqlash mezonlari
# V (ko'rinish) - km da (tasdiqlangan: JASLYK max=23.6, NUKUS max=11.8)
# Vx8 - shamol tezligi (m/s), VxG - shamol loli/gust (m/s)
# U/UN - nisbiy namlik (%)
#
# MUHIM: Tashkentda V o'rtacha 1.28 km (tuman, smog ta'sirida) — 
# shuning uchun faqat V bilan emas, V + shamol + quruqlik birgalikda aniqlaymiz!
#
# Mezon: chang bo'roni = past ko'rinish + KUCHLI shamol + QURUQ havo
# (tuman = past ko'rinish + past shamol + yuqori namlik — bu EMAS)

DUST_V_SEVERE = 0.5          # V < 0.5 km — juda kuchli chang
DUST_V_STRONG = 1.0          # V < 1 km — kuchli chang
DUST_V_MODERATE = 2.0        # V < 2 km — o'rtacha
DUST_WIND_THRESHOLD = 10.0   # VxG >= 10 m/s (chang uchun kerak)
DUST_WIND_MODERATE = 8.0     # VxG >= 8 m/s (o'rtacha)
DUST_HUMIDITY_LOW = 40.0     # UN < 40% — quruq havo (chang uchun)
DUST_HUMIDITY_MODERATE = 60.0  # UN < 60%

# Standart ustun nomlari (aksariyat stansiyalarda)
STANDARD_HEADERS = [
    'Year', 'Mon', 'Day', 'Taav', 'TaX', 'TaN', 'Tg', 'TgX', 'TgN', 'TrN',
    'E', 'U', 'UN', 'Ed', 'EdX', 'StP', 'SeP', 'Lo', 'Ln', 'V',
    'Vx8', 'VxG', 'RSum', 'R', 'Sc', 'Sh', 'Tg5', 'Tg10', 'Tg15', 'Tg20',
    'Te2', 'Te5', 'Te10', 'Te15', 'Te20', 'Te40', 'SunL', 'Tr'
]

# Chang bo'roni modeli uchun eng muhim parametrlar
KEY_PARAMS = ['V', 'Vx8', 'VxG', 'U', 'UN', 'Ed', 'Taav', 'Tg', 'StP', 'SeP']


# ============================================================
# YORDAMCHI FUNKSIYALAR
# ============================================================

def find_header_row(sheet):
    """Header qatorini topish — 'Year' so'zini qidirish"""
    for row_idx in range(min(5, sheet.nrows)):
        for col_idx in range(min(5, sheet.ncols)):
            cell_val = str(sheet.cell(row_idx, col_idx).value).strip()
            if cell_val == 'Year':
                return row_idx
    return -1


def get_headers(sheet, header_row):
    """Ustun nomlarini olish"""
    if header_row < 0:
        return []
    return [str(sheet.cell(header_row, col).value).strip() for col in range(sheet.ncols)]


def safe_float(value):
    """Qiymatni xavfsiz tarzda floatga aylantirish"""
    if value == '' or value == '*' or value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def read_station_data(filepath):
    """
    Bitta stansiya faylini o'qish va tuzilgan ma'lumot qaytarish.
    
    Returns:
        dict: {
            'name': stansiya nomi,
            'headers': ustun nomlari,
            'header_row': header qator indeksi,
            'data_rows': ma'lumot qatorlari soni,
            'years': (boshlanish, tugash),
            'data': [{ustun: qiymat, ...}, ...],
            'error': xato xabari (agar bor bo'lsa)
        }
    """
    result = {
        'name': os.path.basename(filepath),
        'headers': [],
        'header_row': -1,
        'data_rows': 0,
        'years': (None, None),
        'data': [],
        'error': None
    }
    
    try:
        wb = xlrd.open_workbook(filepath)
        sheet = wb.sheet_by_index(0)
        
        header_row = find_header_row(sheet)
        result['header_row'] = header_row
        
        if header_row >= 0:
            headers = get_headers(sheet, header_row)
            result['headers'] = headers
            data_start = header_row + 1
        else:
            # Header topilmadi — birinchi qatordan data boshlaydi deb hisoblaymiz
            # va standart header ishlatamiz
            headers = STANDARD_HEADERS[:sheet.ncols]
            result['headers'] = headers
            data_start = 0
        
        result['data_rows'] = sheet.nrows - data_start
        
        # Yillar oralig'ini aniqlash
        first_year = safe_float(sheet.cell(data_start, 0).value)
        last_year = safe_float(sheet.cell(sheet.nrows - 1, 0).value)
        result['years'] = (int(first_year) if first_year else None, 
                          int(last_year) if last_year else None)
        
        # Ma'lumotni o'qish
        for row_idx in range(data_start, sheet.nrows):
            row_dict = {}
            for col_idx, header in enumerate(headers):
                if col_idx < sheet.ncols:
                    row_dict[header] = sheet.cell(row_idx, col_idx).value
            result['data'].append(row_dict)
            
    except Exception as e:
        result['error'] = str(e)
    
    return result


# ============================================================
# 1. MA'LUMOT SIFATI TAHLILI
# ============================================================

def analyze_data_quality(station_data):
    """
    Stansiya ma'lumotlari sifatini tahlil qilish.
    
    Returns:
        dict: {
            'total_rows': jami qatorlar,
            'missing_by_column': {ustun: bo'sh qiymatlar soni},
            'missing_percent': {ustun: bo'sh qiymatlar foizi},
            'value_ranges': {ustun: (min, max, mean)},
            'key_params_available': True/False
        }
    """
    quality = {
        'total_rows': len(station_data['data']),
        'missing_by_column': {},
        'missing_percent': {},
        'value_ranges': {},
        'key_params_available': True
    }
    
    headers = station_data['headers']
    data = station_data['data']
    
    if not data:
        return quality
    
    for header in headers:
        if not header:
            continue
            
        values = []
        missing = 0
        
        for row in data:
            val = row.get(header, '')
            num = safe_float(val)
            if num is not None:
                values.append(num)
            else:
                missing += 1
        
        quality['missing_by_column'][header] = missing
        quality['missing_percent'][header] = round(missing / len(data) * 100, 1) if data else 0
        
        if values:
            quality['value_ranges'][header] = (
                round(min(values), 2),
                round(max(values), 2),
                round(sum(values) / len(values), 2)
            )
    
    # Asosiy parametrlar mavjudligini tekshirish
    for param in KEY_PARAMS:
        if param not in headers:
            quality['key_params_available'] = False
            break
    
    return quality


# ============================================================
# 2. CHANG BO'RONI KUNLARINI ANIQLASH
# ============================================================

def detect_dust_storms(station_data):
    """
    Chang bo'roni kunlarini aniqlash.
    
    Asosiy mantiq: CHANG BO'RONI = past ko'rinish + kuchli shamol + quruq havo
    TUMAN (false positive) = past ko'rinish + past shamol + yuqori namlik
    
    Mezonlar:
    - KUCHLI: V < 0.5 km VA VxG >= 10 VA UN < 60%
    - ORTACHA: V < 1 km VA (VxG >= 8 YOKI UN < 40%)  
    - YENGIL: V < 2 km VA VxG >= 10 VA UN < 40%
    - TUMAN (chiqarib tashlanadi): V past VA shamol past VA namlik > 70%
    
    Returns:
        list: [{
            'date': '2011-03-15',
            'severity': 'KUCHLI'/'ORTACHA'/'YENGIL',
            'V': ko'rinish qiymati,
            'VxG': shamol tezligi,
            'UN': minimal namlik,
            'Taav': harorat,
            'is_fog': True/False
        }, ...]
    """
    dust_events = []
    data = station_data['data']
    headers = station_data['headers']
    
    if 'V' not in headers:
        return dust_events
    
    for row in data:
        v = safe_float(row.get('V', ''))
        vxg = safe_float(row.get('VxG', ''))
        vx8 = safe_float(row.get('Vx8', ''))
        un = safe_float(row.get('UN', ''))
        u = safe_float(row.get('U', ''))
        taav = safe_float(row.get('Taav', ''))
        
        year = safe_float(row.get('Year', ''))
        mon = safe_float(row.get('Mon', ''))
        day = safe_float(row.get('Day', ''))
        
        if v is None or year is None or mon is None or day is None:
            continue
        
        date_str = f"{int(year)}-{int(mon):02d}-{int(day):02d}"
        
        # Ko'rinish past bo'lmasa — chang bo'roni emas
        if v >= DUST_V_MODERATE:
            continue
        
        # Shamol tezligi (VxG yoki Vx8 dan eng yuqori)
        wind = vxg if vxg is not None else (vx8 if vx8 is not None else None)
        
        # Namlik (UN - minimal, U - o'rtacha)
        humidity = un if un is not None else u
        
        # TUMAN FILTRI: past ko'rinish + past shamol + yuqori namlik = TUMAN, CHANG EMAS
        is_fog = False
        if humidity is not None and humidity > 70:
            if wind is not None and wind < 5:
                is_fog = True
            elif wind is None:
                is_fog = True
        
        if is_fog:
            continue  # Tumanli kunlarni chiqarib tashlaymiz
        
        severity = None
        
        # KUCHLI chang bo'roni: V < 0.5 km VA kuchli shamol VA quruq
        if v < DUST_V_SEVERE:
            if wind is not None and wind >= DUST_WIND_THRESHOLD:
                severity = 'KUCHLI'
            elif humidity is not None and humidity < DUST_HUMIDITY_LOW:
                severity = 'KUCHLI'
            else:
                severity = 'ORTACHA'  # V juda past, lekin shamol/namlik ma'lumoti yo'q
        
        # ORTACHA: V < 1 km VA (shamol YOKI quruqlik)
        elif v < DUST_V_STRONG:
            if wind is not None and wind >= DUST_WIND_MODERATE:
                severity = 'ORTACHA'
            elif humidity is not None and humidity < DUST_HUMIDITY_LOW:
                severity = 'ORTACHA'
            else:
                severity = 'YENGIL'
        
        # YENGIL: 1 <= V < 2 km VA shamol VA quruqlik
        elif v < DUST_V_MODERATE:
            if (wind is not None and wind >= DUST_WIND_THRESHOLD) and \
               (humidity is not None and humidity < DUST_HUMIDITY_LOW):
                severity = 'YENGIL'
        
        if severity:
            dust_events.append({
                'date': date_str,
                'severity': severity,
                'V': v,
                'VxG': vxg,
                'Vx8': vx8,
                'UN': un,
                'Taav': taav,
                'is_fog': False
            })
    
    return dust_events


# ============================================================
# 3. STATISTIKA VA HISOBOT
# ============================================================

def generate_statistics(all_stations_events):
    """
    Barcha stansiyalar bo'yicha chang bo'roni statistikasini hisoblash.
    """
    stats = {
        'total_events': 0,
        'kuchli': 0,
        'ortacha': 0,
        'yengil': 0,
        'by_station': {},
        'by_year': defaultdict(int),
        'by_month': defaultdict(int),
        'by_station_year': defaultdict(lambda: defaultdict(int)),
        'top_stations': [],
        'seasonal_pattern': {}
    }
    
    for station_name, events in all_stations_events.items():
        station_count = {'KUCHLI': 0, 'ORTACHA': 0, 'YENGIL': 0, 'total': 0}
        
        for event in events:
            stats['total_events'] += 1
            severity = event['severity']
            
            if severity == 'KUCHLI':
                stats['kuchli'] += 1
                station_count['KUCHLI'] += 1
            elif severity == 'ORTACHA':
                stats['ortacha'] += 1
                station_count['ORTACHA'] += 1
            else:
                stats['yengil'] += 1
                station_count['YENGIL'] += 1
            
            station_count['total'] += 1
            
            # Yil va oy bo'yicha
            year = event['date'][:4]
            month = event['date'][5:7]
            stats['by_year'][year] += 1
            stats['by_month'][month] += 1
            stats['by_station_year'][station_name][year] += 1
        
        stats['by_station'][station_name] = station_count
    
    # Eng ko'p chang bo'roni bo'lgan stansiyalar (Top-20)
    sorted_stations = sorted(stats['by_station'].items(), 
                            key=lambda x: x[1]['KUCHLI'], reverse=True)
    stats['top_stations'] = sorted_stations[:20]
    
    return stats


# ============================================================
# ASOSIY ISHGA TUSHIRISH
# ============================================================

def main():
    print("=" * 80)
    print("  0-BOSQICH: MA'LUMOT SIFATI TAHLILI VA CHANG BO'RONI ANIQLASH")
    print("=" * 80)
    
    # Papkalar yaratish
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Barcha fayllarni o'qish
    if not os.path.exists(DATA_DIR):
        print(f"\n❌ XATO: Data papkasi topilmadi: {DATA_DIR}")
        print("  Iltimos, ZIP faylni 'data/' papkasiga oching.")
        return
    
    files = sorted([f for f in os.listdir(DATA_DIR) if f.endswith(('.xls', '.XLS', '.xlsx'))])
    print(f"\n📁 Topilgan fayllar: {len(files)} ta stansiya")
    print("-" * 80)
    
    all_quality = {}
    all_events = {}
    skipped_files = []
    
    for i, fname in enumerate(files, 1):
        filepath = os.path.join(DATA_DIR, fname)
        station_name = os.path.splitext(fname)[0].strip()
        
        print(f"  [{i:2d}/{len(files)}] {station_name}...", end=" ")
        
        # Ma'lumotni o'qish
        station_data = read_station_data(filepath)
        
        if station_data['error']:
            print(f"❌ XATO: {station_data['error']}")
            skipped_files.append((fname, station_data['error']))
            continue
        
        # Sifat tahlili
        quality = analyze_data_quality(station_data)
        all_quality[station_name] = quality
        
        # Chang bo'roni aniqlash
        events = detect_dust_storms(station_data)
        all_events[station_name] = events
        
        # Qisqa natija
        kuchli = len([e for e in events if e['severity'] == 'KUCHLI'])
        ortacha = len([e for e in events if e['severity'] == 'ORTACHA'])
        v_missing = quality['missing_percent'].get('V', 100)
        
        print(f"✓ {station_data['data_rows']} kun | "
              f"V bo'sh: {v_missing:.0f}% | "
              f"Chang: {kuchli} kuchli, {ortacha} o'rtacha")
    
    # ============================================================
    # STATISTIKA HISOBLASH
    # ============================================================
    
    print("\n" + "=" * 80)
    print("  UMUMIY STATISTIKA")
    print("=" * 80)
    
    stats = generate_statistics(all_events)
    
    print(f"\n📊 CHANG BO'RONI HODISALARI (tuman filtrlangan):")
    print(f"   Jami hodisalar: {stats['total_events']}")
    print(f"   🔴 Kuchli (V < 0.5 km + shamol/quruqlik): {stats['kuchli']}")
    print(f"   🟠 O'rtacha (V < 1 km + shamol/quruqlik): {stats['ortacha']}")
    print(f"   🟡 Yengil (V < 2 km + shamol + quruqlik): {stats['yengil']}")
    
    print(f"\n📅 YILLAR BO'YICHA:")
    for year in sorted(stats['by_year'].keys()):
        count = stats['by_year'][year]
        bar = "█" * (count // 20)
        print(f"   {year}: {count:4d} hodisa {bar}")
    
    print(f"\n📅 OYLAR BO'YICHA (mavsumiylik):")
    month_names = ['Yan', 'Fev', 'Mar', 'Apr', 'May', 'Iyn', 
                   'Iyl', 'Avg', 'Sen', 'Okt', 'Noy', 'Dek']
    for m_idx, m_name in enumerate(month_names, 1):
        m_key = f"{m_idx:02d}"
        count = stats['by_month'].get(m_key, 0)
        bar = "█" * (count // 30)
        print(f"   {m_name}: {count:4d} {bar}")
    
    print(f"\n🏆 ENG KO'P CHANG BO'RONI BO'LGAN STANSIYALAR (Top-20):")
    ortacha_label = "O'rtacha"
    print(f"   {'Stansiya':<25} {'Kuchli':>8} {ortacha_label:>10} {'Yengil':>9} {'Jami':>8}")
    print(f"   {'-'*62}")
    for station_name, counts in stats['top_stations']:
        print(f"   {station_name:<25} {counts['KUCHLI']:>8} {counts['ORTACHA']:>10} "
              f"{counts['YENGIL']:>9} {counts['total']:>8}")
    
    # ============================================================
    # MA'LUMOT SIFATI XULOSASI
    # ============================================================
    
    print(f"\n\n{'=' * 80}")
    print("  MA'LUMOT SIFATI XULOSASI")
    print("=" * 80)
    
    # Asosiy parametrlar bo'yicha bo'sh qiymatlar
    print(f"\n📋 ASOSIY PARAMETRLAR BO'YICHA BO'SH QIYMATLAR (o'rtacha %):")
    param_missing_avg = defaultdict(list)
    for station_name, quality in all_quality.items():
        for param in KEY_PARAMS:
            if param in quality['missing_percent']:
                param_missing_avg[param].append(quality['missing_percent'][param])
    
    for param in KEY_PARAMS:
        if param_missing_avg[param]:
            avg = sum(param_missing_avg[param]) / len(param_missing_avg[param])
            max_val = max(param_missing_avg[param])
            bar = "█" * int(avg / 2)
            print(f"   {param:<6}: o'rtacha {avg:5.1f}% | max {max_val:5.1f}% {bar}")
    
    # Eng muammoli stansiyalar
    print(f"\n⚠️  ENG KO'P BO'SH QIYMATLI STANSIYALAR (V ustuni bo'yicha):")
    v_missing_list = [(name, q['missing_percent'].get('V', 100)) 
                      for name, q in all_quality.items()]
    v_missing_list.sort(key=lambda x: x[1], reverse=True)
    for name, pct in v_missing_list[:10]:
        if pct > 0:
            print(f"   {name:<25}: V bo'sh {pct:.1f}%")
    
    if skipped_files:
        print(f"\n❌ O'QILMAGAN FAYLLAR ({len(skipped_files)}):")
        for fname, error in skipped_files:
            print(f"   {fname}: {error}")
    
    # ============================================================
    # NATIJALARNI FAYLGA SAQLASH
    # ============================================================
    
    # 1. Chang bo'roni hodisalari CSV
    events_file = os.path.join(OUTPUT_DIR, 'dust_storm_events.csv')
    with open(events_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['station', 'date', 'severity', 'V_km', 'VxG_ms', 'Vx8_ms', 'UN_pct', 'Taav_C'])
        
        for station_name, events in sorted(all_events.items()):
            for event in events:
                writer.writerow([
                    station_name,
                    event['date'],
                    event['severity'],
                    event['V'],
                    event['VxG'] if event['VxG'] is not None else '',
                    event['Vx8'] if event['Vx8'] is not None else '',
                    event['UN'] if event['UN'] is not None else '',
                    event['Taav'] if event['Taav'] is not None else ''
                ])
    
    print(f"\n\n💾 SAQLANGAN FAYLLAR:")
    print(f"   ✓ {events_file}")
    
    # 2. Stansiyalar sifati CSV
    quality_file = os.path.join(OUTPUT_DIR, 'station_quality.csv')
    with open(quality_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['station', 'total_rows', 'V_missing_%', 'VxG_missing_%', 
                        'UN_missing_%', 'V_min', 'V_max', 'V_mean',
                        'dust_kuchli', 'dust_ortacha', 'dust_yengil'])
        
        for station_name in sorted(all_quality.keys()):
            q = all_quality[station_name]
            events = all_events.get(station_name, [])
            kuchli = len([e for e in events if e['severity'] == 'KUCHLI'])
            ortacha = len([e for e in events if e['severity'] == 'ORTACHA'])
            yengil = len([e for e in events if e['severity'] == 'YENGIL'])
            
            v_range = q['value_ranges'].get('V', (None, None, None))
            
            writer.writerow([
                station_name,
                q['total_rows'],
                q['missing_percent'].get('V', ''),
                q['missing_percent'].get('VxG', ''),
                q['missing_percent'].get('UN', ''),
                v_range[0] if v_range[0] is not None else '',
                v_range[1] if v_range[1] is not None else '',
                v_range[2] if v_range[2] is not None else '',
                kuchli,
                ortacha,
                yengil
            ])
    
    print(f"   ✓ {quality_file}")
    
    # 3. Yillik/oylik statistika CSV
    stats_file = os.path.join(OUTPUT_DIR, 'dust_storm_stats.csv')
    with open(stats_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['type', 'period', 'count'])
        
        for year in sorted(stats['by_year'].keys()):
            writer.writerow(['year', year, stats['by_year'][year]])
        
        for month in sorted(stats['by_month'].keys()):
            writer.writerow(['month', month, stats['by_month'][month]])
    
    print(f"   ✓ {stats_file}")
    
    # 4. Umumiy hisobot TXT
    report_file = os.path.join(OUTPUT_DIR, 'analysis_report.txt')
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("  CHANG BO'RONI MA'LUMOTLARI TAHLILI — 0-BOSQICH HISOBOTI\n")
        f.write(f"  Stansiyalar: {len(all_quality)} ta | Davr: 2011-2020\n")
        f.write("=" * 80 + "\n\n")
        
        f.write("UMUMIY STATISTIKA:\n")
        f.write(f"  Jami chang bo'roni hodisalari: {stats['total_events']}\n")
        f.write(f"  Kuchli (V < 0.5 km + shamol/quruqlik): {stats['kuchli']}\n")
        f.write(f"  O'rtacha (V < 1 km + shamol/quruqlik): {stats['ortacha']}\n")
        f.write(f"  Yengil (V < 2 km + shamol + quruqlik): {stats['yengil']}\n\n")
        
        f.write("YILLAR BO'YICHA:\n")
        for year in sorted(stats['by_year'].keys()):
            f.write(f"  {year}: {stats['by_year'][year]} hodisa\n")
        
        f.write("\nOYLAR BO'YICHA (mavsumiylik):\n")
        for m_idx, m_name in enumerate(month_names, 1):
            m_key = f"{m_idx:02d}"
            count = stats['by_month'].get(m_key, 0)
            f.write(f"  {m_name}: {count} hodisa\n")
        
        f.write("\nTOP-20 STANSIYALAR (kuchli chang bo'roni bo'yicha):\n")
        f.write(f"  {'Stansiya':<25} {'Kuchli':>8} {'Ortacha':>8} {'Jami':>8}\n")
        f.write(f"  {'-'*55}\n")
        for station_name, counts in stats['top_stations']:
            f.write(f"  {station_name:<25} {counts['KUCHLI']:>8} "
                   f"{counts['ORTACHA']:>8} {counts['total']:>8}\n")
        
        f.write("\n\nMEZONLAR (ishlatilgan — tuman filtrlangan):\n")
        f.write(f"  KUCHLI chang bo'roni: V < {DUST_V_SEVERE} km VA (VxG >= {DUST_WIND_THRESHOLD} m/s YOKI UN < {DUST_HUMIDITY_LOW}%)\n")
        f.write(f"  O'RTACHA: V < {DUST_V_STRONG} km VA (VxG >= {DUST_WIND_MODERATE} m/s YOKI UN < {DUST_HUMIDITY_LOW}%)\n")
        f.write(f"  YENGIL: V < {DUST_V_MODERATE} km VA VxG >= {DUST_WIND_THRESHOLD} m/s VA UN < {DUST_HUMIDITY_LOW}%\n")
        f.write(f"  TUMAN FILTRI: V past VA shamol < 5 m/s VA namlik > 70% → chiqarib tashlanadi\n")
        f.write(f"\nBIRLIKLAR: V = km, Vx8/VxG = m/s, U/UN = %\n")
    
    print(f"   ✓ {report_file}")
    
    print(f"\n\n{'=' * 80}")
    print("  ✅ TAHLIL YAKUNLANDI!")
    print(f"  Natijalar: {OUTPUT_DIR}/")
    print("=" * 80)


if __name__ == '__main__':
    main()
