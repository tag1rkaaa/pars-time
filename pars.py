import pandas as pd
import re
from pathlib import Path
from typing import Optional, Dict, List
from dataclasses import dataclass
import tkinter as tk
from tkinter import filedialog


@dataclass
class SpeakerStat:
    name: str
    total_seconds: int = 0
    segment_count: int = 0  # Оставляем для внутренней логики, но не выводим

    @property
    def formatted_time(self) -> str:
        h, rem = divmod(self.total_seconds, 3600)
        m, s = divmod(rem, 60)
        if h: return f"{h}ч {m}м {s}с"
        if m: return f"{m}м {s}с"
        return f"{s}с"


def parse_chronometry(filepath: str,
                      output_file: Optional[str] = None,
                      max_silence_sec: int = 90,
                      similarity_threshold: float = 0.85) -> Dict[str, SpeakerStat]:
    """
    Парсит хронометраж спикеров из Excel-файла
    """
    print(f" Загрузка файла: {filepath}")

    # Загрузка Excel
    df = pd.read_excel(filepath, engine='openpyxl', header=None)
    df = df.dropna(how='all')

    # Парсинг сегментов
    segments = []
    for _, row in df.iterrows():
        if len(row) < 2:
            continue

        ts = str(row.iloc[0]).strip() if len(row) > 0 else ""
        speaker = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else "Неизвестный"

        # Пропускаем пустые строки и заголовки
        if not ts or speaker == "nan" or "Транскрипция" in ts:
            continue

        # Парсинг времени (формат H:M:S или M:S)
        match = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$', ts)
        if not match:
            continue

        seconds = int(match.group(1)) * 3600 + int(match.group(2)) * 60
        if match.group(3):
            seconds += int(match.group(3))

        segments.append({'ts': seconds, 'speaker': speaker})

    if not segments:
        raise ValueError("Не найдено валидных записей с таймкодами")

    # Сортировка и расчет длительности
    segments.sort(key=lambda x: x['ts'])
    stats = {}

    for i in range(len(segments) - 1):
        curr = segments[i]
        duration = min(segments[i + 1]['ts'] - curr['ts'], max_silence_sec)

        if curr['speaker'] not in stats:
            stats[curr['speaker']] = SpeakerStat(name=curr['speaker'])

        stats[curr['speaker']].total_seconds += duration
        stats[curr['speaker']].segment_count += 1

    # Подсчёт общего времени для расчёта процентов
    total_seconds = sum(s.total_seconds for s in stats.values())

    # Вывод результатов
    print("\n" + "=" * 80)
    print("ХРОНОМЕТРАЖ ВЫСТУПЛЕНИЙ")
    print("=" * 80)
    print(f"{'№':<3} {'СПИКЕР':<50} {'ВРЕМЯ':<15} {'% ОТ ОБЩЕГО':<10}")
    print("-" * 80)

    for i, stat in enumerate(sorted(stats.values(), key=lambda x: x.total_seconds, reverse=True), 1):
        pct = (stat.total_seconds / total_seconds * 100) if total_seconds else 0
        print(f"{i:<3} {stat.name:<50} {stat.formatted_time:<15} {pct:>6.1f}%")

    print("-" * 80)
    print(f"{'ИТОГО':<50} {total_seconds // 60} мин {total_seconds % 60} сек  (100.0%)")
    print("=" * 80 + "\n")

    # Экспорт в Excel если указан output_file
    if output_file:
        data = []
        for stat in sorted(stats.values(), key=lambda x: x.total_seconds, reverse=True):
            pct = (stat.total_seconds / total_seconds * 100) if total_seconds else 0
            data.append({
                'Спикер': stat.name,
                'Время_сек': stat.total_seconds,
                'Время_формат': stat.formatted_time,
                'Процент_от_общего': round(pct, 1)
            })
        pd.DataFrame(data).to_excel(output_file, index=False, engine='openpyxl')
        print(f"Результаты сохранены в: {output_file}")

    return stats


# ===== ОСНОВНОЙ КОД =====
if __name__ == "__main__":
    # Открываем диалог выбора файла
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Выберите Excel-файл с хронометражем",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    if not file_path:
        print("Файл не выбран!")
        exit()

    print(f"Выбран файл: {file_path}")

    # Запуск парсера
    stats = parse_chronometry(
        filepath=file_path,
        output_file="chronometry_result.xlsx",
        max_silence_sec=90,
        similarity_threshold=0.88
    )

    print("Готово!")