# Prolongation Analysis

проект на pandas для расчета коэффициентов пролонгации по аккаунт-менеджерам

## Что делает проект
- объединяет данные из двух CSV
- агрегирует дубли по id
- рассчитывает коэффициенты пролонгации 1m и 2m
- формирует Excel-отчет

## Файлы
- `main.py` — основной скрипт
- `data/prolongations.csv` — обезличенные данные по проектам
- `data/financial_data.csv` — обезличенные финансовые данные

## Запуск
```bash
pip install -r requirements.txt
python prolongation_report.py
