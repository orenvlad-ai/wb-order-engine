# ⚙️ WB Order Engine

![GitHub release (latest by date)](https://img.shields.io/github/v/release/orenvlad-ai/wb-order-engine?label=версия&color=blue)
[![Download Excel](https://img.shields.io/badge/⬇️_Скачать_шаблон-Excel%20WB%20Engine-brightgreen?style=for-the-badge&logo=microsoft-excel)](https://github.com/orenvlad-ai/wb-order-engine/releases/latest/download/Planner_Latest.xlsx)

---

## 📘 Описание

**WB Order Engine** — сервис для расчёта оптимальных заказов товара с фабрики и формирования Excel-рекомендаций.  
API принимает исходные данные (остатки, план продаж, сроки поставок) и возвращает Excel-файл с листом `Recommendations`, уже оформленным для менеджеров.

---

## 🚀 Запуск локально

```bash
uvicorn app.main:app --reload
