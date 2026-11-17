# Alphacam Batch Processor

Графическое приложение для пакетной подготовки и обработки деталей в Alphacam.

## Возможности

- последовательный импорт файлов Parasolid `X_T`;
- ручной поворот и переворот модели перед сохранением ARD;
- настройка каждой детали чекбоксами:
  - паз;
  - обратная сторона;
  - дополнительная программа `*-ROT`;
  - пропуск обработки;
- сохранение плана пакетной обработки;
- пересчёт `Width`, `Height`, `Thickness` перед выбором геометрии паза;
- создание стандартных, `*-OBR` и `*-ROT` программ;
- извлечение WorkPlane через штатные VBA-макросы Alphacam, если COM не публикует enum ориентации;
- экспорт NC по настройке пользователя;
- копирование лога из окна приложения.

## Требования

- Windows;
- установленный Alphacam с доступным COM API;
- Python 3.10 или новее;
- загруженный в Alphacam VBA-проект `DWFI_PR` с модулями:
  - `Paz`;
  - `Automation`;
  - `ExtractFeature`;
  - `GrooveStorage`.

По умолчанию используется шаблон:

```text
C:\ALPHACAM1\LICOMDIR\Templates\DEFAULT-PAZ.adt
```

и AutoStyle:

```text
C:\ALPHACAM1\LICOMDIR\Styles\DRILL\AUTODRILL.ara
```

## Установка

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Запуск

```powershell
.\.venv\Scripts\python.exe .\app\alphacam_gui.pyw
```

Либо запустите `app\alphacam_gui.pyw` двойным щелчком после установки зависимостей.

## Рабочий процесс

1. Нажмите `X_T → ARD + обработка`.
2. Выберите папку с X_T.
3. Для каждой детали настройте положение и чекбоксы.
4. После завершения импорта подтвердите запуск пакетной обработки.
5. Если выбран паз, параметры размеров сначала пересчитываются, затем Alphacam предлагает выбрать грань паза.

План сохраняется рядом с ARD:

```text
PR\_batch_plan.txt
```

Формат строки:

```text
file|groove|reverse|rotate_program|skip
```

## Настройки

Если VBA-проект называется не `DWFI_PR`, задайте переменную окружения:

```powershell
$env:ALPHACAM_VBA_PROJECT = "ИмяПроекта"
```

Для нестандартного ProgID Alphacam:

```powershell
$env:ALPHACAM_PROGID = "Broutaps.Application"
```

