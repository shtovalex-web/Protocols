# Linux-дистрибутив (PyInstaller, без исходников)

Готовая папка для пользователя Linux: один файл **`ProtocolOOT`** + каталог **`data/`** (шаблоны, FAQ, Excel-образцы).

## Вариант 1: скачать готовую сборку (GitHub Actions)

1. Откройте репозиторий на GitHub → **Actions** → workflow **build-linux-dist**.
2. Выберите последний успешный запуск (ветка **`linux`**, push или **Run workflow**).
3. Внизу страницы — артефакт **`ProtocolOHT_linux_dist`** (zip).
4. Распакуйте на Linux-ПК, например в `/opt/ProtocolOOT/`.

Структура после распаковки:

```text
ProtocolOOT              ← запускать этот файл
data/                    ← шаблоны, инструкции (не удалять)
Data_base.xlsx           ← образец (если был в сборке)
Programs_base.xlsx
ИНСТРУКЦИЯ_папки_сборки.txt
```

Запуск:

```bash
chmod +x ProtocolOOT
./ProtocolOOT
```

## Вариант 2: собрать самому на Linux

Нужна ветка **`linux`** (в ней уже есть `linux_port/app/`):

```bash
git clone -b linux https://github.com/shtovalex-web/Protocols.git
cd Protocols/linux_port
chmod +x build_linux.sh install_deps.sh
./install_deps.sh
python3 -m pip install -r requirements-build.txt
./build_linux.sh
```

Результат: `linux_port/ProtocolOHT_linux_dist/`.

Архив для переноса:

```bash
cd linux_port/ProtocolOHT_linux_dist
zip -r ../ProtocolOHT_linux_dist.zip .
```

## Требования на целевом ПК

| Компонент | Зачем |
|-----------|--------|
| **glibc** той же эпохи, что у сборки (обычно Ubuntu 22.04+ / современный дистрибутив) | совместимость бинарника |
| **python3-tk** не нужен | tkinter вшит в сборку |
| **LibreOffice** (`libreoffice-writer`) | PDF с оформлением DOCX |
| **fonts-dejavu-core**, **fonts-liberation** | кириллица в упрощённом PDF |

Установка на Debian/Ubuntu:

```bash
sudo apt install libreoffice-writer fonts-dejavu-core fonts-liberation
```

## Отличие от запуска из исходников

| | Исходники (`./run.sh`) | Бинарник (`ProtocolOOT`) |
|--|------------------------|---------------------------|
| Нужен git / Python на ПК | да | нет |
| Обновление | `git pull` + `./install_deps.sh` | новый zip из Actions или пересборка |
| Размер | меньше | больше (onefile) |
| Функции | те же | те же (оверлеи Linux уже внутри) |

## Автосборка

Workflow **`.github/workflows/build-linux-dist.yml`** запускается при каждом **push в ветку `linux`** и по кнопке **Run workflow**. После синхронизации `main` → `linux` (см. `docs/LINUX_BRANCH.md`) новый бинарник появится в Actions через несколько минут.
