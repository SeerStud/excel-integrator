Информация о данном ПО в виде презентации: https://disk.yandex.ru/i/qUS_4sIVi7xBYw

Для установки программного обеспечения на Windows понадобится:
1. Установить Python 3.8
1.1. Перейти по ссылке https://www.python.org/downloads/release/python-380/
1.2. Скачать «Windows x86-64 executable installer» или «Windows x86 executable installer» в зависимости от разрядности операционной системы
1.3. Установить файл и при установке поставить галочку "Add Python to PATH"

2. Установить программное обеспечение Excel Integrator
2.1. Перейти по ссылке https://github.com/SeerStud/excel-integrator
2.2. Скачать все файлы с репозитория
2.2. Открыть PowerShell
2.3. Выбрать диск, на который были сохранены файлы, с помощью команды «cd» (например «cd D:»)
2.4. Выбрать путь до папки с программным обеспечением с помощью команды «cd "$env:"») (например «cd "$env:D:\ExcelIntegrator"»)
2.5. Установить виртуальное окружение и зависимости:
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
python -m spacy download ru_core_news_lg

3. Запустить программное обеспечение Excel Integrator
3.1. Открыть PowerShell
3.2. Выбрать путь до папки с программным обеспечением с помощью команды «cd "$env:"») (например  «cd "$env:D:\ExcelIntegrator"»)
3.3. Активировать виртуальное окружение командой «.\venv\Scripts\activate»
3.4. Запустить программное обеспечение командой «python main.py»
