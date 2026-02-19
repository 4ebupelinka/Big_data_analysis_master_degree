import marimo

__generated_with = "0.19.11"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo
    import os
    import tempfile
    import zipfile
    import rarfile
    import pandas as pd
    from pathlib import Path
    import matplotlib.pyplot as plt

    return os, pd, plt, rarfile, tempfile, zipfile


@app.cell
def _():
    YEAR_CONFIG = {
        2015: {
            'archive': 'Data/2015.rar', 'type': 'rar',
            'inner_pattern': 'СВОД_ВПО1_РОССИЯ.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 14, 'spec_row': 204, 'master_row': 311, 
            'col_budget_fed': ['F', 'G', 'H'], 'col_paid': 'I', 'name_row': 'A',
        },
        2016: {
            'archive': 'Data/2016.rar', 'type': 'rar',
            'inner_pattern': 'СВОД_ВПО1_РОССИЯ.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 12, 'spec_row': 191, 'master_row': 294, 
            'col_budget_fed': ['I'],'col_paid': 'O', 'name_row': 'A',
        },
        2017: {
            'archive': 'Data/2017.zip', 'type': 'zip',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 12, 'spec_row': 192, 'master_row': 295, 
            'col_budget_fed': ['J', 'L', 'M'],'col_paid': 'P', 'name_row': 'A',
        },
        2018: {
            'archive': 'Data/2018.rar', 'type': 'rar',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 12, 'spec_row': 194, 'master_row': 297, 
            'col_budget_fed': ['J', 'L', 'M'],'col_paid': 'P', 'name_row': 'A',
        },
        2019: {
            'archive': 'Data/2019.zip', 'type': 'zip',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 12, 'spec_row': 194, 'master_row': 297, 
            'col_budget_fed': ['J', 'L', 'M'],'col_paid': 'P', 'name_row': 'A',
        },
        2020: {
            'archive': 'Data/2020.zip', 'type': 'zip',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 12, 'spec_row': 194, 'master_row': 297, 
            'col_budget_fed': ['J', 'L', 'M'],'col_paid': 'P', 'name_row': 'A',
        },
        2021: {
            'archive': 'Data/2021.zip', 'type': 'zip',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 13, 'spec_row': 197, 'master_row': 303, 
            'col_budget_fed': ['J', 'L', 'M'],'col_paid': 'P', 'name_row': 'A',
        },
        2022: {
            'archive': 'Data/2022.zip', 'type': 'nested_zip',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 13, 'spec_row': 201, 'master_row': 307, 
            'col_budget_fed': ['J', 'L', 'M'],'col_paid': 'P', 'name_row': 'A',
        },
        2023: {
            'archive': 'Data/2023.zip', 'type': 'zip',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 13, 'spec_row': 201, 'master_row': 309, 
            'col_budget_fed': ['K', 'M', 'N'],'col_paid': 'R', 'name_row': 'A',
        },
        2024: {
            'archive': 'Data/2024.zip', 'type': 'zip',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xls', 'format': 'xls',
            'page': 'Р2_1_1', 'Bach_row': 13, 'spec_row': 201, 'master_row': 309, 
            'col_budget_fed': ['K', 'M', 'N'],'col_paid': 'R', 'name_row': 'A',
        },
        2025: {
            'archive': 'Data/2025.zip', 'type': 'zip',
            'inner_pattern': 'СВОД_ВПО1_ВСЕГО.xlsx', 'format': 'xlsx', 
            'page': 'Р2_1_1', 'Bach_row': 14, 'spec_row': 204, 'master_row': 311, 
            'col_budget_fed': ['K', 'M', 'N'],'col_paid': 'R', 'name_row': 'A',
        },
    }
    return (YEAR_CONFIG,)


@app.cell
def _(os, pd, rarfile, tempfile, zipfile):
    RARFILE_AVAILABLE = False
    def col_letter_to_index(col_letter: str) -> int:
        """Преобразует букву столбца Excel в индекс (A->0, B->1, ...)."""
        result = 0
        for char in col_letter.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1

    def extract_file_from_archive(archive_path: str, archive_type: str, inner_pattern: str, extract_to: str) -> str | None:
        """
        Извлекает файл inner_pattern из архива archive_path (тип archive_type) в папку extract_to.
        Возвращает полный путь к извлечённому файлу или None, если не удалось.
        """
        if archive_type == 'zip':
            with zipfile.ZipFile(archive_path, 'r') as zf:
                for member in zf.namelist():
                    if os.path.basename(member) == inner_pattern:
                        zf.extract(member, extract_to)
                        return os.path.join(extract_to, member)
        elif archive_type == 'rar':
            if not RARFILE_AVAILABLE:
                print(f"rarfile не доступен, пропускаем {archive_path}")
                return None
            with rarfile.RarFile(archive_path, 'r') as rf:
                for member in rf.namelist():
                    if os.path.basename(member) == inner_pattern:
                        rf.extract(member, extract_to)
                        return os.path.join(extract_to, member)
        elif archive_type == 'nested_zip':
            # Внешний архив - zip, внутри ещё один zip с нужным файлом
            with tempfile.TemporaryDirectory() as tmp_inner:
                with zipfile.ZipFile(archive_path, 'r') as outer_zip:
                    # Ищем внутренний zip, имя которого может быть любым, но предположим, что он один
                    inner_zips = [f for f in outer_zip.namelist() if f.endswith('.zip')]
                    if not inner_zips:
                        print(f"В архиве {archive_path} нет вложенных zip")
                        return None
                    # Извлекаем все внутренние zip во временную папку
                    for inner_zip_name in inner_zips:
                        outer_zip.extract(inner_zip_name, tmp_inner)
                        inner_zip_path = os.path.join(tmp_inner, inner_zip_name)
                        # Теперь ищем нужный файл внутри внутреннего zip
                        with zipfile.ZipFile(inner_zip_path, 'r') as inner_zip:
                            for member in inner_zip.namelist():
                                if os.path.basename(member) == inner_pattern:
                                    inner_zip.extract(member, extract_to)
                                    return os.path.join(extract_to, member)
        else:
            print(f"Неизвестный тип архива: {archive_type}")
        return None

    def read_excel_value(file_path: str, sheet_name: str, row: int, col_letter: str, file_format: str) -> float:
        """
        Читает числовое значение из Excel-файла.
        Возвращает 0, если ячейка пуста или не является числом.
        """
        try:
            # Выбираем движок в зависимости от формата
            if file_format.lower() == 'xls':
                engine = 'xlrd'  # для старых .xls
            else:
                engine = 'openpyxl'  # для .xlsx

            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine=engine)
            # Индекс строки в pandas (0-based) = row - 1
            idx_row = row - 1
            idx_col = col_letter_to_index(col_letter)
            val = df.iat[idx_row, idx_col]
            # Преобразуем в число, если возможно
            try:
                return float(val)
            except (ValueError, TypeError):
                return 0.0
        except Exception as e:
            print(f"Ошибка при чтении ячейки {col_letter}{row} из {file_path}: {e}")
            return 0.0

    return col_letter_to_index, extract_file_from_archive, read_excel_value


@app.cell
def _(extract_file_from_archive, os, pd, read_excel_value, tempfile):
    def extract_year_data_to_df(config: dict) -> pd.DataFrame:
        """
        Обрабатывает все года из конфига и возвращает DataFrame со столбцами:
        year, bach_budget, bach_paid, spec_budget, spec_paid
        """
        records = []  # список для накопления записей

        for year, params in config.items():
            print(f"\n--- Обработка года {year} ---")
            archive_path = params['archive']
            archive_type = params['type']
            inner_pattern = params['inner_pattern']
            file_format = params['format']
            sheet = params['page']
            bach_row = params['Bach_row']
            spec_row = params['spec_row']
            col_budget_fed = params['col_budget_fed']
            col_paid = params['col_paid']

            if not os.path.exists(archive_path):
                print(f"  Файл архива не найден: {archive_path}")
                continue

            with tempfile.TemporaryDirectory() as tmpdir:
                extracted_file = extract_file_from_archive(archive_path, archive_type, inner_pattern, tmpdir)
                if not extracted_file or not os.path.exists(extracted_file):
                    print(f"  Не удалось извлечь {inner_pattern} из {archive_path}")
                    continue

                print(f"  Извлечён файл: {extracted_file}")

                # Бакалавриат
                bach_budget = sum(
                    read_excel_value(extracted_file, sheet, bach_row, col, file_format)
                    for col in col_budget_fed
                )
                bach_paid = read_excel_value(extracted_file, sheet, bach_row, col_paid, file_format)

                # Специалитет
                spec_budget = sum(
                    read_excel_value(extracted_file, sheet, spec_row, col, file_format)
                    for col in col_budget_fed
                )
                spec_paid = read_excel_value(extracted_file, sheet, spec_row, col_paid, file_format)

                # Добавляем запись
                records.append({
                    'year': year,
                    'bach_budget': bach_budget,
                    'bach_paid': bach_paid,
                    'spec_budget': spec_budget,
                    'spec_paid': spec_paid
                })

                # Для отладки можно оставить печать
                print(f"  Бакалавриат: бюджет = {bach_budget}, платно = {bach_paid}")
                print(f"  Специалитет: бюджет = {spec_budget}, платно = {spec_paid}")

        # Создаём DataFrame
        df = pd.DataFrame(records)
        # Устанавливаем год как индекс (опционально)
        if not df.empty:
            df.set_index('year', inplace=True)
        return df

    return (extract_year_data_to_df,)


@app.cell
def _(YEAR_CONFIG, extract_year_data_to_df):
    df = extract_year_data_to_df(YEAR_CONFIG)
    return (df,)


@app.cell
def _(df):
    df
    return


@app.cell
def _(pd, plt):
    def plot_year_data(df: pd.DataFrame, kind='line', title='Динамика приёма по годам'):
        """
        Строит график показателей приёма (бакалавриат/специалитет, бюджет/платно) по годам.

        Параметры:
        df : DataFrame с индексом 'year' и колонками bach_budget, bach_paid, spec_budget, spec_paid
        kind : тип графика ('line' или 'bar')
        title : заголовок графика

        Возвращает:
        fig : объект matplotlib.figure.Figure для отображения в marimo
        """
        if df.empty:
            print("DataFrame пуст, нечего строить")
            return None

        # Сбрасываем индекс, чтобы год стал обычной колонкой
        plot_df = df.reset_index()
        x = plot_df['year']

        fig, ax = plt.subplots(figsize=(10, 6))

        if kind == 'line':
            ax.plot(x, plot_df['bach_budget'], marker='o', label='Бакалавриат бюджет')
            ax.plot(x, plot_df['bach_paid'], marker='s', label='Бакалавриат платно')
            ax.plot(x, plot_df['spec_budget'], marker='^', label='Специалитет бюджет')
            ax.plot(x, plot_df['spec_paid'], marker='d', label='Специалитет платно')
        elif kind == 'bar':
            width = 0.2
            x_pos = range(len(x))
            ax.bar([p - 1.5*width for p in x_pos], plot_df['bach_budget'], width, label='Бакалавриат бюджет')
            ax.bar([p - 0.5*width for p in x_pos], plot_df['bach_paid'], width, label='Бакалавриат платно')
            ax.bar([p + 0.5*width for p in x_pos], plot_df['spec_budget'], width, label='Специалитет бюджет')
            ax.bar([p + 1.5*width for p in x_pos], plot_df['spec_paid'], width, label='Специалитет платно')
            ax.set_xticks(x_pos)
            ax.set_xticklabels(x)
        else:
            raise ValueError("kind должен быть 'line' или 'bar'")

        ax.set_xlabel('Год')
        ax.set_ylabel('Количество мест')
        ax.set_title(title)
        ax.legend()
        ax.grid(True, linestyle='--', alpha=0.7)

        plt.tight_layout()
        return fig

    return (plot_year_data,)


@app.cell
def _(df, plot_year_data):
    plot_year_data(df)
    return


@app.cell
def _(col_letter_to_index, extract_file_from_archive, os, pd, tempfile):
    def extract_directions_data(config):
        """
        Возвращает два DataFrame (бакалавриат, специалитет) в широком формате:
        индекс — (год, тип), колонки — направления, значения — количество мест.
        """
        records_bach = []
        records_spec = []

        for year, params in config.items():
            print(f"\n--- Обработка года {year} для направлений ---")
            archive_path = params['archive']
            archive_type = params['type']
            inner_pattern = params['inner_pattern']
            file_format = params['format']
            sheet = params['page']
            bach_row = params['Bach_row']
            spec_row = params['spec_row']
            master_row = params['master_row']
            name_col_letter = params['name_row']
            budget_cols_letters = params['col_budget_fed']
            paid_col_letter = params['col_paid']

            if not os.path.exists(archive_path):
                print(f"  Файл архива не найден: {archive_path}")
                continue

            with tempfile.TemporaryDirectory() as tmpdir:
                extracted_file = extract_file_from_archive(archive_path, archive_type, inner_pattern, tmpdir)
                if not extracted_file or not os.path.exists(extracted_file):
                    print(f"  Не удалось извлечь {inner_pattern} из {archive_path}")
                    continue

                print(f"  Извлечён файл: {extracted_file}")

                # Читаем лист Excel
                try:
                    if file_format.lower() == 'xls':
                        engine = 'xlrd'
                    else:
                        engine = 'openpyxl'
                    df = pd.read_excel(extracted_file, sheet_name=sheet, header=None, engine=engine)
                except Exception as e:
                    print(f"  Ошибка чтения Excel: {e}")
                    continue

                # Индексы столбцов
                name_col_idx = col_letter_to_index(name_col_letter)
                budget_cols_idx = [col_letter_to_index(c) for c in budget_cols_letters]
                paid_col_idx = col_letter_to_index(paid_col_letter)

                # ----- Бакалавриат: строки от bach_row+1 до spec_row-1 -----
                bach_start = bach_row + 1
                bach_end = spec_row - 1
                if bach_start <= bach_end:
                    for r in range(bach_start, bach_end + 1):
                        idx = r - 1
                        if idx >= len(df):
                            continue
                        name_val = df.iat[idx, name_col_idx]
                        if pd.isna(name_val) or not isinstance(name_val, str):
                            continue
                        name = str(name_val).strip()
                        if name == '':
                            continue

                        # Сумма по бюджетным колонкам
                        budget_sum = 0.0
                        for col_idx in budget_cols_idx:
                            val = df.iat[idx, col_idx]
                            try:
                                budget_sum += float(val)
                            except (ValueError, TypeError):
                                pass

                        # Платное
                        paid_val = df.iat[idx, paid_col_idx]
                        try:
                            paid_val = float(paid_val)
                        except (ValueError, TypeError):
                            paid_val = 0.0

                        records_bach.append({'year': year, 'type': 'budget', 'direction': name, 'value': budget_sum})
                        records_bach.append({'year': year, 'type': 'paid', 'direction': name, 'value': paid_val})
                else:
                    print(f"  Для года {year} диапазон бакалавриата пуст")

                # ----- Специалитет: строки от spec_row+1 до master_row-1 -----
                spec_start = spec_row + 1
                spec_end = master_row - 1
                if spec_start <= spec_end:
                    for r in range(spec_start, spec_end + 1):
                        idx = r - 1
                        if idx >= len(df):
                            continue
                        name_val = df.iat[idx, name_col_idx]
                        if pd.isna(name_val) or not isinstance(name_val, str):
                            continue
                        name = str(name_val).strip()
                        if name == '':
                            continue

                        budget_sum = 0.0
                        for col_idx in budget_cols_idx:
                            val = df.iat[idx, col_idx]
                            try:
                                budget_sum += float(val)
                            except (ValueError, TypeError):
                                pass

                        paid_val = df.iat[idx, paid_col_idx]
                        try:
                            paid_val = float(paid_val)
                        except (ValueError, TypeError):
                            paid_val = 0.0

                        records_spec.append({'year': year, 'type': 'budget', 'direction': name, 'value': budget_sum})
                        records_spec.append({'year': year, 'type': 'paid', 'direction': name, 'value': paid_val})
                else:
                    print(f"  Для года {year} диапазон специалитета пуст")

        # Создаём широкие таблицы
        if records_bach:
            df_bach_long = pd.DataFrame(records_bach)
            df_bach_wide = df_bach_long.pivot_table(
                index=['year', 'type'],
                columns='direction',
                values='value',
                aggfunc='sum',
                fill_value=0
            )
        else:
            df_bach_wide = pd.DataFrame()

        if records_spec:
            df_spec_long = pd.DataFrame(records_spec)
            df_spec_wide = df_spec_long.pivot_table(
                index=['year', 'type'],
                columns='direction',
                values='value',
                aggfunc='sum',
                fill_value=0
            )
        else:
            df_spec_wide = pd.DataFrame()

        return df_bach_wide, df_spec_wide

    return (extract_directions_data,)


@app.cell
def _(YEAR_CONFIG, extract_directions_data):
    bach_df, spec_df = extract_directions_data(YEAR_CONFIG)
    return bach_df, spec_df


@app.cell
def _(bach_df):
    bach_df
    return


@app.cell
def _(spec_df):
    spec_df 
    return


@app.cell
def _(pd, plt):
    def plot_top5_budget_drop(bach_wide_df, spec_wide_df):
        """
        Строит два графика (бакалавриат и специалитет) для 5 направлений с наибольшим падением бюджетного приёма.
    
        Параметры:
        bach_wide_df, spec_wide_df — датафреймы в широком формате,
            индекс: (год, тип), колонки: направления.
    
        Возвращает:
        fig_bach, fig_spec — объекты matplotlib.figure.Figure для отображения в marimo.
        """
        def process_drop(df_wide, title_prefix):
            if df_wide.empty:
                print(f"Нет данных для {title_prefix}")
                return None
        
            # Отбираем только бюджетные строки
            budget_mask = df_wide.index.get_level_values('type') == 'budget'
            df_budget = df_wide[budget_mask].copy()
        
            # Переформатируем: строки = направления, колонки = годы
            df_budget = df_budget.reset_index(level='type', drop=True)
            # Убедимся, что годы — целые числа
            df_budget.index = df_budget.index.astype(int)
            # Транспонируем: направления как строки, годы как колонки
            df_plot = df_budget.T  # теперь строки — направления, колонки — годы
        
            # Заменяем 0 на NaN (направление отсутствовало)
            df_plot = df_plot.replace(0, pd.NA)
        
            # Словарь для хранения изменений
            drop_data = []
        
            for direction in df_plot.index:
                series = df_plot.loc[direction].dropna()  # убираем NaN (годы без направления)
                if len(series) < 2:
                    continue  # нужно минимум два года с ненулевым приёмом
                # Первый и последний год с ненулевыми значениями
                first_year = series.index[0]
                last_year = series.index[-1]
                first_val = series.iloc[0]
                last_val = series.iloc[-1]
                change = last_val - first_val
                drop_data.append({
                    'direction': direction,
                    'change': change,
                    'first_val': first_val,
                    'last_val': last_val,
                    'first_year': first_year,
                    'last_year': last_year,
                    'series': series  # сохраняем для графика
                })
        
            if not drop_data:
                print(f"Нет направлений с достаточными данными для {title_prefix}")
                return None
        
            # Сортируем по изменению (по возрастанию — самые отрицательные вверху)
            drop_df = pd.DataFrame(drop_data).sort_values('change')
            # Берём 5 с наибольшим падением (самые отрицательные)
            top5 = drop_df.head(5)
        
            # Строим график
            fig, ax = plt.subplots(figsize=(12, 6))
            for _, row in top5.iterrows():
                series = row['series']
                ax.plot(series.index, series.values, marker='o', label=row['direction'])
        
            ax.set_xlabel('Год')
            ax.set_ylabel('Количество бюджетных мест')
            ax.set_title(f'{title_prefix}: 5 направлений с наибольшим падением приёма')
            ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
            ax.grid(True, linestyle='--', alpha=0.7)
            plt.tight_layout()
            return fig
    
        fig_bach = process_drop(bach_wide_df, 'Бакалавриат')
        fig_spec = process_drop(spec_wide_df, 'Специалитет')
    
        return fig_bach, fig_spec

    return (plot_top5_budget_drop,)


@app.cell
def _(bach_df, plot_top5_budget_drop, spec_df):
    fig1, fig2 = plot_top5_budget_drop(bach_df, spec_df)

    fig1
    return (fig2,)


@app.cell
def _(fig2):
    fig2
    return


@app.cell
def _():
    return


if __name__ == "__main__":
    app.run()
