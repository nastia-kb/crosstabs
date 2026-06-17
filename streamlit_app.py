import streamlit as st
import pandas as pd
import numpy as np
import io
import string
from scipy.stats import t

def get_col_letter(idx):
    """0 -> 'A', 1 -> 'B', ..., 25 -> 'Z', 26 -> 'AA', ..."""
    idx += 1
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = string.ascii_uppercase[rem] + letters
    return letters

THIN_BORDER = 1
THICK_BORDER = 5
HEADER_BG = '#e5d6fa'
BASE_BG = '#f2f2f2'
LABEL_COL_WIDTH = 350
DATA_COL_WIDTH = 120

def make_border_pair(col_slice_map, n_cols, col_idx):
    """Borders are thick between columns belonging to different slices, thin otherwise."""
    if col_idx is None:
        return THIN_BORDER, THIN_BORDER
    left = THICK_BORDER if col_idx > 0 and col_slice_map[col_idx] != col_slice_map[col_idx - 1] else THIN_BORDER
    right = THICK_BORDER if col_idx < n_cols - 1 and col_slice_map[col_idx] != col_slice_map[col_idx + 1] else THIN_BORDER
    return left, right

def get_cached_format(workbook, cache, num_format=None, bg_color=None, bold=False,
                       font_color=None, align=None, left=THIN_BORDER, right=THIN_BORDER):
    key = (num_format, bg_color, bold, font_color, align, left, right)
    if key not in cache:
        fmt_dict = {'top': THIN_BORDER, 'bottom': THIN_BORDER, 'left': left, 'right': right, 'border_color': '#000000'}
        if num_format:
            fmt_dict['num_format'] = num_format
        if bg_color:
            fmt_dict['bg_color'] = bg_color
        if bold:
            fmt_dict['bold'] = True
        if font_color:
            fmt_dict['font_color'] = font_color
        if align:
            fmt_dict['align'] = align
        cache[key] = workbook.add_format(fmt_dict)
    return cache[key]

def write_formatted_table(writer, sheet_name, table, col_slice_map, start_row, n_base_rows,
                           col_letters=None, colored_cells=None, sig_letters=None,
                           value_num_format='0.00%'):
    """Writes a percent/value table with column widths, slice-boundary borders, header and
    base-row coloring. If col_letters is given, also writes a letter sub-header row and applies
    significance highlighting/suffixes from colored_cells/sig_letters. Returns the next free row."""
    has_sig = col_letters is not None

    if has_sig:
        letters_row = pd.DataFrame(
            [[col_letters.get(col, "") for col in range(table.shape[1])]],
            columns=table.columns, index=[""]
        )
        export_table = pd.concat([letters_row, table])
        export_table.index.name = table.index.name
        export_table.to_excel(writer, sheet_name=sheet_name, merge_cells=True, startrow=start_row, startcol=0)
    else:
        table.to_excel(writer, sheet_name=sheet_name, merge_cells=True, startrow=start_row, startcol=0)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    n_cols = table.shape[1]
    format_cache = {}

    worksheet.set_column_pixels(0, 0, LABEL_COL_WIDTH)
    worksheet.set_column_pixels(1, n_cols, DATA_COL_WIDTH)

    label_left, label_right = make_border_pair(col_slice_map, n_cols, None)
    header_label_fmt = get_cached_format(workbook, format_cache, bg_color=HEADER_BG, bold=True,
                                          left=label_left, right=label_right)
    index_name = table.index.name if isinstance(table.index.name, str) else ""
    worksheet.write_string(start_row, 0, index_name, header_label_fmt)
    for col_idx in range(n_cols):
        left, right = make_border_pair(col_slice_map, n_cols, col_idx)
        header_fmt = get_cached_format(workbook, format_cache, bg_color=HEADER_BG, bold=True,
                                        align='center', left=left, right=right)
        worksheet.write_string(start_row, col_idx + 1, str(table.columns[col_idx]), header_fmt)

    row_offset = start_row + 1
    if has_sig:
        letter_label_fmt = get_cached_format(workbook, format_cache, bg_color=HEADER_BG,
                                              left=label_left, right=label_right)
        worksheet.write_blank(row_offset, 0, None, letter_label_fmt)
        for col_idx in range(n_cols):
            letter = col_letters.get(col_idx, "")
            left, right = make_border_pair(col_slice_map, n_cols, col_idx)
            letter_fmt = get_cached_format(workbook, format_cache, bg_color=HEADER_BG, bold=True,
                                            align='center', left=left, right=right)
            if letter:
                worksheet.write_string(row_offset, col_idx + 1, letter, letter_fmt)
            else:
                worksheet.write_blank(row_offset, col_idx + 1, None, letter_fmt)
        row_offset += 1

    for row_idx in range(table.shape[0]):
        excel_row = row_offset + row_idx
        is_percent_row = row_idx < table.shape[0] - n_base_rows
        is_base_row = not is_percent_row

        row_label_bg = BASE_BG if is_base_row else None
        row_label_fmt = get_cached_format(workbook, format_cache, bg_color=row_label_bg,
                                           left=label_left, right=label_right)
        worksheet.write_string(excel_row, 0, str(table.index[row_idx]), row_label_fmt)

        for col_idx in range(n_cols):
            excel_col = col_idx + 1
            value = table.iloc[row_idx, col_idx]
            left, right = make_border_pair(col_slice_map, n_cols, col_idx)
            col_base = table.iloc[table.shape[0] - n_base_rows, col_idx]
            is_empty = is_percent_row and (pd.isna(col_base) or col_base == 0)

            if is_percent_row:
                cell_color = colored_cells.get((row_idx, col_idx)) if colored_cells else None
                letters = sig_letters.get((row_idx, col_idx)) if sig_letters else None

                if cell_color or letters:
                    letters_suffix = ",".join(sorted(letters)) if letters else ""
                    num_format = value_num_format
                    if letters_suffix:
                        num_format += '" >' + letters_suffix + '"'
                    bg_color = '#C6EFCE' if cell_color == 'up' else ('#FFC7CE' if cell_color == 'down' else None)
                    fmt = get_cached_format(
                        workbook, format_cache, num_format=num_format, bg_color=bg_color,
                        bold=bool(letters_suffix), font_color='#ED7D31' if letters_suffix else None,
                        align='center', left=left, right=right)
                else:
                    fmt = get_cached_format(workbook, format_cache, num_format=value_num_format,
                                             align='center', left=left, right=right)
            else:
                fmt = get_cached_format(workbook, format_cache, num_format='0', bg_color=BASE_BG,
                                         align='center', left=left, right=right)

            if is_empty:
                worksheet.write_blank(excel_row, excel_col, None, fmt)
            else:
                worksheet.write_number(excel_row, excel_col, value if pd.notna(value) else 0, fmt)

    return row_offset + table.shape[0]

def compute_significance(table, col_slice_map, n_base_rows, z_crit, restrict_to_same_slice=True):
    """Vs-baseline (col 0) and pairwise z-tests for proportions across columns 1..n.
    Returns (col_letters, colored_cells, sig_letters)."""
    bases = table.iloc[-n_base_rows, :]
    shares = table.iloc[:-n_base_rows, :]
    n_cols = table.shape[1]
    colored_cells = {}
    sig_letters = {}
    col_letters = {col: get_col_letter(col - 1) for col in range(1, n_cols)}

    for col in range(1, n_cols):
        for row in range(table.shape[0] - n_base_rows):
            p1 = shares.iloc[row, 0]
            p2 = shares.iloc[row, col]
            n1 = bases.iloc[0]
            n2 = bases.iloc[col]
            if n2 < 30:
                continue
            p = (p1 * n1 + p2 * n2) / (n1 + n2)
            se = np.sqrt(p * (1 - p) * (1 / n1 + 1 / n2))
            if se == 0:
                continue
            z = (p2 - p1) / se
            if abs(z) >= z_crit:
                colored_cells[(row, col)] = 'up' if p2 > p1 else 'down'

    for row in range(table.shape[0] - n_base_rows):
        for col_a in range(1, n_cols):
            for col_b in range(col_a + 1, n_cols):
                if restrict_to_same_slice and col_slice_map[col_a] != col_slice_map[col_b]:
                    continue
                pa = shares.iloc[row, col_a]
                pb = shares.iloc[row, col_b]
                na = bases.iloc[col_a]
                nb = bases.iloc[col_b]
                if na < 30 or nb < 30:
                    continue
                p = (pa * na + pb * nb) / (na + nb)
                se = np.sqrt(p * (1 - p) * (1 / na + 1 / nb))
                if se == 0:
                    continue
                z = (pa - pb) / se
                if abs(z) >= z_crit:
                    if pa > pb:
                        sig_letters.setdefault((row, col_a), set()).add(col_letters[col_b])
                    else:
                        sig_letters.setdefault((row, col_b), set()).add(col_letters[col_a])

    return col_letters, colored_cells, sig_letters

def compute_mean_significance(table, col_slice_map, n_base_rows, z_crit):
    """Welch's t-test on the 'Среднее' row vs baseline (col 0) and pairwise within same slice.
    Returns (col_letters, colored_cells, sig_letters) in the same format as compute_significance."""
    n_cols = table.shape[1]
    col_letters = {col: get_col_letter(col - 1) for col in range(1, n_cols)}
    colored_cells = {}
    sig_letters = {}

    try:
        mean_row = list(table.index).index("Среднее")
    except ValueError:
        return col_letters, colored_cells, sig_letters

    alpha = 0.05 if z_crit == 1.96 else 0.10
    bases = table.loc["База"].astype(float)
    means = table.loc["Среднее"].astype(float)
    stds = table.loc["Стандартное отклонение"].astype(float)

    def _welch(n1, s1, n2, s2, m1, m2):
        if n1 < 2 or n2 < 2 or s1 == 0 or s2 == 0:
            return None
        se = np.sqrt(s1**2/n1 + s2**2/n2)
        if se == 0:
            return None
        dfree = (s1**2/n1 + s2**2/n2)**2 / ((s1**2/n1)**2/(n1-1) + (s2**2/n2)**2/(n2-1))
        return (m2 - m1) / se, t.ppf(1 - alpha/2, dfree)

    for col in range(1, n_cols):
        n1, n2 = bases.iloc[0], bases.iloc[col]
        if n2 < 30:
            continue
        result = _welch(n1, stds.iloc[0], n2, stds.iloc[col], means.iloc[0], means.iloc[col])
        if result and abs(result[0]) >= result[1]:
            colored_cells[(mean_row, col)] = 'up' if means.iloc[col] > means.iloc[0] else 'down'

    for col_a in range(1, n_cols):
        for col_b in range(col_a + 1, n_cols):
            if col_slice_map[col_a] != col_slice_map[col_b]:
                continue
            na, nb = bases.iloc[col_a], bases.iloc[col_b]
            if na < 30 or nb < 30:
                continue
            result = _welch(na, stds.iloc[col_a], nb, stds.iloc[col_b],
                            means.iloc[col_a], means.iloc[col_b])
            if result and abs(result[0]) >= result[1]:
                if means.iloc[col_a] > means.iloc[col_b]:
                    sig_letters.setdefault((mean_row, col_a), set()).add(col_letters[col_b])
                else:
                    sig_letters.setdefault((mean_row, col_b), set()).add(col_letters[col_a])

    return col_letters, colored_cells, sig_letters


def write_matrix_block(writer, matrix_table, matrix_is_numeric, matrix_row_n, matrix_row_n_sig, z_crit, n_base_rows):
    if isinstance(matrix_table, pd.DataFrame) and matrix_table.empty:
        return matrix_row_n, matrix_row_n_sig
    if isinstance(matrix_table, pd.Series):
        matrix_table = matrix_table.to_frame()
    if matrix_is_numeric:
        bases = matrix_table.loc["База"].astype(float)
        data_rows = matrix_table.iloc[:-n_base_rows].astype(float)
        avg_data = {}
        for row_label in data_rows.index:
            if row_label == "Сумма":
                avg_data[row_label] = data_rows.loc[row_label].sum()
            else:
                avg_data[row_label] = (data_rows.loc[row_label] * bases).sum() / bases.sum() if bases.sum() > 0 else np.nan
        avg_data[matrix_table.index[-n_base_rows]] = bases.sum()
        if n_base_rows == 2:
            avg_data[matrix_table.index[-1]] = matrix_table.iloc[-1].astype(float).sum()
        avg_col = pd.DataFrame({"Средневзвешенное": avg_data})
        full = pd.concat([avg_col, matrix_table], axis=1)
        full.index.name = matrix_table.index.name
        col_map = [None] * full.shape[1]
        matrix_row_n = write_formatted_table(writer, 'matrixes', full, col_map, matrix_row_n, n_base_rows, value_num_format='0.00')
        if z_crit > 0:
            m_cl, m_cc, m_sl = compute_mean_significance(full, col_map, n_base_rows, z_crit)
            matrix_row_n_sig = write_formatted_table(writer, 'matrixes_sig', full, col_map, matrix_row_n_sig, n_base_rows, col_letters=m_cl, colored_cells=m_cc, sig_letters=m_sl, value_num_format='0.00')
        else:
            matrix_row_n_sig = write_formatted_table(writer, 'matrixes_sig', full, col_map, matrix_row_n_sig, n_base_rows, value_num_format='0.00')
    else:
        weights = matrix_table.iloc[-n_base_rows].astype(float)
        shares = matrix_table.iloc[:-n_base_rows].astype(float)
        weighted_avg = shares.mul(weights, axis=1).sum(axis=1) / weights.sum()
        avg_col = pd.DataFrame({"Средневзвешенное": weighted_avg})
        avg_base = pd.DataFrame({"Средневзвешенное": [weights.sum()]}, index=[matrix_table.index[-n_base_rows]])
        avg_col = pd.concat([avg_col, avg_base])
        if n_base_rows == 2:
            avg_base_wt = pd.DataFrame({"Средневзвешенное": [matrix_table.iloc[-1].astype(float).sum()]}, index=[matrix_table.index[-1]])
            avg_col = pd.concat([avg_col, avg_base_wt])
        full = pd.concat([avg_col, matrix_table], axis=1)
        full.index.name = matrix_table.index.name
        col_map = [None] * full.shape[1]
        matrix_row_n = write_formatted_table(writer, 'matrixes', full, col_map, matrix_row_n, n_base_rows)
        if z_crit > 0:
            m_cl, m_cc, m_sl = compute_significance(full, col_map, n_base_rows, z_crit, restrict_to_same_slice=False)
            matrix_row_n_sig = write_formatted_table(writer, 'matrixes_sig', full, col_map, matrix_row_n_sig, n_base_rows, col_letters=m_cl, colored_cells=m_cc, sig_letters=m_sl)
        else:
            matrix_row_n_sig = write_formatted_table(writer, 'matrixes_sig', full, col_map, matrix_row_n_sig, n_base_rows)
    return matrix_row_n + 3, matrix_row_n_sig + 3


st.title("Вывод таблиц")

st.info("Здесь можно вывести кросс-таблицы со взвешиванием и без для опроса EnjoySurvey\n\n" \
"[Инструкция по работе](%s)" % "https://drive.google.com/file/d/1qnguxoMI2PBqxKcI_zZxYvncynFqylJp/view?usp=drive_linkl", 
        icon="💡")

with st.expander("Краткое описание"):
            st.write("Для вывода таблиц загрузите файл Excel EnjoySurvey (с лейблами, не кодами), проверьте корректность распознования типов вопросов и отметьте вопросы, по которым нужны разрезы. \n\n" \
            "В базу можно добавить собственные переменные, но их формат должен быть таким же, как у EnjoySurvey: уникальный код переменной, название вопроса, для множественного выбора - вариант ответа в названии вопроса (Вопрос - ответ) \n\n" \
            "Для взвешивания добавьте в базу **столбец wt**\n\n" \
            "Скрипт обрабатывает следующие типы вопросов:\n\n" \
            "- **Один ответ** - вопрос, состоящий из одного столбца в базе\n\n" \
            "- **Множестсвенный ответ** - вопрос, состоящий из нескольких столбцов в базе\n\n" \
            "- **Шкала** (5-балльная) - вопрос с 5 вариантами ответа, которые содержат характерные слова (скорее не, возможно, полностью и т.д.), поэтому при использовании нестандартных шкал данные могут обрабатываться некорректно. В таблицах строки сортируются по возрастанию оценок, выводятся Топ-2 и Боттом-2\n\n" \
            "- **Число** - вопрос, больше 75% ответов в котором являются числами. Исключаются выбросы (+/- 1.5 квартильных размаха), в таблицах показывается сумма, среднее, стандартное отклонение и база. При взвешивании числовые переменные также взвешиваются.\n\n" \
            "- **Матрица** - Определяется по наличию 'r_' в коде переменной. Обрабатывается так же, как и типы вопросов выше, но дополнительно выводится сводная таблица общих итогов на отдельном листе\n\n" \
            "- **Тестовый ответ** - вопрос с более чем 150 уникальными вариантами ответа, не выводится в таблицах. Если базе есть закрытые вопросы с таким количеством вариантов, измените тип вопроса на Один ответ вручную\n\n" \
            "Комментарии к другому (в базе отмечены 'o' в коде вопроса) исключаются из анализа. Если в них содержатся важные данные (например, чеки), перед загрузкой файла удалите 'o' из кодов переменных")


uploaded_file = st.file_uploader(
    "Загрузите файл Excel", type='xlsx', accept_multiple_files=False)

if 'stage' not in st.session_state:
    st.session_state.stage = 0

def set_state(i):
    st.session_state.stage = i

if st.session_state.stage == 0:
    st.button('Запустить', on_click=set_state, args=[1])

if st.session_state.stage == 1:
    data = pd.read_excel(uploaded_file, engine="openpyxl")

    param_names = data.iloc[0, :]

    unique_vars = []

    for i in param_names.index:
        if "q" in i:
            if "r" in i and "o" not in i:
                if i.count("_") == 1:
                    data.rename(columns={i: i+"_"}, inplace = True)
                    unique_vars.append(i+"_")
                else:
                    clean = i[:i.rfind("_")]+"_"
                    data.rename(columns={i: i+"_"}, inplace = True)
                    if clean not in unique_vars:
                        unique_vars.append(clean)
            elif "_" in i and "r" not in i:
                clean = i.split("_")[0]+"_"
                if clean not in unique_vars:
                    unique_vars.append(clean)
            elif "o" not in i:
                data.rename(columns={i: i+"_"}, inplace = True)
                unique_vars.append(i+"_")
        elif "Q" in i:
            data.rename(columns={i: i+"_"}, inplace = True)
            unique_vars.append(i+"_")
    
    data.drop(columns = ['№ записи', 'id', 'ac', 'starttime', 'endtime', 'surveytime', 'status'], inplace = True)
    param_names = data.iloc[0, :].to_frame()
    data.drop([0, 1], inplace = True)
    data.replace(" ", np.nan, inplace = True)

    var_df = pd.DataFrame()
    for var in unique_vars:
        temp = data.filter(like = var)
        if "_r" in var:
                if temp.shape[1] > 1:
                    var_type = "Матрица. Множественный ответ"
                    var_name = param_names.at[temp.columns[0], 0]
                    var_name = var_name[:var_name.rfind(" - ")]
                else:
                    temp_s = pd.Series(temp[var].dropna())
                    if (pd.to_numeric(temp_s, errors='coerce').count()) > (temp_s.shape[0] * 0.5):
                        var_type = "Матрица. Число"
                        var_name = param_names.at[f"{var}", 0]
                    elif temp_s.nunique() == 5:
                        answers = "".join(temp_s.astype(str).unique())
                        answers = answers.lower().replace(" ", "")
                        if "скореене" in answers:
                            var_type = "Матрица. Шкала"
                            var_name = param_names.at[f"{var}", 0]
                        else:
                            var_type = "Матрица. Один ответ"
                            var_name = param_names.at[f"{var}", 0]

                    else:
                        var_type = "Матрица. Один ответ"
                        var_name = param_names.at[f"{var}", 0]
                
        elif temp.shape[1] > 1:
             var_type = "Множественный ответ"
             var_name = param_names.at[temp.columns[0], 0]
             var_name = var_name[:var_name.rfind(" - ")]
        elif temp.shape[1] == 1:
            temp = temp.dropna().squeeze()
            if temp.nunique() <= 1:
                var_type = "Служебная переменная (не будет в таблицах)"
                var_name = param_names.at[f"{var}", 0]
            elif pd.to_numeric(temp, errors = "coerce").count() > temp.shape[0]*0.75:
                var_type = "Число"
                var_name = param_names.at[f"{var}", 0]
            elif temp.nunique() == 5:
                answers = "".join(temp.dropna().unique())
                answers = answers.lower()
                answers = answers.replace(" ", "")
                if "скореене" in answers or "скорееулуч" in answers:
                    var_type = "Шкала"
                    var_name = param_names.at[f"{var}", 0]
                else:
                    var_type = "Один ответ"
                    var_name = param_names.at[f"{var}", 0]
            elif temp.nunique() > 100: 
                var_type = "Открытый вопрос (не будет в таблицах)"
                var_name = param_names.at[f"{var}", 0]
            else:
                var_type = "Один ответ"
                var_name = param_names.at[f"{var}", 0]

        fin_var = pd.DataFrame({"Переменная": [var], "Вопрос": [var_name], "Тип вопроса": [var_type]})
        var_df = pd.concat([var_df, fin_var], axis = 0)

    var_df.insert(loc=3, column='Вывести разрез', value=False)
    
    st.session_state["var_df"] = var_df
    st.session_state["data"] = data
    st.session_state["param_names"] = param_names
    
    set_state(2)

if st.session_state.stage == 2:
    with st.form(key='my_form'):
        need_weight = st.checkbox("Взвесить данные (в базе должен быть **столбец wt**)")
        need_freqs = st.checkbox("Вывести частотные таблицы")

        st.session_state["need_weight"] = need_weight
        st.session_state["need_freqs"] = need_freqs

        option_map = {
            0: "Нет",
            1.645: "90%",
            1.96: "95%",
        }
        selection = st.segmented_control(
        "Сравнить разрезы с тоталом?",
        options=option_map.keys(),
        format_func=lambda option: option_map[option],
        selection_mode="single",
        default = 0,
        )
        
        st.write("Проверьте, верно ли определены типы вопросов, и отметьте переменные, по которым необходимо вывести разрезы")
        edited_df_in_form = st.data_editor(
            st.session_state["var_df"],
            column_config={
            "Переменная": None,
            "Вопрос": "Вопрос",
            "Тип вопроса": st.column_config.SelectboxColumn(
                "Тип вопроса",
                width="medium",
                options=[
                    "Один ответ",
                    "Множественный ответ",
                    "Шкала",
                    "Число",
                    "Матрица. Один ответ",
                    "Матрица. Множественный ответ",
                    "Матрица. Шкала",
                    "Матрица. Число",
                    "Открытый вопрос (не будет в таблицах)",
                    "Служебная переменная (не будет в таблицах)"
                ],
                required=True,
            ),
            "Вывести разрез":"Вывести разрез?"
            },
            hide_index=True,
            disabled = ["Вопрос"],
            )
        submit_button = st.form_submit_button("Все готово, вывести таблицы")

    if submit_button:
        st.session_state["var_df"] = edited_df_in_form
        st.session_state["z_score"] = selection
        set_state(3)

if st.session_state.stage == 3:
    var_df = st.session_state["var_df"]
    data = st.session_state["data"]
    param_names = st.session_state["param_names"].T
    slices = var_df.loc[var_df["Вывести разрез"], "Переменная"].to_list()
    unique_vars = var_df["Переменная"].to_list()
    need_weight = st.session_state["need_weight"]
    z_crit = st.session_state["z_score"]
    
    new_data = pd.DataFrame()
    data.replace(0, np.nan, inplace = True)
    data = data.astype("object")

    for var in unique_vars:
        temp_data = data.filter(like = var)

        if temp_data.shape[1] == 1:
            dummy_temp_data = pd.get_dummies(temp_data)
            dummy_temp_data.replace({True: 1, False: 0}, inplace = True)
            new_data = pd.concat([new_data, dummy_temp_data], axis = 1)
        else:
            temp_data.mask(temp_data.notna(), 1, inplace = True)
            new_names = []
            names = param_names.filter(like = var).iloc[0,:].tolist()
            for i in names:
                if len(i.split(" - "))>1:
                    name = i.split(" - ")[-1]
                else:
                    name = " "
                new_names.append(var+"_"+name)
            temp_data.columns = new_names
            new_data = pd.concat([new_data, temp_data], axis = 1)
        new_data.replace(0, np.nan, inplace = True)

    if need_weight:
        new_data["wt"] = data["wt"]
        new_data.iloc[:, :-1] = new_data.iloc[:, :-1].mul(data["wt"], axis=0)
    
    buffer = io.BytesIO()
    rows_n = 0
    rows_n_sig = 0

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:

        matrix_var_prev = 0
        matrix_table = pd.DataFrame()
        matrix_row_n = 0
        matrix_row_n_sig = 0
        matrix_is_numeric = False
        last_matrix = var_df.loc[var_df["Тип вопроса"].isin(["Матрица. Один ответ", "Матрица. Множественный ответ", "Матрица. Шкала", "Матрица. Число"]), "Переменная"].values
        if len(last_matrix) > 0:
            last_matrix = last_matrix[-1]

        for var in unique_vars:
            var_type = var_df.loc[var_df["Переменная"] == var, "Тип вопроса"].values[0]
            if var_type in ["Один ответ", "Множественный ответ", "Матрица. Один ответ", "Матрица. Множественный ответ", "Шкала", "Матрица. Шкала"]:
                temp_data = new_data.filter(like = var)
                temp_data.dropna(axis = 0, how = "all", inplace = True)
                count = pd.DataFrame({"Общий итог": temp_data.sum()})
                new_index = count.index.map(lambda x: x[x.find("__")+2:])
                count.index = new_index
                base = pd.DataFrame({"Общий итог": [temp_data.shape[0]]}, index = ["База"])
                if need_weight:
                    base_wt = pd.DataFrame({"Общий итог" : [new_data.loc[temp_data.index, "wt"].sum()]}, index = ["Взвешенная база"])
                    table = pd.concat([count, base, base_wt], axis = 0)
                else:
                    table = pd.concat([count, base])

                col_slice_map = [None]
                for slice in slices:
                    for group in new_data.filter(like = slice).columns:
                        group_index = new_data.loc[new_data[group] > 0].index
                        group_index_fin = [item for item in group_index if item in temp_data.index]
                        group_data = temp_data.loc[group_index_fin]
                        count = pd.DataFrame({group[group.find("__")+2:]: group_data.sum()})
                        count.index = new_index
                        base = pd.DataFrame({group[group.find("__")+2:]: [group_data.shape[0]]}, index = ["База"])
                        if need_weight:
                            base_wt = pd.DataFrame({group[group.find("__")+2:] : [new_data.loc[group_index_fin, "wt"].sum()]}, index = ["Взвешенная база"])
                            temp_table = pd.concat([count, base, base_wt], axis = 0)
                        else:
                            temp_table = pd.concat([count, base])
                        table = pd.concat([table, temp_table], axis = 1)
                        col_slice_map.append(slice)
                
                if var_type in ["Шкала", "Матрица. Шкала"]:
                    if need_weight:
                        bases = table.iloc[-2:,:]
                        counts = table.iloc[:-2,:]
                        test = counts.index
                    else:
                        bases = table.iloc[-1:,:]
                        counts = table.iloc[:-1,:]
                        test = counts.index
                    
                    new_index = []

                    for i in test:
                        text = i.lower().replace(" ", "")
                        if "отчасти" in text or "возможно" in text or "никак" in text or "неизмен" in text:
                            new_i = "3. "+i
                        elif "совсемне" in text or "совершенноне" in text or "точноне" in text or "определенноухудш" in text or "однозначноне" in text or "определенноне" in text:
                            new_i = "1. "+i
                        elif "скореене" in text or "скорееухудш" in text:
                            new_i = "2. "+i
                        elif "скорее" in text:
                            new_i = "4. "+i
                        else:
                            new_i = "5. "+i
                        new_index.append(new_i)
                    
                    counts.index = new_index
                    counts = counts.sort_index(ascending=True)
                    bottom2 = pd.DataFrame(counts.iloc[:2,:].sum(axis = 0)).T
                    bottom2.index = ["Боттом-2"]
                    top2 = pd.DataFrame(counts.iloc[3:,:].sum(axis = 0)).T
                    top2.index = ["Топ-2"]
                    counts = pd.concat([counts, bottom2, top2])
                    table = pd.concat([counts, bases])               

                table = table.astype(float)
                freq_table = table.copy()
                if need_weight:
                    table[:-2] = table[:-2].div(table.iloc[-1])
                else:
                    table[:-1] = table[:-1].div(table.iloc[-1])

                table.index.name = var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0]
                freq_table.index.name = var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0]

                if var_type in ["Матрица. Один ответ", "Матрица. Множественный ответ", "Матрица. Шкала"]:
                    matrix_var_curr = var.split("_")[0]

                    if matrix_var_curr != matrix_var_prev or var == last_matrix:
                        if var == last_matrix:
                            temp_matrix = table["Общий итог"]
                            temp_matrix.rename(var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0], inplace = True)
                            matrix_table = pd.concat([matrix_table, temp_matrix], axis = 1)

                        matrix_row_n, matrix_row_n_sig = write_matrix_block(
                            writer, matrix_table, matrix_is_numeric, matrix_row_n, matrix_row_n_sig,
                            z_crit, 2 if need_weight else 1)

                        matrix_table = table["Общий итог"]
                        matrix_table.rename(var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0], inplace = True)
                        matrix_is_numeric = False
                        matrix_var_prev = matrix_var_curr

                    else:
                        temp_matrix = table["Общий итог"]
                        temp_matrix.rename(var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0], inplace = True)
                        matrix_table = pd.concat([matrix_table, temp_matrix], axis = 1)
                
                
                n_base_rows = 2 if need_weight else 1

                write_formatted_table(writer, 'tables', table, col_slice_map, rows_n, n_base_rows)

                if z_crit > 0:
                    col_letters, colored_cells, sig_letters = compute_significance(
                        table, col_slice_map, n_base_rows, z_crit, restrict_to_same_slice=True)
                    write_formatted_table(writer, 'tables_sig', table, col_slice_map, rows_n_sig, n_base_rows,
                                          col_letters=col_letters, colored_cells=colored_cells,
                                          sig_letters=sig_letters)
                else:
                    write_formatted_table(writer, 'tables_sig', table, col_slice_map, rows_n_sig, n_base_rows)

                if need_freqs:
                    freq_table.to_excel(writer, sheet_name='frequencies', merge_cells = True, startrow=rows_n, startcol=0)
                    workbook = writer.book
                    worksheet = writer.sheets["frequencies"]

                rows_n = rows_n + table.shape[0] + 3
                rows_n_sig = rows_n_sig + table.shape[0] + (1 if z_crit > 0 else 0) + 3

            if var_type == "Матрица. Число":
                temp_data = pd.to_numeric(data[var].dropna(), errors="coerce")
                temp_check_list = pd.Series(temp_data.to_numpy().flatten())
                Q1 = temp_check_list.quantile(0.25)
                Q3 = temp_check_list.quantile(0.75)
                IQR = Q3 - Q1
                temp_data.loc[temp_data >= Q3 + 1.5*IQR] = np.nan
                temp_data.loc[temp_data <= Q1 - 1.5*IQR] = np.nan
                temp_data.dropna(inplace=True)

                if need_weight:
                    wts = data.loc[temp_data.index, "wt"]
                    sums = np.sum(temp_data * wts)
                    if wts.sum() > 0:
                        average = np.average(temp_data, weights=wts)
                        variance = np.average((temp_data-average)**2, weights=wts)
                        std = np.sqrt(variance)
                    else:
                        average = std = sums = np.nan
                    base = temp_data.count()
                    base_wt = wts.sum()
                    table = pd.DataFrame({"Общий итог": [sums, average, std, base, base_wt]},
                                          index=["Сумма", "Среднее", "Стандартное отклонение", "База", "Взвешенная база"])
                else:
                    sums = np.sum(temp_data)
                    average = np.average(temp_data)
                    variance = np.average((temp_data-average)**2)
                    std = np.sqrt(variance)
                    base = temp_data.count()
                    table = pd.DataFrame({"Общий итог": [sums, average, std, base]},
                                          index=["Сумма", "Среднее", "Стандартное отклонение", "База"])

                col_slice_map = [None]
                for slice in slices:
                    for group in new_data.filter(like=slice).columns:
                        group_index = new_data.loc[new_data[group] > 0].index
                        group_index_fin = [item for item in group_index if item in temp_data.index]
                        group_data = temp_data.loc[group_index_fin]
                        if need_weight:
                            wts = data.loc[group_index_fin, "wt"]
                            sums = np.sum(group_data * wts)
                            if wts.sum() > 0:
                                avg_g = np.average(group_data, weights=wts)
                                var_g = np.average((group_data-avg_g)**2, weights=wts)
                                std_g = np.sqrt(var_g)
                            else:
                                avg_g = std_g = sums = np.nan
                            base_g = group_data.count()
                            base_wt_g = wts.sum()
                            group_table = pd.DataFrame(
                                {group[group.find("__")+2:]: [sums, avg_g, std_g, base_g, base_wt_g]},
                                index=["Сумма", "Среднее", "Стандартное отклонение", "База", "Взвешенная база"])
                        else:
                            sums = np.sum(group_data)
                            avg_g = np.average(group_data)
                            var_g = np.average((group_data-avg_g)**2)
                            std_g = np.sqrt(var_g)
                            base_g = group_data.count()
                            group_table = pd.DataFrame(
                                {group[group.find("__")+2:]: [sums, avg_g, std_g, base_g]},
                                index=["Сумма", "Среднее", "Стандартное отклонение", "База"])
                        table = pd.concat([table, group_table], axis=1)
                        col_slice_map.append(slice)

                counts_d = {}
                table.columns = [f"{c}_{counts_d.setdefault(c,-1)+1}" if table.columns.tolist().count(c) > 1 else c
                                  for c in table.columns if not counts_d.update({c: counts_d.get(c,-1)+1})]
                table.index.name = var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0]

                num_n_base_rows = 2 if need_weight else 1
                write_formatted_table(writer, 'tables', table, col_slice_map, rows_n,
                                       num_n_base_rows, value_num_format='0.00')
                if z_crit > 0:
                    nc, ncc, nsl = compute_mean_significance(table, col_slice_map, num_n_base_rows, z_crit)
                    write_formatted_table(writer, 'tables_sig', table, col_slice_map, rows_n_sig,
                                           num_n_base_rows, col_letters=nc, colored_cells=ncc,
                                           sig_letters=nsl, value_num_format='0.00')
                else:
                    write_formatted_table(writer, 'tables_sig', table, col_slice_map, rows_n_sig,
                                           num_n_base_rows, value_num_format='0.00')

                rows_n = rows_n + table.shape[0] + 3
                rows_n_sig = rows_n_sig + table.shape[0] + (1 if z_crit > 0 else 0) + 3

                # Matrix accumulation: full stats row per sub-question
                if need_weight:
                    m_series = pd.Series(
                        [table.loc["Сумма", "Общий итог"], table.loc["Среднее", "Общий итог"],
                         table.loc["Стандартное отклонение", "Общий итог"],
                         table.loc["База", "Общий итог"], table.loc["Взвешенная база", "Общий итог"]],
                        index=["Сумма", "Среднее", "Стандартное отклонение", "База", "Взвешенная база"],
                        name=table.index.name)
                else:
                    m_series = pd.Series(
                        [table.loc["Сумма", "Общий итог"], table.loc["Среднее", "Общий итог"],
                         table.loc["Стандартное отклонение", "Общий итог"], table.loc["База", "Общий итог"]],
                        index=["Сумма", "Среднее", "Стандартное отклонение", "База"],
                        name=table.index.name)

                matrix_var_curr = var.split("_")[0]
                if matrix_var_curr != matrix_var_prev or var == last_matrix:
                    if var == last_matrix:
                        if not (isinstance(matrix_table, pd.DataFrame) and matrix_table.empty):
                            matrix_table = pd.concat([matrix_table, m_series], axis=1)
                        else:
                            matrix_table = m_series.to_frame()
                    matrix_row_n, matrix_row_n_sig = write_matrix_block(
                        writer, matrix_table, matrix_is_numeric, matrix_row_n, matrix_row_n_sig,
                        z_crit, num_n_base_rows)
                    matrix_table = m_series.to_frame()
                    matrix_is_numeric = True
                    matrix_var_prev = matrix_var_curr
                else:
                    if not (isinstance(matrix_table, pd.DataFrame) and matrix_table.empty):
                        matrix_table = pd.concat([matrix_table, m_series], axis=1)
                    else:
                        matrix_table = m_series.to_frame()

            if var_type == "Число":
                temp_data = pd.to_numeric(data[var].dropna(), errors = "coerce")
                temp_check_list = pd.Series(temp_data.to_numpy().flatten())
                Q1 = temp_check_list.quantile(0.25)
                Q3 = temp_check_list.quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                temp_data.loc[temp_data >= upper_bound] = np.nan
                temp_data.loc[temp_data <= lower_bound] = np.nan
                temp_data.dropna(inplace = True)

                if need_weight:
                    wts = data.loc[temp_data.index, "wt"]
                    sums = np.sum(temp_data * wts)
                    if wts.sum() > 0:
                        average = np.average(temp_data, weights=wts)
                        variance = np.average((temp_data-average)**2, weights=wts)
                        std = np.sqrt(variance)
                    else:
                        average = np.nan
                        variance = np.nan
                        std = np.nan

                    base = temp_data.count()
                    base_wt = wts.sum()
                    table = pd.DataFrame({"Общий итог" : [sums, average, std, base, base_wt]}, index = ["Сумма", "Среднее", "Стандартное отклонение", "База", "Взвешенная база"])

                else:
                    sums = np.sum(temp_data)
                    average = np.average(temp_data)
                    variance = np.average((temp_data-average)**2)
                    std = np.sqrt(variance)
                    base = temp_data.count()
                    table = pd.DataFrame({"Общий итог" : [sums, average, std, base]}, index = ["Сумма", "Среднее", "Стандартное отклонение", "База"])           
                
                col_slice_map = [None]
                for slice in slices:
                    for group in new_data.filter(like = slice).columns:
                        group_index = new_data.loc[new_data[group] > 0].index
                        group_index_fin = [item for item in group_index if item in temp_data.index]
                        group_data = temp_data.loc[group_index_fin]

                        if need_weight:
                            wts = data.loc[group_index_fin, "wt"]
                            sums = np.sum(group_data * wts)
                            if wts.sum()>0:
                                average = np.average(group_data, weights=wts)
                                variance = np.average((group_data-average)**2, weights=wts)
                                std = np.sqrt(variance)
                            else:
                                average = np.nan
                                variance = np.nan
                                std = np.nan
                            base = group_data.count()
                            base_wt = wts.sum()
                            group_table = pd.DataFrame({group[group.find("__")+2:]: [sums, average, std, base, base_wt]}, index = ["Сумма", "Среднее", "Стандартное отклонение", "База", "Взвешенная база"])

                        else:
                            sums = np.sum(group_data)
                            average = np.average(group_data)
                            variance = np.average((group_data-average)**2)
                            std = np.sqrt(variance)
                            base = group_data.count()
                            group_table = pd.DataFrame({group[group.find("__")+2:] : [sums, average, std, base]}, index = ["Сумма", "Среднее", "Стандартное отклонение", "База"])

                        table = pd.concat([table, group_table], axis = 1)
                        col_slice_map.append(slice)
                counts = {}

                table.columns = [f"{c}_{counts.setdefault(c, -1) + 1}" if table.columns.tolist().count(c) > 1 else c 
                    for c in table.columns if not counts.update({c: counts.get(c, -1) + 1})]
                
                table.index.name = var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0]

                num_n_base_rows = 2 if need_weight else 1

                write_formatted_table(writer, 'tables', table, col_slice_map, rows_n,
                                       num_n_base_rows, value_num_format='0.00')

                if z_crit > 0:
                    num_col_letters, num_colored_cells, num_sig_letters = compute_mean_significance(
                        table, col_slice_map, num_n_base_rows, z_crit)
                    write_formatted_table(writer, 'tables_sig', table, col_slice_map, rows_n_sig,
                                           num_n_base_rows, col_letters=num_col_letters,
                                           colored_cells=num_colored_cells, sig_letters=num_sig_letters,
                                           value_num_format='0.00')
                else:
                    write_formatted_table(writer, 'tables_sig', table, col_slice_map, rows_n_sig,
                                           num_n_base_rows, value_num_format='0.00')

                if need_freqs:
                    table.to_excel(writer, sheet_name='frequencies', merge_cells = True, startrow=rows_n, startcol=0)

                rows_n = rows_n + table.shape[0] + 3
                rows_n_sig = rows_n_sig + table.shape[0] + (1 if z_crit > 0 else 0) + 3

    st.download_button(
        label="Скачать результаты",
        data=buffer,
        file_name="crosstables.xlsx",
        mime="application/vnd.ms-excel",
        on_click=set_state, args=[0])

                                           

