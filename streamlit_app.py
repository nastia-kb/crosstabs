import streamlit as st
import pandas as pd
import numpy as np
import io
from scipy.stats import t

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
                elif temp[var].nunique() == 5:
                    answers = "".join(temp[var].dropna().unique())
                    answers = answers.lower()
                    answers = answers.replace(" ", "")
                    if "скореене" in answers:
                        var_type = "Матрица. Шкала"
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
            if temp.nunique() == 1:
                unique_vars.remove(var)
                continue
            if pd.to_numeric(temp, errors = "coerce").count() > temp.shape[0]*0.75:
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
                    "Открытый вопрос (не будет в таблицах)"
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

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:

        matrix_var_prev = 0
        matrix_table = pd.DataFrame()
        matrix_row_n = 0
        last_matrix = var_df.loc[var_df["Тип вопроса"].isin(["Матрица. Один ответ", "Матрица. Множественный ответ", "Матрица. Шкала"]), "Переменная"].values
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

                        if matrix_table.shape[0] > 0:

                            matrix_table.to_excel(writer, sheet_name='matrixes', merge_cells = True, startrow=matrix_row_n, startcol=0)
                            workbook = writer.book
                            worksheet = writer.sheets["matrixes"]
                            
                            percent_format = workbook.add_format({'num_format': '0.00%'})

                            if need_weight:        
                                rows_to_format = [r for r in range(matrix_row_n, (matrix_row_n+matrix_table.shape[0])-1)]
                            else:
                                rows_to_format = [r for r in range(matrix_row_n, (matrix_row_n+matrix_table.shape[0]))]

                            for row in rows_to_format:
                                worksheet.set_row(row, cell_format = percent_format)

                            matrix_row_n = matrix_row_n + matrix_table.shape[0]+3

                        matrix_table = table["Общий итог"]
                        matrix_table.rename(var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0], inplace = True)
                        matrix_var_prev = matrix_var_curr

                    else:
                        temp_matrix = table["Общий итог"]
                        temp_matrix.rename(var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0], inplace = True)
                        matrix_table = pd.concat([matrix_table, temp_matrix], axis = 1)
                
                
                if z_crit > 0:

                    if need_weight:
                        bases = table.iloc[-2, :]
                        shares = table.iloc[:-2,:]
                    else:
                        bases = table.iloc[-1, :]
                        shares = table.iloc[:-1,:]
                    colored_cells = {}
                    for col in range(1, table.shape[1]):
                        for row in range(table.shape[0]):
                            if need_weight:
                                is_percent_row = row < table.shape[0] - 2
                            else:
                                is_percent_row = row < table.shape[0] - 1
                            
                            if not is_percent_row:
                                continue
                            p1 = shares.iloc[row, 0]
                            p2 = shares.iloc[row, col]
                            n1 = bases.iloc[0]
                            n2 = bases.iloc[col]
                            if n2 < 30:
                                continue
                            p = (p1*n1 + p2*n2) / (n1+n2)
                            se = np.sqrt(p*(1-p)*(1/n1 + 1/n2))
                            if se == 0:
                                continue
                            z = (p2 - p1) / se
                            if abs(z) >= z_crit:
                                if p2 > p1:
                                    colored_cells[(row, col)] = 'up'
                                else:
                                    colored_cells[(row, col)] = 'down'
                    
                    table.to_excel(writer, sheet_name='tables', merge_cells = True, startrow=rows_n, startcol=0)
                    workbook = writer.book
                    worksheet = writer.sheets["tables"]

                    percent_format = workbook.add_format({'num_format': '0.00%'})
    
                    green_percent = workbook.add_format({
                        'bg_color': '#C6EFCE',
                        'num_format': '0.00%'
                    })
                    
                    red_percent = workbook.add_format({
                        'bg_color': '#FFC7CE',
                        'num_format': '0.00%'
                    })
                    
                    number_format = workbook.add_format({'num_format': '0'})
                    
                    for row_idx in range(table.shape[0]):
                        excel_row = rows_n + row_idx + 1
                        
                        if need_weight:
                            is_percent_row = row_idx < table.shape[0] - 2
                        else:
                            is_percent_row = row_idx < table.shape[0] - 1
                        
                        for col_idx in range(table.shape[1]):
                            excel_col = col_idx + 1 
                            value = table.iloc[row_idx, col_idx]
                            
                            if is_percent_row:
                                cell_color = colored_cells.get((row_idx, col_idx))
                                
                                if cell_color == 'up':
                                    format_to_use = green_percent
                                elif cell_color == 'down':
                                    format_to_use = red_percent
                                else:
                                    format_to_use = percent_format
                            else:
                                format_to_use = number_format
                            
                            try:
                                worksheet.write_number(excel_row, excel_col, value, format_to_use)
                            except:
                                pass

                else:
                    table.to_excel(writer, sheet_name='tables', merge_cells = True, startrow=rows_n, startcol=0)
                    workbook = writer.book
                    worksheet = writer.sheets["tables"]
                    percent_format = workbook.add_format({'num_format': '0.00%'})
                    if need_weight:        
                        rows_to_format = [r for r in range(rows_n, (rows_n+table.shape[0])-1)]
                    else:
                        rows_to_format = [r for r in range(rows_n, (rows_n+table.shape[0]))]
                    for row in rows_to_format:
                        worksheet.set_row(row, cell_format = percent_format)
                
                if need_freqs:
                    freq_table.to_excel(writer, sheet_name='frequencies', merge_cells = True, startrow=rows_n, startcol=0)
                    workbook = writer.book
                    worksheet = writer.sheets["frequencies"]

                rows_n = rows_n + table.shape[0]+3

            def highlight_mean_significance(df, alpha=z_crit):

                bases = df.loc["База"].astype(float)
                means = df.loc["Среднее"].astype(float)
                stds = df.loc["Стандартное отклонение"].astype(float)

                if alpha == 1.96:
                    alpha = 0.05
                else:
                    alpha = 0.10

                colors = pd.DataFrame("", index=df.index, columns=df.columns)
                m1 = means["Общий итог"]
                s1 = stds["Общий итог"]
                n1 = bases["Общий итог"]

                for col in range(1, table.shape[0]):
                    m2 = means.iloc[col]
                    s2 = stds.iloc[col]
                    n2 = bases.iloc[col]
                    if n2 < 30:
                        continue
                    se = np.sqrt((s1**2)/n1 + (s2**2)/n2)
                    if se == 0:
                        continue
                    t_stat = (m2 - m1) / se
                    dfree = ((s1**2/n1 + s2**2/n2)**2) / (
                        ((s1**2/n1)**2)/(n1-1) + ((s2**2/n2)**2)/(n2-1)
                        )
                    t_crit = t.ppf(1 - alpha/2, dfree)

                    if abs(t_stat) >= t_crit:
                        if m2 > m1:
                            colors.loc["Среднее", df.columns[col]] = "background-color: #C6EFCE"
                        else:
                            colors.loc["Среднее", df.columns[col]] = "background-color: #FFC7CE"
                styler = df.style

                styler = styler.apply(
                    lambda _: colors, axis=None)

                return styler
            
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
                counts = {}

                table.columns = [f"{c}_{counts.setdefault(c, -1) + 1}" if table.columns.tolist().count(c) > 1 else c 
                    for c in table.columns if not counts.update({c: counts.get(c, -1) + 1})]
                
                table.index.name = var_df.loc[var_df["Переменная"] == var, "Вопрос"].values[0]
                
                if z_crit > 0:
                    fin_table = highlight_mean_significance(table, alpha = z_crit)
                    fin_table.to_excel(writer, sheet_name='tables', merge_cells = True, startrow=rows_n, startcol=0)
                else:
                    table.to_excel(writer, sheet_name='tables', merge_cells = True, startrow=rows_n, startcol=0)

                if need_freqs:
                    table.to_excel(writer, sheet_name='frequencies', merge_cells = True, startrow=rows_n, startcol=0)
                    workbook = writer.book
                    worksheet = writer.sheets["frequencies"]

                workbook = writer.book
                worksheet = writer.sheets["tables"]
                rows_n = rows_n + table.shape[0]+3           

    st.download_button(
        label="Скачать результаты",
        data=buffer,
        file_name="crosstables.xlsx",
        mime="application/vnd.ms-excel",
        on_click=set_state, args=[0])

                                           

