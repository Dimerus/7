import pandas as pd
from difflib import get_close_matches

def find_best_matches(exact_names, distorted_names):
    matches = {}
    for distorted in distorted_names:
        # Находим наиболее похожее название (только 1 лучший вариант)
        closest = get_close_matches(distorted, exact_names, n=1, cutoff=0)
        if closest:
            matches[distorted] = closest[0]
        else:
            matches[distorted] = None  # Если не найдено соответствие
    return matches

def main():
    # Чтение файла Excel
    try:
        df = pd.read_excel('input.xlsx')  # Или 'input.xls' для старого формата
    except FileNotFoundError:
        print("Ошибка: Файл 'input.xlsx' не найден.")
        return
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")
        return
    
    # Проверка наличия двух колонок
    if len(df.columns) < 2:
        print("Ошибка: В файле должно быть как минимум две колонки.")
        return
    
    # Предполагаем, что первая колонка - точные названия, вторая - искажённые
    exact_col = df.columns[0]
    distorted_col = df.columns[1]
    
    exact_names = df[exact_col].dropna().unique().tolist()
    distorted_names = df[distorted_col].dropna().unique().tolist()
    
    # Находим соответствия
    matches = find_best_matches(exact_names, distorted_names)
    
    # Создаём DataFrame с результатами
    result = pd.DataFrame({
        'Искажённое название': matches.keys(),
        'Соответствие': matches.values()
    })
    
    # Сохраняем результаты в новый файл
    result.to_excel('matches_output.xlsx', index=False)
    print("Результаты сохранены в файл 'matches_output.xlsx'")

if __name__ == "__main__":
    main()
