import time

import openpyxl

from duckduckgo_search import DDGS

SLEEP = 1
# С какой строки начать
FIRST_ROW = 8
# Колонка с текстом для поиска
SEARCH_COLUMN = 3
# Колонка с URL
URL_COLUMN = 11
# Сколько максимально нужно картинок
MAX_RESULTS = 2
# Сохранять после обработки N строк:
SAVE_AFTER_N_ROWS = 100


def search_for_image(keyword):
    with DDGS() as ddgs:
        results = ddgs.images(
            keyword,
            max_results=MAX_RESULTS,
            region="ru-ru"        
        )
        image_urls = [r['image'] for r in results]

    return(image_urls)

def main():
    workbook = openpyxl.load_workbook('raw.xlsx')
    
    result_filename = 'results.xlsx'
    
    sheet = workbook.active
    
    print(f'Всего строк: {sheet.max_row}')
    seconds = sheet.max_row * (SLEEP+0.5)
    print(f'Примерно займет времени: {seconds // 60} минут {seconds % 60} секунд')

        
    for row in range(FIRST_ROW, sheet.max_row + 1):
        if row % SAVE_AFTER_N_ROWS == 0:
            workbook.save(result_filename)
            print(f'Промежуточный Результат сохранен в {result_filename}')
        
        keyword = sheet.cell(row=row, column=SEARCH_COLUMN).value
        
        print(f'Строка {row} поиск: {keyword}')
        try:
            image_urls = search_for_image(keyword)
        except Exception as err:
            print("Произошла ошибка:")
            print(err)
            print("Сохраняю результат")
            result_filename = 'results_error.xlsx'
            break
        
        print(f'Найдено картинок: {image_urls}')
        
        # Сохранить урлы картинок
        for url_idx in range(len(image_urls)):
            sheet.cell(row=row, column=URL_COLUMN+url_idx).value = image_urls[url_idx]
            
        time.sleep(SLEEP)

    workbook.save(result_filename)
    print(f'Результат сохранен в {result_filename}')
    
if __name__ == '__main__':
    main()