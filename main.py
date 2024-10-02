import os

from Excel import read_excel_file, save_excel
from config import OUT_DIR, EXCEL_FULL_PATH, SHEET_FOR_SPLIT


def extract_elements_by_index(source_list, index_list):
    new_list = []
    for index in index_list:
        try:
            new_list.append(source_list[index])
        except IndexError:
            pass
    return new_list


def get_range_header(data) -> dict:
    result = {}
    key = None
    for index, value in enumerate(data):
        if value is not None:
            key = value
        if value is not None or key is not None:
            result.setdefault(key, []).append(index)
    return result


def split_excel_file(input_file_excel=EXCEL_FULL_PATH, out_dir=OUT_DIR):
    import_data = read_excel_file(input_file_excel)

    files = {}
    for sheet, data in import_data.items():
        if sheet not in SHEET_FOR_SPLIT:
            continue
        header = data[0]
        for name_range, renge_val in get_range_header(header).items():
            data_without_header = [data[1], *data[3:]]
            data_temp = []
            for i, row in enumerate(data_without_header):
                try:
                    renge_val.remove(0)
                except ValueError:
                    pass
                d = extract_elements_by_index(row, renge_val)
                __temp_data = set(d)
                if __temp_data != {None}:
                    data_temp.append([row[0], *d])
            files[f'{sheet}/{name_range}.xlsx'] = data_temp

    for f, data_save_file in files.items():
        print(f, end='')
        try:
            dist_path = os.path.join(out_dir, f)
            os.makedirs(os.path.dirname(dist_path), exist_ok=True)
            save_excel(dist_path, data=data_save_file)
            print('  : OK')
        except IOError:
            print(f'  : ERROR')


if __name__ == '__main__':
    split_excel_file(input_file_excel=EXCEL_FULL_PATH, out_dir=OUT_DIR)
