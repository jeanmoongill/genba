from openpyxl import Workbook, load_workbook

class ProcessExcel:

    def main():

        CATEGORY = 1
        SHIHYO = 2
        KAMOKU = 3

        print("Process starting....")

        wb = load_workbook('./for_test.xlsx')
        ws = wb.active

        _temp_shihyo = None
        
        data_list = []
        shihyo_temp = None
        category_temp = None

        type_list = []
        for row in range(1, ws.max_row):
            _value = ws.cell(row = row, column = 1).value

            type_dict = {}
            if _value in ['連携', '単体']:
                type_dict['type'] = _value
                type_dict['row'] = row
                type_list.append(type_dict)

        for index, row in enumerate(range(KAMOKU, ws.max_row + 1)):

            _data_dict = {}
            _shigyo = ws.cell(row = row, column = 2).value
            _kamoku = ws.cell(row = row, column = 3).value
            _category = ws.cell(row = row - 1, column = 1).value

            if _category is not None:
                category_temp = _category

            sub_dict = {}
            data = []

            _type = None
            for row_type in type_list:
                if row > row_type.get('row'):
                    _type = row_type.get('type')

            if _temp_shihyo != _shigyo and _shigyo is not None and _kamoku is not None:
                _kamoku_for_store = ws.cell(row = row - 1, column = KAMOKU).value
                if _kamoku_for_store:
                    
                    for data_row in range(1, 6):
                        _data = ws.cell(row = row - 1, column = KAMOKU + data_row).value
                        data.append(_data)
                
                    if _kamoku_for_store == '科目':
                        shihyo_temp = _shigyo
                        sub_dict['years'] = data
                    else:
                        sub_dict['type'] = _type
                        sub_dict['category'] = category_temp
                        sub_dict['shihyo'] = shihyo_temp
                        sub_dict[_kamoku_for_store] = data
                        shihyo_temp = _shigyo
                    
                    data_list.append(sub_dict)
                    _temp_shihyo = _shigyo

            if _kamoku is None:

                _last_kamoku = ws.cell(row = row - 1, column = 3).value

                if _last_kamoku is not None:
                    last_komoku_list = []
                    _last_komoku_dict = {}
                    for _last_komoku in range(1, 6):
                        _last_value = ws.cell(row = row - 1, column = KAMOKU + _last_komoku).value
                        last_komoku_list.append(_last_value)

                    _last_komoku_dict['type'] = _type
                    _last_komoku_dict['category'] = category_temp
                    _last_komoku_dict['shihyo'] = shihyo_temp
                    _last_komoku_dict[_last_kamoku] = last_komoku_list   
                    shihyo_temp = _shigyo

                    data_list.append(_last_komoku_dict)

                # 規準を格納する
                _kizyun_list = []
                _kizyun_dict = {}
                for _kizyun_col in range(1, 6):
                    _kizyun_data = ws.cell(row = row, column = KAMOKU + _kizyun_col).value

                    if _kizyun_data is not None:
                        _kizyun_list.append(_kizyun_data)

                if len(_kizyun_list) > 0:
                    _kizyun_dict['kizyun'] = _kizyun_list
                    data_list.append(_kizyun_dict)

        print("data_list: {}".format(data_list))
    

    def get_kizyun(self, row):
        """
        基準をセットする。
        """
        pass

    if __name__ == "__main__":
        main()