from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
import pprint
import json
from datetime import date

class ProcessExcel:

    def main():

        CATEGORY = 1
        SHIHYO = 2
        KAMOKU = 3

        wb = load_workbook('./for_test.xlsx')
        ws = wb["Sheet1"]

        _temp_shihyo = None
        
        data_list = []
        shihyo_temp = None
        category_temp = None

        type_list = []
        _years_list = []
        _kizyun_list = []
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
                        _years_list = data
                        # sub_dict['years'] = data
                        continue
                    else:

                        sub_dict['type'] = _type
                        sub_dict['category'] = category_temp
                        sub_dict['shihyo'] = shihyo_temp
                        sub_dict['kamoku'] = _kamoku_for_store
                        sub_dict['data'] = data
                        sub_dict['years'] = _years_list

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
                    _last_komoku_dict['kamoku'] = _last_kamoku
                    _last_komoku_dict['data'] = last_komoku_list   
                    _last_komoku_dict['years'] = _years_list
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

        # draw chart
        _rows = [
            ['Date', '指標', '規準', 'value']
        ]

        _years_list = []
        _kizyun_list = []
        _data_data_list = []

        for _dict_data in data_list:
            _keys = list(_dict_data.keys()) 
           
            if 'data' in _keys and 'years' in _keys:
                for _index in range(0, len(_dict_data.get('years'))):
                    _tmp = []
                    _tmp.append(_dict_data.get('years')[_index])
                    _tmp.append(_dict_data.get('data')[_index])
                    _tmp.append(_dict_data.get('type'))
                    _tmp.append(_dict_data.get('category'))
                    _tmp.append(_dict_data.get('shihyo'))
                    _tmp.append(_dict_data.get('kamoku'))
            
                    if len(_tmp) > 0:
                        _data_data_list.append(_tmp)

        print("list {}".format(_data_data_list))

        rows = [
            ['Date', 'Batch 1'],
            [date(2015,9, 1), 40],
            [date(2015,9, 2), 40],
            [date(2015,9, 3)],
            [date(2015,9, 4), 30],
            [date(2015,9, 5), 25],
            [date(2015,9, 6), 20],
        ]

        for row in rows:
            ws.append(row)

        c1 = LineChart()
        c1.title = "test"
        c1.style = 13
        c1.y_axis.title = 'Y axis'
        c1.x_axis.title = 'X axis'

        data = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=7)
        c1.add_data(data, titles_from_data=True)

        ws.add_chart(c1, "A10")
        wb.save("line.xlsx")

    def get_kizyun(self, row):
        """
        基準をセットする。
        """
        pass

    if __name__ == "__main__":
        main()