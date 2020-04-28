import openpyxl
import json


class ConvertToSomething():
    def __init__(self, sheet, output):
        self._sheet = sheet
        self._output = output
        self._columns = []

    def _write_header(self):
        with open(self._output, mode='w') as f:
            f.write("[")

    def _write_footer(self):
        with open(self._output, mode='a') as f:
            f.write("]")

    def _set_columns(self):
        column = 1
        while self._check(1, column):
            self._columns.append(self._sheet.cell(row=1, column=column).value)
            column += 1

    def _check(self, row, column):
        data = self._sheet.cell(row=row, column=column).value
        if data != '' and data != None:
            return True
        return False

    def _write_body(self):
        row = 2
        while self._check(row, 1):
            # add comma
            if row != 2:
                with open(self._output, mode='a') as f:
                    f.write(",")

            data = {}
            for index, column in enumerate(self._columns):
                data[column] = self._sheet.cell(row=row, column=index+1).value

            self._write_data(data)

            row += 1

    def _write_data(self, data):
        pass

    def convert(self):
        self._write_header()
        self._set_columns()
        self._write_body()
        self._write_footer()


class ConvertToJson(ConvertToSomething):
    def _write_data(self, data):
        with open(self._output, mode='a') as f:
            f.write(json.dumps(data, sort_keys=False,
                               ensure_ascii=False, indent=4))


class ConvertToPhpArray(ConvertToSomething):
    def _write_data(self, data):
        with open(self._output, mode='a') as f:
            f.write('[\n    ')
            for key, value in data.items():
                f.write(self.__add_quote(key)+' => ' +
                        self.__format_value(value))
                if key != next(iter(reversed(data))):
                    f.write(',')
            f.write('\n]')

    def __add_quote(self, value):
        return '\"'+str(value)+'\"'

    def __format_value(self, value):
        if isinstance(value, (int, float)):
            return str(value)
        else:
            return self.__add_quote(value)


def main():
    with open('settings.json', 'r') as f:
        settings = json.load(f)

        book = openpyxl.load_workbook(settings['dataFile'])
        sheet = book[settings['dataSheet']]
        output = settings['outputFile']
        output_type = settings['outputFileType']

    if output_type == 'json':
        converter = ConvertToJson(sheet, output)
    elif output_type == 'php':
        converter = ConvertToPhpArray(sheet, output)
    else:
        return False

    converter.convert()


main()
