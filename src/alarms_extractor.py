import re
import xlsxwriter
from dataclasses import dataclass
from codesys_symbols_parser import CodesysSymbolParser


@dataclass
class AlarmType:
    pattern: str
    offset: int


alarm_types = [
    AlarmType(r'Application\.S(\d+)\.stDefImdt\.', 1),
    AlarmType(r'Application\.S(\d+)\.stDefFcy\.', 1000),
    AlarmType(r'Application\.S(\d+)\.stDefAttente\.', 2000)
]


def get_alarm_list(symbols_list):
    alarm_list = {}
    for symbol in symbols_list:
        for alarm_type in alarm_types:
            res = re.search(alarm_type.pattern, symbol['name'])
            if not res:
                continue

            symbol['alarm_offset'] = alarm_type.offset + symbol['byteoffset']

            station_id = res.group(1)
            if station_id in alarm_list:
                alarm_list[station_id].append(symbol)
            else:
                alarm_list[station_id] = [symbol]

            break
    return alarm_list


def write_headers(worksheet):
    headers = [
        "Station",
        "Alarm ID",
        "Symbol",
        "Text"
    ]
    for i, header in enumerate(headers):
        worksheet.write(0, i, header)


def write_alarms(worksheet, symbols_list):
    row_id = 1  # Starts writing at row 2
    alarm_list = get_alarm_list(symbols_list)
    print(f'{len(alarm_list)} alarms found.')
    for station in alarm_list:
        # Sort alarm list by alarm_offset
        alarm_list[station].sort(key=lambda alarm: alarm['alarm_offset'])
        for alarm in alarm_list[station]:
            row_data = [station, alarm['alarm_offset'], alarm['name'], alarm['comment']]
            worksheet.write_row(row_id, 0, row_data)
            row_id += 1


def write_xls(fname, symbols_list):
    with xlsxwriter.Workbook(fname) as workbook:
        worksheet = workbook.add_worksheet()
        write_headers(worksheet)
        write_alarms(worksheet, symbols_list)


if __name__ == '__main__':
    symbols_filepath = '../assets/PZ_PLC.MyController.Application_withoutAttributes.xml'
    parser = CodesysSymbolParser(symbols_filepath)
    parser.parse()
    symbols = parser.get_symbols()

    print(f'{len(symbols)} symbols found.')

    write_xls("../assets/test.xlsx", symbols)
