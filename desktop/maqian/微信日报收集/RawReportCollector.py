# coding=utf-8
import csv
import datetime
import re
import sys

MANAGE_NAMES = [u'何明学', u'何锴鹏', u'ks陈静', u'钟小平', u'黄宋雪', u'胡强', u'王新', u'ks刘宁', u'路小鹿', u'林增玉' u'ks孙涛']


def _is_manage_starts(line):
    for managerName in MANAGE_NAMES + [u'马倩']:
        if re.match(ur'{0}:'.format(managerName), line):
            return True
    return False


def _is_find_chara_in_report(chara, lines):
    for line in lines:
        if line.find(chara) >= 0:
            return True
    return False


def _is_valid_raw_data(lines):
    characters = []
    if lines[0][:-1] not in MANAGE_NAMES:
        return False
    for character in characters:
        if _is_find_chara_in_report(character, lines) is False:
            return False
    return True


def _get_client_name(lines):
    for line in lines:
        m = re.search(ur'(?:供应商名称|客户|客户名|对应厂家|供应商班名称|供应商名|供应商)\s*(?:\:|：)\s*(.+)', line)
        if m:
            return m.group(1)
    return u''


def _get_date(lines):
    dayOrNight = u''
    for line in lines:
        m = re.search(ur'(?:任务进度|任务日期|任务时间|任务时期|检验日期)\s*(?:\:|：)\s*(.+)', line)
        if m:
            rawTestDate = m.group(1)
            testDate = rawTestDate.replace(u'白班', u'').replace(u'夜班', u'').strip()
            if rawTestDate.find(u'白班') >= 0:
                dayOrNight = u'白班'
            if rawTestDate.find(u'夜班') >= 0:
                dayOrNight = u'夜班'
            return re.sub(' ', '', testDate), dayOrNight
    return u'', u''


def _get_workhour(lines):
    for line in lines:
        m = re.search(ur'(?:总计工时|共计\:|任务总工时|工时合计|合计工时|任务总时|总工时|总共时|工时|总工时计|共计时间)\s*(?:\:|：?)\s*([0-9\.]+)', line)
        if m:
            return m.group(1)
    return u''


def _get_service_addr(lines):
    for line in lines:
        m = re.search(ur'(?:服务地点|地点)\s*(?:\:|：)\s*(.+)', line)
        if m:
            return m.group(1)
    return u''


def _get_test_daynight_of_csv_input(lines):
    dayNight = ur''
    for line in lines:
        if line.find(ur'白班') >= 0:
            dayNight = ur'白班'
        if line.find(ur'夜班') >= 0:
            dayNight = ur'夜班'
    return dayNight


def _get_test_date_of_csv_type(line):
    m = re.search(ur'([\d]{4}).+?([\d]{1,2}).+?([\d]{1,2}).?', line)
    if m:
        return m.group(0).replace(' ', '')
    m = re.search(ur'([\d]{1,2}).+([\d]{1,2}).?', line)
    if m:
        return ur'{0}年{1}'.format(datetime.datetime.now().strftime('%Y'), m.group(0)).replace(' ', '')
    return ur''


def write_data_to_csv(rawDataLiness):
    with open(ur'report-马倩.csv', 'wb') as f:
        writer = csv.writer(f)
        heads = [ur'主管', ur'客户名称', ur'日期', ur'日报详情', ur'工时']
        writer.writerow(map(lambda s: s.encode('utf-8'), heads))
        for rawDataLines in rawDataLiness:
            engineerName = re.sub('\d+', '', rawDataLines[0][:-1])
            clientName = _get_client_name(rawDataLines)
            testDate, _ = _get_date(rawDataLines)
            testHours = _get_workhour(rawDataLines)
            reportDetail = u'\n'.join(rawDataLines[1:])
            rowLine = [engineerName, clientName, testDate, reportDetail, testHours]
            rowLine = map(lambda s: s.encode('utf-8'), rowLine)
            writer.writerow(rowLine)


def collect_raw_data():
    splitLines = []
    file = open('data.txt', 'r')
    lines = file.readlines()
    curManageMsgLines = []
    lines = map(lambda s: s.strip().decode('utf-8'), lines)
    for line in lines:
        if _is_manage_starts(line):
            if curManageMsgLines:
                splitLines.append(curManageMsgLines)
                curManageMsgLines = []
        line = line.replace(u'\xa0', ' ')
        curManageMsgLines.append(line)
    if curManageMsgLines:
        splitLines.append(curManageMsgLines)
    rawDataLiness = filter(lambda lines: _is_valid_raw_data(lines) is True, splitLines)
    write_data_to_csv(rawDataLiness)


def collect_raw_data_from_csv():
    with open(ur'日报输入.csv', 'r') as f:
        reader = csv.reader(f)
        rows = [row for row in reader]
    with open(ur'日报输出-孙丹.csv', 'wb') as f:
        writer = csv.writer(f)
        heads = [ur'主管', ur'客户名称', ur'日期', ur'日报详情', ur'工时']
        writer.writerow(map(lambda s: s.encode('utf-8'), heads))
        for row in rows[1:]:
            reportDetail = row[0].decode('gbk').replace(u'\xa0', ' ').replace('?', '')
            rawDataLines = reportDetail.split('\n')
            engineerName = ur''
            clientName = _get_client_name(rawDataLines)
            testDate = _get_test_date_of_csv_type(filter(lambda s: len(s) > 0, rawDataLines[:3])[0])
            # dayOrNight = _get_test_daynight_of_csv_input(rawDataLines)
            testHours = _get_workhour(rawDataLines)
            # serviceAddr = _get_service_addr(rawDataLines)
            rowLine = [engineerName, clientName, testDate, reportDetail, testHours]
            rowLine = map(lambda s: s.encode('utf-8'), rowLine)
            writer.writerow(rowLine)


if __name__ == '__main__':
    collect_raw_data_from_csv()
    if len(sys.argv) == 1:
        collect_raw_data()
    else:
        collect_raw_data_from_csv()
