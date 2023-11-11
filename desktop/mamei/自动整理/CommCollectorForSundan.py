# coding=utf-8
import csv
import re
PATTERN = {'testDate': u'(?:任务进度|任务日期|任务时间|任务时期|检验日期)',
            'vendorName': u'(?:供应商名称|客户|客户名)',
            'serviceAddr': u'(?:服务地点|地点)',
            'serviceDetail': u'(?:服务明细|检验项目|挑选内容|启动原因|项目挑选内容)',
            'partName': u'^(?:产品名称|零件名称|名称|物料名称|物料描述)',
            'partNo': u'^(?:物料P.?N|零件号|产品物料|物料号|图号|型号|物料PN|物料:|物料号P.?N)',
            'batchNo': u'^(?:批次|产品批次|批次号|管理号|批次B.?N)',
            'vendorNo': u'^(?:供应商批次|供应商次|供应商批|供应商批次号)',
            'testSum': u'^(?:数量|检验数量|检查数量|电检数量|总数量|扫码数量|排查数量|返工数量|分拣数量|分检数量|擦拭数量|挑选数量)',
            'okSum': u'^(?:良品数量|合格数量|合格总数|合格品|良品总数量)',
            'badSum': u'(?:不良品数量|不良品数|不良数量|不良品)',
            'badType': u'不良(?:原因|类型|原因变形|内容)',
            'testPerson': u'(?:服务人员|任务人员|作业人员|检验人|派工人|派工人员)',
            'taskDate': u'(?:任务时间|检验时间|时间)',
            'testHour': u'(?:总计工时|共计|任务总工时|工时合计|合计工时|任务总时|总工时)',
            'remark': u'备注'
            }


def match_str(pattern, s):
    m = re.search(pattern + u'\s*(?:[:：；]+)\s*(.*)', s)
    m = m or re.search(pattern + u'\s*(?:[:：；]*)\s*(.*)', s)
    return m.group(1).strip().replace('?', '') if m else ''


class CommCollectorForSundan():

    def __init__(self):
        self._rowDict = {}
        self.clientOrderNoDict = {}
        self.clientChargeRateDict = {}
        for k, v in self.specailMatchPattern.items():
            PATTERN.update({k: v})
        for k in PATTERN:
            self._rowDict[k] = ''
        self._init_clientname_orderno_dict()
        self._init_clientname_charge_rate_dict()

    def _init_clientname_orderno_dict(self):
        f = open('clientOrder.csv')
        for row in csv.DictReader(f):
            self.clientOrderNoDict.update({row[u'付款客户'.encode('gbk')]: row[u'K&S订单号'.encode('gbk')]})
        f.close()

    def _init_clientname_charge_rate_dict(self):
        with open('clientOrder.csv') as f:
            for row in csv.DictReader(f):
                self.clientChargeRateDict.update({row[u'付款客户'.encode('gbk')]: row[u'费率'.encode('gbk')]})

    def find_order_no(self, clientName):
        for client, orderNo in self.clientOrderNoDict.items():
            if client.find(clientName) >= 0:
                return orderNo
        return ''

    def find_charge_rate(self, clientName):
        for client, chargeRate in self.clientChargeRateDict.items():
            if client.find(clientName) >= 0:
                return chargeRate
        return ''

    def _process_num(self, num):
        m = re.search('(\d+)', num)
        if m:
            return m.group(1)
        return num

    @property
    def specailMatchPattern(self):
        return {}

    @property
    def endPattern(self):
        return u'不良品数'

    def _get_workhour(self, line):
        m = re.search(ur'([0-9\.]+)', line)
        if m:
            return m.group(1)
        return u''

    def _process_task_date(self, line):
        dates = line.split('##')[1:]
        taskDates = []
        for date in dates:
            m = re.search(ur'(\d+(?:\:|：|\.)\d+\s*(?:\-|\-|\~)\s*\d+(?:\:|：|\.)\d+)', date)
            if m:
                taskDates.append(m.group(1).strip().replace(u'：', u':').replace(u'~', u'-').replace(u'.', u':').replace(u' ', ''))
        if len(taskDates) > 1:
            taskDates = map(lambda taskDate: taskDate + '()', taskDates)
        result = ' '.join(taskDates)
        return result

    def get_matched_info(self, patternItem, matchedStr):
        if patternItem == 'testDate':
            dateStr = match_str(PATTERN.get(patternItem), matchedStr) or self._rowDict.get(patternItem, '')
            if len(dateStr) == 0:
                m = re.search(u'\d+年\d+月\d+日', matchedStr)
                return m.group(0) if m else ''
        return match_str(PATTERN.get(patternItem), matchedStr) or self._rowDict.get(patternItem, '')

    def _decode_row(self, row):
        try:
            return row.decode('gbk')
        except:
            return row.decode('utf-8')

    def collect_all_rows(self):
        testItems = []
        f = open('input.csv', 'r')
        reader = csv.reader(f)
        rows = [row for row in reader]
        i = 1
        for row in rows[1:]:
            rowStr = self._decode_row(row[0])
            flag = False
            
            for k in self._rowDict:
                self._rowDict[k] = ''
            for line in rowStr.split('\n'):
                eachItemNames = ['vendorName', 'testDate', 'partName', 'partNo', 'batchNo', 'vendorNo', 'testSum', 'okSum', 'badSum', 'badType', 'serviceDetail', 'serviceAddr']
                for item in eachItemNames:
                    self._rowDict[item] = self.get_matched_info(item, line)
                
                if self._rowDict['testSum'] != '' and line.find(u'数量') == -1 and line.find(self.endPattern) == -1:
                    testItems.append(map(lambda k: self._rowDict[k], eachItemNames))
                    for item in ['batchNo', 'testSum', 'vendorNo', 'okSum', 'badSum', 'badType']:
                        self._rowDict[item] = ''
        
                self._rowDict['testPerson'] = self._rowDict['testPerson'] or match_str(PATTERN.get('testPerson'), line)
                if re.search(PATTERN.get('testPerson'), line):
                    flag = True
                    continue
                if flag and match_str(PATTERN.get('taskDate'), line) == '' and match_str(PATTERN.get('testHour'), line) == '':
                    self._rowDict['testPerson'] += line
                
                self._rowDict['taskDate'] = self.get_matched_info('taskDate', line)
                self._rowDict['testHour'] = self.get_matched_info('testHour', line)
                if self._rowDict['taskDate'] or self._rowDict['testHour']:
                    flag = False
                if self._rowDict['taskDate'] and not self._rowDict['testHour']:
                    self._rowDict['taskDate'] += '##' + line
                self._rowDict['remark'] = ''
            taskHourflag = True
            for testItem in testItems:
                if len(testItem) == len(eachItemNames):
                    if taskHourflag:
                        testItem.extend([self._rowDict['testPerson'], self._rowDict['taskDate'], self._get_workhour(self._rowDict['testHour']), self._rowDict['remark'],
                                         self.find_charge_rate(self._rowDict['vendorName'].encode('gbk'))])
                        taskHourflag = False
                        testItem[-4] = self._process_task_date(testItem[-4])
                    testItem[6] = self._process_num(testItem[6])
                    testItem[7] = self._process_num(testItem[7])
                    testItem[8] = self._process_num(testItem[8])
                    testItem.insert(1, str(i))
                    testItem.insert(0, self.find_order_no(self._rowDict['vendorName'].encode('gbk')))
            i += 1
        return testItems

    def save_collect_result(self, testItems):
        f = open('output.csv', 'wb+')
        writer = csv.writer(f)
        testItems = [[u'订单号', u'供应商名称', u'序号', u'检验日期', u'零件名称', u'零件号', u'批次号', u'供应商批次', u'检验总数', u'合格数量', u'不良数量', u'不良类型', u'服务明细', u'服务地点',
                      u'任务人员', u'任务时间', u'任务工时', u'备注', u'费率']] + testItems
        for testItem in testItems:
            testItem = map(lambda i: i.encode('gbk'), testItem)
            writer.writerow(testItem)
        f.close()

    def main(self):
        results = self.collect_all_rows()
        self.save_collect_result(results)


if __name__ == '__main__':
    CommCollectorForSundan().main()
