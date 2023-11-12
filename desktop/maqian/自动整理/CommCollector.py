# coding=utf-8
import csv
import re
PATTERN = {'testDate': u'(?:任务进度|任务日期|任务时间|任务时期|检验日期)',
            'vendorName': u'(?:供应商班名称|供应商名称|客户名称|供应商名|供应商(?::|：))',
            'serviceAddr': u'(?:服务地点|地点|检验基地)',
            'serviceDetail': u'(?:服务明细|检验项目|挑选内容)',
            'partName': u'(?:产品名称|零件名称|物料描述|检验车型)',
            'partNo': u'^(?:物料P.?N|零件号|产品物料|物料(?::|：)|图号|型号|物料PN|物料号)',
            'batchNo': u'^(?:批次|产品批次|批次号|管理号|物料批次|批次B.?N)',
            'vendorNo': u'^(?:供应商批次|供应商次|供应商批)',
            'testSum': u'^(?:数量|检验数量|检查数量|电检数量|未检发货数量|清洁数量|排查数量|分选数量|返工数量|拆包数量|擦拭数量|送检数量|打磨数量|测试数量|拆包数)',
            'okSum': u'^(?:良品数量|合格数量|合格总数)',
            'badSum': u'(?:不良品数量|不良品数|不良数量|不良品)',
            'badType': u'不良(?:原因|类型|原因变形|内容)',
            'testPerson': u'(?:服务人员|任务人员|作业人员|检验人员|检验人)',
            'taskDate': u'(?:任务时间|检验时间)',
            'testHour': u'(?:总计工时|共计:|任务总工时|工时合计|合计工时|任务总时|总工时|总共时|工时|总工时计|任务共计|共计工时|检验有效工时|共计时间)',
            'remark': u'备注',
            'dayNight': u'(白班|夜班)'
            }


def match_str(pattern, s):
    m = re.search(pattern + u'\s*(?:[:：；]+)\s*(.*)', s)
    m = m or re.search(pattern + u'\s*(?:[:：；]*)\s*(.*)', s)
    return m.group(1).strip().replace('?', '') if m else ''


def Singleton(cls):
    instance = {}

    def _singleton_wrapper(*args, **kargs):
        if cls not in instance:
            instance[cls] = cls(*args, **kargs)
        return instance[cls]

    return _singleton_wrapper


@Singleton
class OrderNoSearcher(object):

    def __init__(self):
        self.clientOrderNoDict = {}
        self._init_client_info_dict()

    def _init_client_info_dict(self):
        f = open('clientOrder.csv')
        for row in csv.DictReader(f):
            key = (row[u'付款客户'.encode('gbk')], row[u'服务地点'.encode('gbk')], row[u'零件号'.encode('gbk')], row[u'零件名称'.encode('gbk')])
            value = (row[u'K&S订单号'.encode('gbk')], row[u'费率'.encode('gbk')], row[u'录入方式'.encode('gbk')])
            self.clientOrderNoDict.update({key: value})
        f.close()

    def find_orderno_and_chargerate(self, clientName, serviceAddr, patchNo, patchName):
        filteredMatchDict = self._filter_match_info(clientName)
        if len(filteredMatchDict) == 0:
            return '', '', ''
        if len(filteredMatchDict) == 1:
            return filteredMatchDict.values()[0]
        for matchInput, matchInputIndex in [(serviceAddr, 1), (patchNo, 2), (patchName, 3)]:
            filteredMatchDict = self.__filter_match_dict_by_extral_info(filteredMatchDict, matchInput, matchInputIndex)
            if len(filteredMatchDict) == 1:
                return filteredMatchDict.values()[0]
        return '', '', ''

    def __filter_match_dict_by_extral_info(self, filteredMatchDict, matchInput, matchInputIndex):
        tempFilteredMatchDict = {}
        mathedKeys = map(lambda key: key[matchInputIndex], filteredMatchDict.keys())
        mathedKeys = filter(lambda mathedKey: len(mathedKey) > 0, mathedKeys)
        if len(mathedKeys) == 0:
            return filteredMatchDict
        for k, v in filteredMatchDict.items():
            if len(k[matchInputIndex]) > 0:
                if matchInput.find(k[matchInputIndex]) >= 0:
                    tempFilteredMatchDict.update({k: v})
            else:
                isMatched = False
                for mathedKey in mathedKeys:
                    if matchInput.find(mathedKey) >= 0:
                        isMatched = True
                        break
                if not isMatched:
                    tempFilteredMatchDict.update({k: v})
        return tempFilteredMatchDict


    def _filter_match_info(self, clientName):
        if len(clientName) == 0:
            return {}
        filteredMatchDict = {}
        for key, value in self.clientOrderNoDict.items():
            if key[0].find(clientName) >= 0:
                filteredMatchDict.update({key: value})
        return filteredMatchDict


class CommCollector(object):

    def __init__(self):
        self._rowDict = {}
        self.clientOrderNoDict = {}
        self.clientChargeRateDict = {}
        for k, v in self.specailMatchPattern.items():
            PATTERN.update({k: v})
        for k in PATTERN:
            self._rowDict[k] = ''

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

    def _format_test_date(self, testDate):
        testDate = testDate.replace(' ', '').replace(u' ', u'')
        m = re.search(u'(\d+年\d+月\d+(?:日|号))',  testDate)
        if m:
            return m.group(1)
        m = re.search(u'(\d+\.\d+\.\d+)', testDate)
        if m:
            return m.group(1)
        return testDate

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
                eachItemNames = ['vendorName', 'testDate', 'partName', 'partNo', 'batchNo', 'vendorNo', 'testSum', 'okSum', 'badSum', 'badType', 'serviceDetail', 'serviceAddr', 'dayNight']
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
                                         self.get_order_no(self._rowDict, 1)])
                        testCountsOfAnOrder = 0
                        overallItem = testItem
                        taskHourflag = False
                        try:
                            testHour = float(self._get_workhour(self._rowDict['testHour']))
                        except:
                            testHour = 0
                        testItem[-4] = self._process_task_date(testItem[-4])
                    testItem[6] = self._process_num(testItem[6])
                    testItem[7] = self._process_num(testItem[7])
                    testItem[8] = self._process_num(testItem[8])
                    testItem[1] = self._format_test_date(testItem[1])

                    testCountsOfAnOrder += int(testItem[6])
                    testItem.insert(1, str(i))
                    testItem.insert(0, self.get_order_no(self._rowDict, 0).decode('gbk'))
                    testItem.insert(2, self.get_order_no(self._rowDict, 2).decode('gbk'))
            if not taskHourflag:
                overallItem.append(self._count_test_efficiency(testCountsOfAnOrder, testHour))
            i += 1
        return testItems

    def _count_test_efficiency(self, testCountsOfAnOrder, testHour):
        if testHour == 0:
            return ''
        return str(round(testCountsOfAnOrder / testHour, 2)).decode('gbk')

    def get_order_no(self, rowDict, index=0):
        return OrderNoSearcher().find_orderno_and_chargerate(rowDict['vendorName'].encode('gbk'),
                                                             rowDict['serviceAddr'].encode('gbk'),
                                                             rowDict['partNo'].encode('gbk'),
                                                             rowDict['partName'].encode('gbk'))[index]

    def save_collect_result(self, testItems):
        f = open('output.csv', 'wb')
        writer = csv.writer(f)
        testItems = [[u'订单号', u'供应商名称', u'录入方式', u'序号', u'检验日期', u'零件名称', u'零件号', u'批次号', u'供应商批次', u'检验总数', u'合格数量', u'不良数量', u'不良类型', u'服务明细', u'服务地点', u'班次',
                      u'任务人员', u'任务时间', u'任务工时', u'备注', u'费率', u'效率']] + testItems
        for testItem in testItems:
            testItem = map(lambda i: i.encode('gbk'), testItem)
            writer.writerow(testItem)
        f.close()

    def main(self):
        results = self.collect_all_rows()
        self.save_collect_result(results)


if __name__ == '__main__':
    CommCollector().main()
