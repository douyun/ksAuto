# _*_ coding: utf-8 -*
import csv
import datetime
import time
import re
from KsAuto import KsAuto


def format_date(rawDate):
    m = re.search(ur'([\d]{4}).+?([\d]{1,2}).+?([\d]{1,2})', rawDate)
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime('%Y-%m-%d')
    return rawDate


def format_batch_no(batchNo):
    if re.search(ur'^[0-9]+$', batchNo):
        return str(batchNo.zfill(10))
    return batchNo


def main():
    ks = KsAuto()
    ks.to_common_page()

    f = open('output.csv')
    orderDetailDict = {}

    reportLines = []
    reportLines.append(['orderNo', 'result', 'faileDetail'])


    for row in csv.DictReader(f):
        orderNo = row[u'订单号'.encode('gbk')]
        orderDetailDict.setdefault(orderNo, []).append(row)


    for orderNo, orderInfoDicts in orderDetailDict.items():
        try:
            ks.query_order(orderNo)
            time.sleep(0.5)
            ks.wait_query_order_finished()
            time.sleep(0.5)
            startIndex, endIndex = ks.add_new_row(len(orderInfoDicts))
            if startIndex == -1:
                reportLines.append([orderNo, 'failed', 'add new row failed'])
                continue
            index = 0
            for i in range(startIndex, endIndex + 1):
                orderInfoDict = orderInfoDicts[index]
                ks.update_row(i, 1, format_date(orderInfoDict[u'检验日期'.encode('gbk')].decode('gbk')))
                ks.update_row(i, 2, orderInfoDict[u'零件名称'.encode('gbk')].decode('gbk'))
                ks.update_row(i, 3, orderInfoDict[u'零件号'.encode('gbk')].decode('gbk'))
                ks.update_row(i, 4, format_batch_no(orderInfoDict[u'批次号'.encode('gbk')].decode('gbk')))
                ks.update_row(i, 5, orderInfoDict[u'检验总数'.encode('gbk')].decode('gbk'))
                ks.update_row(i, 9, orderInfoDict[u'任务人员'.encode('gbk')].decode('gbk'))
                ks.update_row(i, 10, orderInfoDict[u'任务时间'.encode('gbk')].decode('gbk'))
                ks.update_row(i, 12, orderInfoDict[u'任务工时'.encode('gbk')].decode('gbk'))
                ks.update_row(i, 13, orderInfoDict[u'备注'.encode('gbk')].decode('gbk'))
                index += 1
            ks._double_submit()

            ks.goto_expense_report()
            expenseRows = ks.get_expense_row_element(len(orderInfoDicts))
            for i in range(len(orderInfoDicts)):
                expenseRow = expenseRows[i].get_attribute('tridx')
                orderInfoDict = orderInfoDicts[i]
                ks.update_row(expenseRow, 3, orderInfoDict[u'服务明细'.encode('gbk')].decode('gbk'))
                ks.update_row(expenseRow, 4, orderInfoDict[u'服务地点'.encode('gbk')].decode('gbk'))
                if len(orderInfoDict[u'费率'.encode('gbk')]) > 0:
                    ks.update_row(expenseRow, 10, orderInfoDict[u'费率'.encode('gbk')].decode('gbk') + '\n', False)
            ks.submit_on_expense_page()
            ks.refresh_by_click_button()
            print 'success:' + '-' * 30 + orderNo + '-' * 30
            reportLines.append([orderNo, 'success', ''])
        except Exception as e:
            print 'failed:' + '-' * 30 + orderNo + '-' * 30
            try:
                reportLines.append([orderNo, 'failed', e.message])
                ks.refresh_by_click_button()
            except:
                ks.refresh()
                ks.to_common_page()
            time.sleep(1)


    f = open('./report/report-{0}.csv'.format(datetime.datetime.now().strftime('%Y%m%d%H%M%S')), 'wb+')
    writer = csv.writer(f)
    writer.writerows(reportLines)
    f.close()


if __name__ == '__main__':
    main()
