#encoding=utf-8
from CommCollector import CommCollector

class JintaiCollector(CommCollector):
    
    def __init__(self):
        CommCollector.__init__(self)

    @property
    def specailMatchPattern(self):
        return {
                'partNo': u'(?:物料 P.N|零件号|产品物料|物料|物料编码)',
                'batchNo': u'^(?:批次|产品批次|批次号|产品批|批次)',
                }

    @property
    def endPattern(self):
        return u'不良'

if __name__ == '__main__':
    JintaiCollector().main()
    
