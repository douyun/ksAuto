#encoding=utf-8
from CommCollector import CommCollector

class YibinKedaliCollector(CommCollector):
    
    def __init__(self):
        CommCollector.__init__(self)

    @property
    def specailMatchPattern(self):
        return {}

    @property
    def endPattern(self):
        return u'不良'

if __name__ == '__main__':
    YibinKedaliCollector().main()
    
