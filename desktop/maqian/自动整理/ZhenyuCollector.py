#encoding=utf-8
from CommCollector import CommCollector

class ZhenyuCollector(CommCollector):
    
    def __init__(self):
        CommCollector.__init__(self)

    @property
    def specailMatchPattern(self):
        return {}

if __name__ == '__main__':
    ZhenyuCollector().main()
    
