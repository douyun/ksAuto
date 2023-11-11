# coding=utf-8
from CommCollectorForSundan import CommCollectorForSundan


class YibinKedaliCollectorForSundan(CommCollectorForSundan):
    
    def __init__(self):
        CommCollectorForSundan.__init__(self)

    @property
    def specailMatchPattern(self):
        return {}

    @property
    def endPattern(self):
        return u'不良'


if __name__ == '__main__':
    YibinKedaliCollectorForSundan().main()
    
