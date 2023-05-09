
class exapp:
    pass

ts=[]
ee = []
ng= []
for i in dir(exapp):
    try:
        tt = getattr(exapp, i)
        if str(type(tt)).startswith('''<class 'win32com.'''):
            ts.append(i)
        else:
            ng.append(i+str(type(tt)))
    except:
        ee.append(i)

print(len(ts), len(ee), len(ng))


class AA:
     v23=234234
     ''' asdf  '''



class DD:
    def __get__(self, ii):
        print(ii)
        return 1
    def __call__(self, ii):
        '''return item '''
        print(ii)
        return 1

dd = DD()
dd(3)

