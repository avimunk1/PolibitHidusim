class myclass(object):
    """docstring for ."""

    def __init__(self,strt):
        myobjid=id(self)
        self.id=myobjid
        self.a=strt+1
        self.b='abc'
        #print('init',self.id,self.a,self.b)

#myid=id(sumeraPoliciesObj)

mydic=dict()

i=1
while i<10:
    i=i+1
    x= i
    x=myclass(i)
    mydic[i]=x
    #print('this is i',i)
c=2
while c<10:
    c=c+1
    #print('this is c',c)
    zz=mydic[5]
    print('zz',zz.id,zz.a,zz.b)
