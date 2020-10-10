a = [(1,'asd',1), (1,'qwe',2),(2,'asd',1), (2,'qwe',2)]
f = []
tempdict, templist = dict(), list()
for i, j, k in a:
    tempdict[str(i)+'_'+str(k)] = {"identity": j, "i_name": k}
    templist.append(i)
templist = list(set(templist))
for i in templist:
    j, k = "", ""
    if str(i)+'_1' in tempdict:
        j = tempdict[str(i)+'_1']['identity']
    if str(i)+'_2' in tempdict:
        k = tempdict[str(i)+'_2']['identity']
    f.append((i, j, k))

pass