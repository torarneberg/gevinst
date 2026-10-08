import openpyxl, sys, json, statistics as st, collections, random
S=sys.argv[1]
random.seed(1)
def rows(name):
    ws=openpyxl.load_workbook(f'{S}/src/{name}.xlsx').worksheets[0]
    return list(ws.iter_rows(values_only=True))
def num(v):
    try: return int(v)
    except: return None
def areas(v): return [('Regulatorisk produktinformasjon' if a.strip()=='Regulatorisk PI' else a.strip()) for a in (v or '').split(';') if a.strip()]
SHORT=['Faglig støtte','Finne faglig info','Ekstern nasjonal samhandling','Status i saksbehandling','Arbeide effektivt','Enkle å bruke','Prioritering/ressursstyring','Intern samhandling','Oversikt frister','Oversikt oppgaver']
MI1={'Kvalitet':[0,1],'Samhandling nasjonal':[2],'Medarbeidertilfredshet':[3,4,5,6,7,8,9]}
out={}
# ---- R1
r22=rows('08c8a52d-R1september2022'); r23=rows('04e79576-R1mai2023'); r26=rows('d32ea74a-R1oktober2026')
F=['Apotekkonsesjoner','Tilsyn','Virksomhetstillatelser']
R1={'2022':{},'2023':{},'2026':{}}
for lab,rr in (('2022',r22),('2023',r23)):
    for j,f in enumerate(F): R1[lab][f]=[float(r[j+1]) for r in rr[1:]]
h=r26[0]; data=r26[1:]
ns={}
for f in F:
    sub=[r for r in data if f in areas(r[6])]; ns[f]=len(sub)
    R1['2026'][f]=[round(st.mean([num(r[q]) for r in sub if num(r[q])]),2) for q in range(8,18)]
R1['2026']['Alle']=[round(st.mean([num(r[q]) for r in data if num(r[q])]),2) for q in range(8,18)]
def mi(v,idx): return round(st.mean([v[i] for i in idx]),2)
out['R1']={'q':SHORT,'vals':R1,'n2026':ns,'nTotal':len(data),
  'mi':{y:{f:{m:mi(R1[y][f],ix) for m,ix in MI1.items()} for f in F} for y in R1}}
# R1 dist 2026 alle (andel 1-2, 3-4, 5-6)
out['R1']['dist2026']=[[sum(1 for r in data if num(r[q]) and lo<=num(r[q])<=hi)/len(data) for lo,hi in ((1,2),(3,4),(5,6))] for q in range(8,18)]
# ---- R2
SHORT2=['Info til eksterne brukere']+SHORT[:3]+['Ekstern internasjonal samhandling']+SHORT[3:]
MI2={'Informasjonstilgjengelighet':[0],'Kvalitet':[1,2],'Samhandling nasjonal':[3],'Samhandling internasjonal':[4],'Medarbeidertilfredshet':[5,6,7,8,9,10,11]}
A=rows('76d8d138-R2sommer2025'); B=rows('0025ce6a-R2oktober2026')
GROUPS={'Regulatorisk produktinformasjon':['Regulatorisk produktinformasjon'],'Regulatorisk før MT':['Regulatorisk før MT'],'Regulatorisk etter MT':['Regulatorisk etter MT'],
 'Sikkerhet – HUM':['Sikkerhet - HUM'],'Effekt – HUM':['Effekt - HUM'],'Kvalitet (kjemisk/biologisk)':['Kvalitet - Kjemisk','Kvalitet - Biologisk'],
 'Preklinikk, farmakologi og VET':['Preklinikk','Farmakologi - BE/AM','Farmakologi - Klinisk','Effekt og sikkerhet - VET']}
def stats(sub,q):
    v=[num(r[q]) for r in sub]; vv=[x for x in v if x]
    ir=sum(1 for r in sub if r[q]=='Ikke relevant')
    if not vv: return None
    return {'n':len(vv),'m':round(st.mean(vv),2),'sd':round(st.stdev(vv),2) if len(vv)>1 else 0,'enig':round(sum(x>=5 for x in vv)/len(vv),3),'uenig':round(sum(x<=2 for x in vv)/len(vv),3),'ir':ir,'irShare':round(ir/len(sub),3)}
def perm(a,b,n=20000):
    obs=st.mean(b)-st.mean(a); pool=a+b; k=0
    for _ in range(n):
        random.shuffle(pool); d=st.mean(pool[len(a):])-st.mean(pool[:len(a)])
        if abs(d)>=abs(obs)-1e-9: k+=1
    return k/n
R2={}
for lab,rr in (('2025',A),('2026',B)):
    d=rr[1:]; res={'Alle':[stats(d,q) for q in range(8,20)]}
    for g,members in GROUPS.items():
        sub=[r for r in d if set(areas(r[6]))&set(members)]
        res[g]=[stats(sub,q) for q in range(8,20)]; res[g+'_n']=len(sub)
    for role in ('saksbehandler/lagleder','enhetsleder'):
        sub=[r for r in d if r[7].lower()==role]; res['ROLE:'+role]=[stats(sub,q) for q in range(8,20)]; res['ROLE:'+role+'_n']=len(sub)
    R2[lab]=res
p=[]
for i,q in enumerate(range(8,20)):
    a=[num(r[q]) for r in A[1:] if num(r[q])]; b=[num(r[q]) for r in B[1:] if num(r[q])]
    p.append(round(perm(a,b),3))
mi2={}
for lab in R2:
    mi2[lab]={}
    for g in ['Alle']+list(GROUPS)+['ROLE:saksbehandler/lagleder','ROLE:enhetsleder']:
        mi2[lab][g]={m:round(st.mean([R2[lab][g][i]['m'] for i in ix if R2[lab][g][i]]),2) for m,ix in MI2.items() if any(R2[lab][g][i] for i in ix)}
out['R2']={'q':SHORT2,'vals':R2,'p':p,'mi':mi2,'n':{'2025':len(A)-1,'2026':len(B)-1}}
# free-text vs quant
def ft(rr,first,last):
    d=rr[1:]; c=[];nc=[]
    for r in d:
        vv=[num(r[q]) for q in range(first,last)]; vv=[x for x in vv if x]
        (c if r[-1] else nc).append(st.mean(vv))
    return {'withComment':len(c),'meanC':round(st.mean(c),2) if c else None,'meanNC':round(st.mean(nc),2) if nc else None}
out['ft']={'R1_2026':ft(r26,8,18),'R2_2025':ft(A,8,20),'R2_2026':ft(B,8,20)}
json.dump(out,open(f'{S}/work/data.json','w'),ensure_ascii=False,indent=1)
print(json.dumps(out['R1']['mi'],ensure_ascii=False,indent=0)); print(out['R1']['n2026'])
print('p',list(zip(SHORT2,p)))
for g in mi2['2025']: print(g, R2['2025'].get(g+'_n'), R2['2026'].get(g+'_n'), mi2['2025'][g], mi2['2026'][g])
print(out['ft'])
