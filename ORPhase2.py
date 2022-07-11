import pulp as pl
import pandas as pd


c = [i for i in range(1, 21)]
f = [j for j in range(1, 9)]
m = [1, 2, 3]

Model = pl.LpProblem('Transportation', pl.LpMinimize)

# تعریف متعیرها
X1 = pl.LpVariable.dicts('X1' ,(c ,f ,m) ,lowBound=0 ,cat = 'Integer')
A1 = pl.LpVariable.dicts('A1' ,(c ,f ,m) ,lowBound=0 ,cat = 'Integer')
B1 = pl.LpVariable.dicts('B1' ,(c ,f ,m) ,lowBound=0 ,cat = 'Integer')

X2 = pl.LpVariable.dicts('X2' ,(c ,f ,m) ,lowBound=0 ,cat = 'Integer')
A2 = pl.LpVariable.dicts('A2' ,(c ,f ,m) ,lowBound=0 ,cat = 'Integer')
B2 = pl.LpVariable.dicts('B2' ,(c ,f ,m) ,lowBound=0 ,cat = 'Integer')

X3 = pl.LpVariable.dicts('X3' ,(c ,f) ,lowBound=0 ,cat = 'Integer')
A3 = pl.LpVariable.dicts('A3' ,(c ,f) ,lowBound=0 ,cat = 'Integer')
B3 = pl.LpVariable.dicts('B3' ,(c ,f) ,lowBound=0 ,cat = 'Integer')

N = pl.LpVariable.dicts('N' ,f ,lowBound=0 ,cat = 'Integer')
Y = pl.LpVariable.dicts('Y' ,f ,0 ,1 ,cat = 'Binary')
K1 = pl.LpVariable.dicts('K1' ,(c ,f) ,0 ,1 ,cat = 'Binary')
K2 = pl.LpVariable.dicts('K2' ,(c ,f) ,0 ,1 ,cat = 'Binary')
K3 = pl.LpVariable.dicts('K3' ,(c ,f) ,0 ,1 ,cat = 'Binary')

path = 'C:/Users/pc/Desktop/Danial/University/Term 4/او آر/پروژه/OR1 Project Data.xlsx'
C1 = pd.read_excel(path, sheet_name = 'هزینه حمل ونقل بین c و f ')
C2 = pd.read_excel(path, sheet_name = 'هزینه حمل ونقل بین m و f ')
C3 = pd.read_excel(path, sheet_name = 'هزینه نصب یک واحدتجهیز در مرکزf')
C4 = pd.read_excel(path, sheet_name = 'تقاضای مرجوعی ')

C1.columns = [0, 1, 2, 3, 4, 5, 6, 7, 8]
del C1[0]
C1 = C1.set_index([pd.Index([i for i in range(1, 21)])])

C2.columns = [0, 1, 2, 3]
del C2[0]
C2 = C2.set_index([pd.Index([i for i in range(1, 9)])])

C3.columns = [0, 1]
del C3[0]
C3 = C3.set_index([pd.Index([i for i in range(1, 9)])])

C4.columns = [0, 1, 2, 3]
del C4[0]
C4 = C4.set_index([pd.Index([i for i in range(1, 21)])])

# تابع هدف
ans = 0
for i in c:
    for j in f:
        for k in m:
           ans += C1[j][i] * (2*(X1[i][j][k] +X2[i][j][k] +A1[i][j][k] +A2[i][j][k] +B1[i][j][k] +B2[i][j][k]) + 2*(X3[i][j] + A3[i][j] + B3[i][j])) + C2[k][j]*(X1[i][j][k] + X2[i][j][k] + A1[i][j][k] + A2[i][j][k] + B1[i][j][k] + B2[i][j][k]) + C3[1][j]*N[j]
Model += ans

# اختصاص مقادیر تقاضای مرجوعی به متعیرها
for i in c:
    sum1 = 0
    sum2 = 0
    sum3 = 0
    for j in f:
        for k in m:
            sum1 += X1[i][j][k] + X2[i][j][k]
            sum2 += A1[i][j][k] + A2[i][j][k]
            sum3 += B1[i][j][k] + B2[i][j][k]
    for j in f:
        sum1 += X3[i][j]
        sum2 += A3[i][j]
        sum3 += B3[i][j]
    Model += sum1 == C4[1][i]
    Model += sum2 == C4[2][i]
    Model += sum3 == C4[3][i]
    
# محدودیت‌ها
for i in c:
    for j in f:
        for k in m:
            Model += X1[i][j][k] <= 1000000*Y[j]
            Model += X2[i][j][k] <= 1000000*Y[j]
            Model += X3[i][j] <= 1000000*Y[j]
            Model += A1[i][j][k] <= 1000000*Y[j]
            Model += A2[i][j][k] <= 1000000*Y[j]
            Model += A3[i][j] <= 1000000*Y[j]
            Model += B1[i][j][k] <= 1000000*Y[j]
            Model += B2[i][j][k] <= 1000000*Y[j]
            Model += B3[i][j] <= 1000000*Y[j]
sum1 = 0
sum2 = 0
sum3 = 0
sum4 = 0
sum5 = 0
sum6 = 0
sum7 = 0
sum8 = 0
sum9 = 0
for i in c:
     for j in f:
        for k in m:
             sum1 += X1[i][j][k]
             sum2 += X2[i][j][k]
             sum4 += A1[i][j][k]
             sum5 += A2[i][j][k]
             sum7 += B1[i][j][k]
             sum8 += B2[i][j][k]
for i in c:
     for j in f:
         sum3 += X3[i][j]
         sum6 += A3[i][j]
         sum9 += B3[i][j]
Model += (sum2 + sum3) * 0.3 == 0.7 * sum1
Model += (sum1 + sum3) * 0.2 == 0.8 * sum2
Model += (sum1 + sum2) * 0.5 == 0.5 * sum3
Model += (sum5 + sum6) * 0.3 == 0.7 * sum4
Model += (sum4 + sum6) * 0.2 == 0.8 * sum5
Model += (sum4 + sum5) * 0.5 == 0.5 * sum6
Model += (sum8 + sum9) * 0.3 == 0.7 * sum7
Model += (sum7 + sum9) * 0.2 == 0.8 * sum8
Model += (sum7 + sum8) * 0.5 == 0.5 * sum9

for i in c:
    for j in f:
        Model += K1[i][j] <= 1000000*Y[j]
        Model += K2[i][j] <= 1000000*Y[j]
        Model += K3[i][j] <= 1000000*Y[j]
sum1 = 0      
for j in f:
    Model += N[j] <= 1000000*Y[j]
    Model += N[j] <= 130
for j in f:
    sum1 += N[j]
Model += sum1 <= 269
for j in f:
    sum2 = 0
    for i in c:
        sum2 += X3[i][j] + A3[i][j] + B3[i][j]
    Model += sum2 <= 117.5*N[j]
    sum3 = 0
    sum4 = 0
    sum5 = 0
    for i in c:
        sum3 += K1[i][j]
        sum4 += K2[i][j]
        sum5 += K3[i][j]
    Model += sum3 == 1
    Model += sum4 == 1
    Model += sum5 == 1

# به دست آوردن جواب نهایی تابع هدف و متغیرها    
A = []
B = []

Model.solve()
print(pl.LpStatus[Model.status])
print(pl.value(Model.objective))
for v in Model.variables():
     A.append(v.name)
     B.append(v.varValue)
     df = pd.DataFrame({'col1':A , 'col2':B})
print(df)

o = [{'name':name,'shadow price':c.pi,'slack': c.slack} for name, c in Model.constraints.items()]
p = pd.DataFrame(o)
print(pd.DataFrame(o))
df.to_csv(r'C:/Users/pc/Downloads/Telegram Desktop/pandas.txt', header=None, index=None, sep=' ', mode='a')
p.to_csv(r'C:/Users/pc/Downloads/Telegram Desktop/pandas.txt', header=None, index=None, sep=' ', mode='a')
