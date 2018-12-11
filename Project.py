import openpyxl
from collections import Counter
f = openpyxl.load_workbook(r"C:\Users\MR-X\Desktop\aum123.xlsx")
sheet = f.active

name = []
for r in sheet.rows:
    s = []
    for i in range(1, 6):
        s.append(r[i].value)
    name.append(s)

print('Popping things...')
name.pop(0)

# 각 배열에 A~G 입력
r_max_people = [32, 100, 11, 11, 15, 30, 8]
regions = [[], [], [], [], [], [], []]
temp = []
for i in name:
    regions['ABCDEFG'.index(i[3])].append(i)

# 성비 확인
def RatioCheck():
    ratio = []
    for i in range(7):
        s = list(map(list, zip(*regions[i])))[1]
        w = s.count('여자')
        m = s.count('남자')
        ratio.append((round(w * 100 / (w + m), 1), round(m * 100 / (w + m), 1)))
    return ratio

# 정원 확인
def PersonCheck():
    ppl = []
    for i in range(7):
        ppl.append(len(regions[i]) - r_max_people[i])
    return ppl

def GetManRatio():
    s = list(map(list, zip(*regions[i])))[1]
    m = s.count('남자')
    w = s.count('여자')
    manratio = round(m * 100 / (w + m), 1)
    return manratio

ratio = RatioCheck()
for i in range(7):
    print('Ratio on {} -> {} : {} ... {}'.format('ABCDEFG'[i], ratio[i][0], ratio[i][1], 'OK' if (ratio[i][1]>= 30 and ratio[i][1] <= 50) else 'Not OK'))

ppl = PersonCheck()
for i in range(7):
    print('People on {} -> {} / {} ... {} ({})'.format('ABCDEFG'[i], len(regions[i]), r_max_people[i], ppl[i] if ppl[i] == 'OK' else 'Not OK',ppl[i]))
print(regions[0])
ages = list(map(list, zip(*regions[0])))[2]
print(ages)

for _ in range(20): # While
    print("ATTEMPT {}".format(_))
    ppl = PersonCheck()
    isChanged = 0
    for i in range(7):
        g = 'ABCDEFG'[i]
        if ppl[i] == 0:
            print('OK with {}, continue...'.format(g))

        elif ppl[i] > 0:
            print('Too many people in {}, pushing people to temporary list...'.format(g))
            for j in range(ppl[i]):
                for a in regions[i]:
                    cnt = Counter(list(map(list, zip(*regions[i])))[2])
                    most_age = cnt.most_common(1)[0][0]
                    manratio = GetManRatio()
                    if manratio <= 40:
                        if a[1] == '여자' and a[2] == most_age:
                            print("Pushing {} into temp...".format(a[0]))
                            temp.append(a)
                            regions[i].pop(regions[i].index(a))
                            isChanged = 1
                    elif manratio > 40:
                        if a[1] == '남자' and a[2] == most_age:
                            print("Pushing {} into temp...".format(a[0]))
                            temp.append(a)
                            regions[i].pop(regions[i].index(a))
                            isChanged = 1

        # 로직상 성비를 맞추기 위해 거르기 + 나이 거르기 동시에 걸러지는 게 없을 때 옮기지 못한다.

        else:
            print('Low people in {}, pulling people from temporary list...'.format(g))
            for a in temp:
                manratio = GetManRatio()
                if manratio <= 40:
                    if a[1] == '남자' and a[4] == 'ABCDEFG'[i]:
                        regions[i].append(temp.pop(temp.index(a)))
                else:
                    if a[1] == '여자' and a[4] == 'ABCDEFG'[i]:
                        regions[i].append(temp.pop(temp.index(a)))
            if not temp:
                print('No people in temporary list, continue...'.format(g))
            else:
                isChanged = 1
    if not isChanged:
        break

for i in range(7):
    print(len(regions[i]),end=', ')
print()
print(regions[2])
print(RatioCheck())

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Sheet'

for i in range(7):
    for r in range(len(regions[i])):
        for k in range(5):
            ws.cell(row=r+1, column=i*5+k+1).value = regions[i][r][k]
for r in range(len(temp)):
    for k in range(5):
        ws.cell(row=r+1, column=36+k).value = temp[r][k]
wb.save(r'test.xlsx')
