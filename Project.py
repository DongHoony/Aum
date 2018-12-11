import openpyxl
from collections import Counter
f = openpyxl.load_workbook(r"C:\Users\So young\Desktop\aum123.xlsx")
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
        m = w = 0
        for j in range(len(regions[i])):
            if regions[i][j][1] == '여자':
                w += 1
            else:
                m += 1
        ratio.append((round(w * 100 / (w + m), 1), round(m * 100 / (w + m), 1)))
    return ratio

# 정원 확인
def PersonCheck():
    ppl = []
    for i in range(7):
        ppl.append(len(regions[i]) - r_max_people[i])
    return ppl

ratio = RatioCheck()
for i in range(7):
    print('Ratio on {} -> {} : {} ... {}'.format('ABCDEFG'[i], ratio[i][0], ratio[i][1], 'OK' if (ratio[i][1]>= 30 and ratio[i][1] <= 50) else 'Not OK'))

ppl = PersonCheck()
for i in range(7):
    print('People on {} -> {} / {} ... {} ({})'.format('ABCDEFG'[i], len(regions[i]), r_max_people[i], ppl[i] if ppl[i] == 'OK' else 'Not OK',ppl[i]))
print(regions[0])
ages = list(map(list, zip(*regions[0])))[2]
print(ages)

for _ in range(10): # While
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
                cnt = Counter(list(map(list, zip(*regions[i])))[2])
                most_age = cnt.most_common(1)[0][0]
                print('MODE AGE : {}'.format(most_age))
                for a in regions[i]:
                    print(a)
                    if a[1] == '여자' and a[2] == most_age:
                        print("Pushing {} into temp...".format(a[0]))

                        temp.append(a)
                        regions[i].pop(regions[i].index(a))
                    else:
                        continue

            isChanged = 1

        else:
            print('Low people in {}, pulling people from temporary list...'.format(g))
            if not temp:
                print('No people in temporary list, continue...'.format(g))
            else:
                isChanged = 1
    if not isChanged:
        break

for i in range(7):
    print(regions[i])
print(RatioCheck())
