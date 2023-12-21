import openpyxl as op
import matplotlib.pyplot as plt

file_name = 'data.xlsx'

wb = op.load_workbook(file_name)

sheet = wb.active

max_row = sheet.max_row


#Запишем в словарь координаты ячеейк с началом месяца
month_cords = {}
year = '2021'
for i in sheet ['C1' :f'C{max_row}']:
    if year in i[0].value:
        month = i[0].value.split()
        month_cords[month[0]] = i[0].coordinate[1:]



#Вопрос 1
def sum_cash(up_border:str,down_border:str):

    Jule_sum = 0

    for cellobj in sheet[f'B{month_cords[up_border]}' : f'C{month_cords[down_border]}']:
        row = cellobj
        if cellobj[1].value == 'ОПЛАЧЕНО':
            Jule_sum+=cellobj[0].value

    print('Вопрос 1 ',Jule_sum)



#Вопрос 2

def all_cash(start_cell:str):

    months = [i for i in month_cords.keys()]

    last_cell = start_cell[0:1]+str(max_row)

    cash = []
    cash_at_month = 0
    for i in sheet[start_cell : last_cell]:
        if i[0].value  and i[0].coordinate not in month_cords.values() and int(i[0].coordinate[1:]) != max_row :
            cash_at_month+=i[0].value
        else:
            if int(i[0].coordinate[1:] )== max_row:
                cash.append(cash_at_month+i[0].value)
            else:
                cash.append(cash_at_month)
            cash_at_month = 0

    print('Вопрос 2 ',cash)

    cord_x = months
    cord_y = cash

    plt.plot(cord_x,cord_y)
    plt.grid()
    plt.show()



#Вопрос 3
def best_stuff(up_border:str,down_border:str):
    names = {}

    for name in sheet[f'B{month_cords[up_border]}' : f'D{month_cords[down_border]}']:
        row = name
        if row[2].value in names and row[1].value == 'ОПЛАЧЕНО':
            names[row[2].value]+=row[0].value
        elif row[2].value not in names and row[1].value == 'ОПЛАЧЕНО':
            names[row[2].value] = row[0].value

    max_name = max(names.values())
    final_name = {k:v for k, v in names.items() if v == max_name}

    print("Вопрос 3 ",final_name)

#Вопрос 4
def most_type(up_border : str,down_border : str or max_row):
    status_dict = {}
    
    for status in sheet[f'E{month_cords[up_border]}' : f'E{down_border}']:
        if status[0].value in status_dict:
            status_dict[status[0].value]+=1
        else:
            status_dict[status[0].value] = 1

    max_status = max(status_dict.values())
    final_status = {k:v for k,v in status_dict.items() if v == max_status}

    print("Вопрос 4 ",final_status)

#Вопрос 5
def OrgiginalsAtMonth(up_border:str,down_border:str,month_search:int):
    originals_at_month = 0
    
    for date in sheet[f'H{month_cords[up_border]}' : f'H{month_cords[down_border]}']:
        date_str = date[0].value
        if date_str and date_str.month == month_search :
            originals_at_month+=1
        else:
            continue

    print("Вопрос 5 ",originals_at_month)


#Задание

def prize_remains(up_border: str,down_border: str,search_month: int ):
    remains = 0
    
    for row in sheet[f"B{month_cords[up_border]}" : f"H{month_cords[down_border]}"]:
        if row[5].value == 'оригинал' and row[6].value.month>search_month:
            if row[3].value == "новая" and row[1].value == 'ОПЛАЧЕНО':
                remains += (row[0].value*7)/100
            elif row[3].value == 'текущая' and row[1].value != 'ПРОСРОЧЕНО':
                if row[0].value>10000:
                    remains+=(row[0].value*5)/100
                else:
                    remains+=(row[0].value*3)/100
    print('Задание ',remains)



def main():

    sum_cash("Июль","Август")

    best_stuff('Сентябрь','Октябрь')

    most_type('Октябрь',max_row)

    OrgiginalsAtMonth('Июнь','Июль',5)

    prize_remains('Май','Июль',6)

    all_cash('B3')

if __name__ == '__main__':
    main()
