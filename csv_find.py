import csv
import re
import analizze
name = '../../ЦНАП/result(АвтоматическиВосстановлено).csv'
with open(name) as f:
    reader = csv.reader(f, delimiter=';')
    data = list(reader)
    headers = data.pop(0)
    # print('Headers: ', headers)
    
def find_posl(text):   
    find_text = [] 
    t = text
    for i in range(len(data)):
        row = data[i]
        
        text = re.sub(r'КАТЕГОРІЙ', 'КАТЕГОРІЇ', text, flags=re.IGNORECASE )
        text = re.sub(r'ГРОШОВИХ', 'ГРОШОВОЇ', text, flags=re.IGNORECASE )
        text = re.sub(r'КОМПЕНСАЦІЙ', 'КОМПЕНСАЦІЇ', text, flags=re.IGNORECASE )
        text = re.sub(r'ння', 'нь', text, flags=re.IGNORECASE )
        text = re.sub(r'i|і|и|__', '', text, flags=re.IGNORECASE )
        row[2] = re.sub(r'КАТЕГОРІЙ', 'КАТЕГОРІЇ', row[2], flags=re.IGNORECASE )
        row[2] = re.sub(r'КОМПЕНСАЦІЙ', 'КОМПЕНСАЦІЇ', row[2], flags=re.IGNORECASE )
        row[2] = re.sub(r'КАТЕГОРІЙ', 'КАТЕГОРІЇ', row[2], flags=re.IGNORECASE )
        row[2] = re.sub(r'ння', 'нь', row[2], flags=re.IGNORECASE )
        row[2] = re.sub(r'i|і|и', '', row[2], flags=re.IGNORECASE )
        
        if 'Призначення тимчасової державної допомоги дітям, батьки яких ухиляються від сплати аліментів, не мають можливості утримувати дитину або місце проживання їх невідоме'.upper() in t.upper() and '00154' == row[1]:
            text = row[2]
        if 'видача посвідчень особам з інвалідністю та особам з інвалідністю з дитинства'.upper() in t.upper() and '00242' == row[1]:
            text = row[2]
        if 'ЗАБЕЗПЕЧЕННЯ направлення на комплексну реабілітацію (абілітацію) осіб з інвалідністю, дітей з інвалідністю, дітей віком до трьох років (включно)'.upper() in t.upper() and '01997' == row[1]:
            text = row[2]
        if 'ВИДАЧА НАПРАВЛЕННЯ НА ПРОХОДЖЕННЯ ОБЛАСНОЇ, ЦЕНТРАЛЬНОЇ МІСЬКОЇ У ММ. КИЄВІ ТА СЕВАСТОПОЛІ'.upper() in t.upper() and '00117' == row[1]:
            text = row[2]
        if 'ЗАБЕЗПЕЧЕННЯ НАПРАВЛЕННЯ ДІТЕЙ З ІНВАЛІДНІСТЮ ДО РЕАБІЛІТАЦІЙНОЇ УСТАНОВИ ДЛЯ НАДАННЯ РЕАБІЛІТАЦІЙНИХ ПОСЛУГ ЗА ПРОГРАМОЮ'.upper() in t.upper() and '01996' == row[1]:
            text = row[2]
        if 'Призначення грошової допомоги особі, яка проживає разом з особою з інвалідністю I або II групи внаслідок психічного розладу'.upper() in t.upper() and '00103' == row[1]:
            text = row[2]
            # print([text, row[2]], re.sub(r'„|”|ˮ|"|\'|’|`|,|;| у | в |ЩОДО|СТОСОВНО|на|для|ПОСВІДЧЕННЯ|ПОСВІДЧЕННЮ|ЧИ|АБО|/|\s{1,}', '', text, flags=re.IGNORECASE).upper() in re.sub(r'„|”|ˮ|"|\'|’|`|,|;| у | в |ЩОДО|СТОСОВНО|на|для|ПОСВІДЧЕННЯ|ПОСВІДЧЕННЮ|ЧИ|АБО|/|\s{1,}', '', row[2], flags=re.IGNORECASE).upper())
        if re.sub(r'«|»|„|”|ˮ|"|\'|’|`|,|;| у | в |ЩОДО|СТОСОВНО|на|для|ПОСВІДЧЕННЯ|ПОСВІДЧЕННЮ|ЧИ|АБО|/|\s{1,}', '', text, flags=re.IGNORECASE).upper() in re.sub(r'«|»|„|”|ˮ|"|\'|’|`|,|;| у | в |ЩОДО|СТОСОВНО|на|для|ПОСВІДЧЕННЯ|ПОСВІДЧЕННЮ|ЧИ|АБО|/|\s{1,}', '', row[2], flags=re.IGNORECASE).upper():
            find_text = row
            break
    if len(find_text) == 0:
        d = analizze.get_data_523()
        for i in range(len(d)):

            row = d[i]
            # if '00154' in row[1]:
            #     print([re.sub(r'„|”', '', text, flags=re.IGNORECASE), row[2]])
            text = re.sub(r'КАТЕГОРІЙ', 'КАТЕГОРІЇ', text, flags=re.IGNORECASE )
            text = re.sub(r'ГРОШОВИХ', 'ГРОШОВОЇ', text, flags=re.IGNORECASE )
            text = re.sub(r'КОМПЕНСАЦІЙ', 'КОМПЕНСАЦІЇ', text, flags=re.IGNORECASE )
            text = re.sub(r'ння', 'нь', text, flags=re.IGNORECASE )
            text = re.sub(r'i|і|и', '', text, flags=re.IGNORECASE )
            row[1] = re.sub(r'КАТЕГОРІЙ', 'КАТЕГОРІЇ', row[1], flags=re.IGNORECASE )
            row[1] = re.sub(r'ГРОШОВИХ', 'ГРОШОВОЇ', row[1], flags=re.IGNORECASE )
            row[1] = re.sub(r'КОМПЕНСАЦІЙ', 'КОМПЕНСАЦІЇ', row[1], flags=re.IGNORECASE )
            row[1] = re.sub(r'ння', 'нь', row[1], flags=re.IGNORECASE )
            row[1] = re.sub(r'i|і|и|__', '', row[1], flags=re.IGNORECASE )
            if 'ВЗЯТТЯ НА ОБЛІК ДЛЯ ЗАБЕЗПЕЧЕННЯ САНАТОРНО-КУРОРТНИМ ЛІКУВАННЯМ (ПУТІВКАМИ) ветеранів війни та осІб, на яких поширюється дія законів УКРАЇНИ'.upper() in t.upper() and '00131' == row[0]:
                text = row[1]
            if '„ВИДАЧА ПІКЛУВАЛЬНИКУ ДОЗВОЛУ НА НАДАННЯ ЗГОДИ ОСОБІ, ДІЄЗДАТНІСТЬ ЯКОЇ ОБМЕЖЕНА, НА ВЧИНЕННЯ ПРАВОЧИНІВ ЩОДО УКЛАДЕННЯ ДОГОВОРІВ, ЯКІ ПІДЛЯГАЮТЬ НОТАРІАЛЬНОМУ ПОСВІДЧЕННЮ ТА (АБО) ДЕРЖАВНІЙ РЕЄСТРАЦІЇ'.upper() in t.upper() and '00228' == row[0]:
                text = row[1]

            if re.sub(r'«|»|„|”|ˮ|"|\'|’|`|,|;| у | в |ЩОДО|СТОСОВНО|на|для|ПОСВІДЧЕННЯ|ПОСВІДЧЕННЮ|ЧИ|АБО|/|\s{1,}', '', text, flags=re.IGNORECASE).upper() in re.sub(r'«|»|„|”|ˮ|"|\'|’|`|,|;| у | в |ЩОДО|СТОСОВНО|на|для|ПОСВІДЧЕННЯ|ПОСВІДЧЕННЮ|ЧИ|АБО|/|\s{1,}', '', row[1], flags=re.IGNORECASE).upper():
                find_text = ['523p'] + row
                break
    return find_text

# with open(name) as f:
#     reader = csv.DictReader(f)
#     for row in reader:
#         print(row)

# with open('sw_data_new.csv', 'w') as f:
#     writer = csv.writer(f, quoting=csv.QUOTE_NONNUMERIC, lineterminator='\n')
#     for row in data:
#         writer.writerow(row)

# with open('csv_write_dictwriter.csv', 'w') as f:
#     writer = csv.DictWriter(
#         f, fieldnames=list(data[0].keys()), quoting=csv.QUOTE_NONNUMERIC)
#     writer.writeheader()
#     for d in data:
#         writer.writerow(d)