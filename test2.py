employees = {1001: {'name': 'lzp', 'age': 18, 'wage': '10000'}, 1002: {'name': 'lzp', 'age': 25, 'wage': '100000'}}
employees_number = list(employees.keys())
info = list(employees.values())
for i in range(0, len(employees_number)):
    number = employees_number[i]
    for j in range(0, len(employees_number)):
        print(len(employees_number))
        name = info[j]['name']
        age = info[j]['age']
        wage = info[j]['wage']
        print("工号", number,"姓名：", name, "年龄", age, "工资", wage)
