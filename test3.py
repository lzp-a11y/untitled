employees = {}
def show_menu():
    print("*"*20, "员工管理系统", "*"*20)
    print("1.添加员工信息")
    print("2.删除员工信息")
    print("3.修改员工信息")
    print("4.显示员工信息")
    print("5.退出系统")
    print("*"*52)


def add_employees():
    employees_id = int(input("请输入员工编号"))
    all_employees = list(employees.keys())
    if employees_id in all_employees:
        print("员工编号重复，添加失败")
        return
    employees_name = input("请输入员工姓名：")
    employees_age = int(input("请输入员工年龄："))
    employees_wage = int(input("请输入员工工资："))
    employees_info = {"name":employees_name, "age":employees_age, "wage":employees_wage}
    employees[employees_id] = employees_info
    print(employees)


def query_info():
    employees_number = list(employees.keys())
    info = list(employees.values())
    for i in range(0, len(employees_number)):
        number = employees_number[i]
        for j in range(0, len(employees_number)):
            name = info[j]['name']
            age = info[j]['age']
            wage = info[j]['wage']
            print("工号", number,"姓名：", name, "年龄", age, "工资", wage)


while True:
    show_menu()
    a = int(input("请输入您的操作："))
    if a == 1:
        add_employees()
    elif a == 2:
        pass
    elif a == 3:
        pass
    elif a == 4:
        query_info()
    elif a == 5:
        print("退出成功!")
        break
    else:
        print("您输入的操作有误，请重新输入")