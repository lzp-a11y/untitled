from test1 import operation_excel
import time

class test(operation_excel):
    def loading(self):
        while True:
            usrename = 'admin'
            password = 123456
            my_usrename = input("请输入用户名：")
            my_password = int(input("请输入密码："))
            if usrename == my_usrename and password == my_password:
                print("登录成功！")
                return True
            else:
                print("用户名活密码有误！")
                b = int(input("重新输入请按0，退出系统请按除0任意数字键："))
                if b != 0:
                    break

    def operation(self):
        if_loading = self.loading()
        while if_loading:
            print("*" * 20, "员工管理系统", "*"*20)
            print("输入1：展示所有的员工信息")
            print("输入2：新增一个员工信息")
            print("输入3：修改一个员工信息")
            print("输入4：删除一个员工信息")
            print("输入5：退出员工管理系统")
            print("*" * 50)
            a = int(input("请输入您的操作："))
            if a == 1:
                self.query()
            elif a == 2:
                self.new_inced()
            elif a == 3:
                self.amend()
            elif a == 4:
                self.delete_data()
            else:
                print("退出登录成功！")
                break

test().operation()