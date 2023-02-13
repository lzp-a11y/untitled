import unittest
import HTMLTestRunner
import HTMLTestRunnerCN
import os
# dri = './case'
root_dir = os.path.dirname(os.path.abspath(__file__))
# path.join路径拼接， path.join("home","code")  = home\code
dir = os.path.join(root_dir, "case")
suite = unittest.defaultTestLoader.discover(start_dir=dir, pattern='unit*.py')

if __name__ == '__main__':
    runner = HTMLTestRunnerCN.HTMLTestRunnerCN(open("./result.html", 'wb'), title='测试报告',
                                               description="说明")
    runner.run(suite)
