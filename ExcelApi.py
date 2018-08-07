import xlrd
import unittest
import requests
import time
import os
import HTMLTestRunner

class MyTest_Case(unittest.TestCase):

    def setUp(self):
        print("测试开始")

    def tearDown(self):
        print("测试结束")

    def testAlpha(self, filename='test.xlsx'):  # 读取 Excel文件
        data = xlrd.open_workbook(filename)  # 打开文件把参数传给 data
        table = data.sheets()[0]  # 通过索引顺序获取Excel 文件
        cookies = None
        for i in range(1,table.nrows):      # table.nrows 为Excel行数
            parameter = table.row_values(i)   # 获取整行数据 获取Excel 第二行数据
            data = eval(parameter[3])  # 因为 parameter 变量类型是str 需要用eval函数转换成dict
            url = 'http://192.168.8.19' + parameter[2]  # 基础地址 加上测试接口测路径
            statusCode = parameter[4]
            duanyan = parameter[6]
            headers = parameter[7]

            if cookies:
                r= requests.request("%s" % parameter[1],url,json=data,cookies=cookies)

            else:
                r= requests.request("%s" % parameter[1],url,json=data)
            # r.encoding='unicode-escape'
            r.encoding="utf-8"
            cookies = r.cookies
            # return url, parameter[i], jsoninfo,duanyan,statusCode
            self.assertEqual(r.status_code, statusCode)  # 断言code 是否等于200
            self.assertIn(duanyan,r.text)
            print("接口",parameter[0],"的测试返回值：",r.text)
            print("*************************************************************************************************")
if __name__ == "__main__":
    report_title = u'接口测试报告Beta_1'

    # 定义脚本内容，加u为了防止中文乱码
    desc = u'接口测试报告详情：'

    # 定义date为日期，time为时间
    date = time.strftime("%Y%m%d")
    time = time.strftime("%Y%m%d%H%M%S")
    testsuite = unittest.TestSuite()
    testsuite.addTest(MyTest_Case("testAlpha"))

    # filename = 'F:\\temp.html'
    with open('TestResult.html', 'wb') as fp:
        runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title=report_title, description=desc)
        runner.run(testsuite)
        fp.close()