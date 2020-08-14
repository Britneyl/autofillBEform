from selenium import webdriver
import time
import xlwings as xw
from selenium.webdriver.common.keys import Keys
def sign_up():
    get_tab = input('1.资产负债表\n2.利润表\n:')
    file_path = input('请输入路径:')  
    driver = webdriver.Chrome()
    driver.set_page_load_timeout(10)
    driver.get('https://etax.*******.chinatax.gov.cn/zjgfdzswj/main/index.html')#通过谷歌浏览器的驱动获取网址并打开
    driver.find_element_by_id('wybs').click()
    time.sleep(65)
    driver.get('https://etax.zhejiang.*******.gov.cn/zjgfzjdzswjsbweb/pages/sb/nssb/sb_nssb.html')
    driver.find_element_by_link_text('财务报表（小企业会计准则）').click()
    wb = xw.Book(file_path)#文件的路径获取
    sht = wb.sheets['Recovered_Sheet1']#获取要控制表单的名字
    if(get_tab == '1'):
        driver.find_element_by_link_text('资产负债表').click()
        driver_search_send("xx_001_qms","E8",driver,sht)#货币资金
        driver_search_send("xx_033_qms","I10",driver,sht)#应付账款
        driver_search_send("xx_004_qms","E13",driver,sht)#应收账款净额
    elif(get_tab == '2'):
        driver.find_element_by_link_text('利润表').click()
        driver_search_send("xx_001_sns","D7",driver,sht)#产品销售收入
        driver_search_send("xx_002_sns","D8",driver,sht)#产品销售成本
        driver_search_send("xx_011_sns","D9",driver,sht)#产品销售费用
        driver_search_send("xx_003_sns","D10",driver,sht)#产品销售税金及附加
        #driver_search_send("","D11",driver,sht)#产品销售利润
        driver_search_send("","D12",driver,sht)#其他业务利润
        driver_search_send("xx_014_sns","D13",driver,sht)#管理费用
        driver_search_send("xx_018_sns","D14",driver,sht)#财务费用
        driver_search_send("","D15",driver,sht)#利息支出
        driver_search_send("","D16",driver,sht)#汇兑损益
        #driver_search_send("","D17",driver,sht)#营业利润
        driver_search_send("","D18",driver,sht)#投资收益
        driver_search_send("xx_022_sns","D19",driver,sht)#营业外收入
        driver_search_send("","D20",driver,sht)#营业外支出
        driver_search_send("","D21",driver,sht)#以前年度损益调整
        #driver_search_send("","D22",driver,sht)#利润总额
        driver_search_send("xx_031_sns","D23",driver,sht)#所得税
        if sht.range("E24").value.replace(',','') == driver.find_element_by_id("xx_032_bns").text:
            driver.find_element_by_id("").click()
        else:
            print("请校对一下哪里有错")
            time.sleep(120)

    wb.close()
    driver.close()

def driver_search_send(htm_id,form,driver,sht):
    if htm_id is None:
        print(htm_id + "是空值，填写下一个")
        pass
    else:
        driver.find_element_by_id(htm_id).send_keys(Keys.BACKSPACE*4+sht.range(form).value.replace(',',''))
        #通过驱动的调用，四个退格键删除格子内容，然后通过找到excel里面的东西写到网页

if __name__ == "__main__":
    sign_up()
