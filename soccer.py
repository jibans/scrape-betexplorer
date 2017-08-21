from selenium import webdriver
import unicodedata
import time
import xlwt


def export_data(data_list,pagedata):

    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
    style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')
    row=1
    for data in data_list:
        print(data)
        if data !="":
            cell=0
            ws.write(row, cell, pagedata[1])
            cell +=1
            ws.write(row, cell, pagedata[2])
            cell +=1
            ws.write(row, cell, pagedata[3])
            cell +=1
            for row1 in data:
                print(row1)
                ws.write(row, cell, row1)
                cell +=1
            row +=1
            print(row)
    wb.save('soccer.xls')


def extractTab(tabLink,pagedata):
    driver = webdriver.Chrome("C:/chromedriver_win32/chromedriver.exe")
    driver.get(tabLink)
    tab_data = []
    for tr in driver.find_elements_by_xpath('//table[@class="table-main sortable"]//tr'):
        tds = tr.find_elements_by_tag_name('td')
        if tds:
            tab_data.append([unicodedata.normalize('NFKD', td.text).encode('ascii','ignore') for td in tds])
    print(tab_data)
    export_data(tab_data,pagedata)
    driver.quit()
    time.sleep(20)

def extractResultPage(data):
    str=data[0]
    print(str)
    tab1= str+"#1x2"
    extractTab(tab1,data)
    tab2= str+"#ou"
    extractTab(tab2,data)
    tab3= str+"#ah"
    extractTab(tab3,data)
    tab4= str+"#ha"
    extractTab(tab4,data)
    tab5 = str+"#dc"
    extractTab(tab5,data)
    tab6 = str+"#bts"
    extractTab(tab6,data)


def mainPageExtract() :

    driver = webdriver.Chrome("C:/chromedriver_win32/chromedriver.exe")
    driver.get("http://www.betexplorer.com/soccer/austria/tipico-bundesliga/results/")
    print(driver.title)
    all_data = []
    for tr in driver.find_elements_by_xpath('//table[@class="table-main js-tablebanner-t js-tablebanner-ntb"]//tr'):
        data = []
        tds = tr.find_elements_by_tag_name('td')
        if tds:
            i=0
            for td in tds :

                anchors = td.find_elements_by_tag_name('a')
                for a in anchors :
                    if i <1 :
                        link= driver.find_element_by_link_text(a.text)
                        data.append(unicodedata.normalize('NFKD', link.get_attribute('href')).encode('ascii','ignore'))
                        i=i+1
                        str=unicodedata.normalize('NFKD', td.text).encode('ascii','ignore')
                        data.append(str)
                        print(td.text)
                data.append(td.text)
            all_data.append(data)

    print(len(all_data))
    driver.quit()
    time.sleep(10)
    print("Extract page")
    for data in all_data :
        print(data)
        extractResultPage(data)


if __name__=="__main__":
    print("Main method")
    mainPageExtract()


