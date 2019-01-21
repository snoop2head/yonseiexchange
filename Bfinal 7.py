#THIS ONE WORKS

import re
from urllib.request import urlopen
from bs4 import BeautifulSoup
from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1', cell_overwrite_ok=True)

#HTML tag removal function
def remove_tag(content):
    cleanr = re.compile('<.*?>')
    cleantext = re.sub(cleanr, '', content)
    return cleantext

doc = """
<option value="AF">AFGHANISTAN</option><option value="AR">ARGENTINA</option><option value="AU">AUSTRALIA</option><option value="AT">AUSTRIA</option><option value="BD">BANGLADESH</option><option value="BE">BELGIUM</option><option value="BR">BRAZIL</option><option value="BN">BRUNEI</option><option value="BG">BULGARIA</option><option value="KH">CAMBODIA</option><option value="CA">CANADA</option><option value="CL">CHILE</option><option value="CN">CHINA</option><option value="CR">COSTA RICA</option><option value="CZ">CZECH REPUBLIC</option><option value="DK">DENMARK</option><option value="EE">ESTONIA</option><option value="FI">FINLAND</option><option value="FR">FRANCE</option><option value="DE">GERMANY</option><option value="GH">GHANA</option><option value="HK">HONG KONG</option><option value="HU">HUNGARY</option><option value="IS">ICELAND</option><option value="IN">INDIA</option><option value="ID">INDONESIA</option><option value="IR">IRAN</option><option value="IE">IRELAND</option><option value="IL">ISRAEL</option><option value="IT">ITALY</option><option value="JP">JAPAN</option><option value="JO">JORDAN</option><option value="KZ">KAZAKHSTAN</option><option value="LV">LATVIA</option><option value="LT">LITHUANIA</option><option value="MY">MALAYSIA</option><option value="MT">MALTA</option><option value="MX">MEXICO</option><option value="MN">MONGOLIA</option><option value="NP">NEPAL</option><option value="NL">NETHERLANDS</option><option value="NZ">NEW ZEALAND</option><option value="NI">NICARAGUA</option><option value="NO">NORWAY</option><option value="PH">PHILIPPINES</option><option value="PL">POLAND</option><option value="PT">PORTUGAL</option><option value="PR">PUERTO RICO</option><option value="RU">RUSSIAN FEDERATION</option><option value="SG">SINGAPORE</option><option value="ZA">SOUTH AFRICA</option><option value="ES">SPAIN</option><option value="SE">SWEDEN</option><option value="CH">SWITZERLAND</option><option value="TW">TAIWAN</option><option value="TH">THAILAND</option><option value="TR">TURKEY</option><option value="GB">UNITED KINGDOM</option><option value="US">UNITED STATES</option><option value="UY">URUGUAY</option><option value="VN">VIETNAM</option><option value="ZW">ZIMBABWE</option>
"""

#Separating doc HTML
soup_concod=BeautifulSoup(doc,'lxml')
list_concod=soup_concod.findAll('option') 
#print(list_concod) 

#Parsing Country Name String List
soup_con=BeautifulSoup(doc,'lxml')
list_con= [str(x.text) for x in soup_con.find_all('option')] 
#print(list_con) 


x=0
numb=0
con_numb=0
while con_numb < len(list_con) :
    #extracting individual country codes
    item_concod=list_concod[con_numb]
    str_item_concod=str(item_concod)
    print(str_item_concod[15:17])

    #finding table data, which is university queries, from country page
    url_cn= "https://oia.yonsei.ac.kr/partner/expReport.asp?yn=Y&country_code="+str_item_concod[15:17]+"&univ="
    html_cn=urlopen(url_cn)
    soup = BeautifulSoup(html_cn, 'lxml')
    type(soup)
    list_cn = soup.find_all('td')

    index = 0 #or start at 3, pending on where you want to start
    while index < len(list_cn) :
        #making a query corresponding for universities in a country
        list_a = list_cn[index].find('a')
        univ_query_full = list_a['href']
        univ_query=univ_query_full[-15:]
        print(univ_query)
        #extracting university name and writing on the beginning of excel file
        str_url=str(list_a) #stringfy in order to use regex
        univ_name=remove_tag(str_url)
        print(univ_name)
        sheet1.write(numb, x, univ_name)
        index +=4
        page=1
        while page < 15:
            #crawling and parsing department items according to page
            url = "https://oia.yonsei.ac.kr/partner/expReport.asp?page=" + str(page)+"&cur_pack=0&ucode="+str(univ_query)
            html = urlopen(url)
            soup = BeautifulSoup(html, 'lxml')
            type(soup)
            list = soup.find_all('td')
            #Not crawling pages from universities that no students have ever went
            if not list:
                page=15
            else:
                postit = 1 #starting point of list
                #writing each department items on excel sheet
                while postit < len(list) :
                    x +=1
                    item=list[postit]
                    str_item=str(item) #stringfy in order to use regex
                    dep=remove_tag(str_item)
                    sheet1.write(numb, x, dep)
                    wb.save('Final.xls')
                    postit +=4
                page +=1
        x=0
        numb+=1
    con_numb+=1
