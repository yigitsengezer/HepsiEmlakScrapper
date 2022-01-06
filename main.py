import requests
from bs4 import BeautifulSoup, BeautifulStoneSoup
from fake_useragent import UserAgent
from openpyxl import load_workbook, workbook
from requests.api import request
from time import sleep

print("""Şehir kodları:
0-Adana
1-Antalya
2-Aydın
3-Denizli
4-Hatay
5-Isparta
6-Muğla
7-Bütün İller
""")

location_code = input("Arama yapmak istediğiniz şehir veya şehirlerin kodlarını girin: ")
location_list = ["adana", "antalya", "aydin", "denizli", "hatay", "isparta", "mugla"]
selected_locations = list(map(int, location_code.split()))

if(selected_locations[0]==7):
    selected_locations = location_list[0:11]
else:
    for i in range(len(selected_locations)):
        selected_locations[i] = location_list[selected_locations[i]]

print("Seçtiğiniz Şehirler =>", selected_locations)
workbook = load_workbook(filename="Template.xlsx")
sheet = workbook.active

def beautify(string):
    string=" ".join(string.replace("\n","").split())
    return string

for selected_location in selected_locations:
    page_index = 1
    while(True):
        user_agent = {'User-Agent':str(UserAgent().random)}
        url = f"https://www.hepsiemlak.com/{selected_location}-satilik?page={page_index}"
        page = requests.get(url,headers=user_agent)
        while(page.status_code !=200):
            print("Çok fazla istek atildi")
            sleep(60*5)
            user_agent = {'User-Agent':str(UserAgent().random)}
            url = f"https://www.hepsiemlak.com/{selected_location}-satilik?page={page_index}"
            page = requests.get(url,headers=user_agent)
            

        soup = BeautifulSoup(page.content,"html.parser")
        home_links = soup.find_all("a",class_="card-link")
        if(home_links == []):
            break

        for home_link in home_links:
            home_link = home_link['href']
            home_url = f"https://www.hepsiemlak.com{home_link}"
            home_page = requests.get(home_url,headers=user_agent)
            while(home_page.status_code !=200):
                print("Çok fazla istek atildi")
                sleep(60*5)
                user_agent = {'User-Agent':str(UserAgent().random)}
                home_url = f"https://www.hepsiemlak.com{home_link}"
                home_page = requests.get(home_url,headers=user_agent)
            
            home_soup = BeautifulSoup(home_page.content,"html.parser")

            home_title = beautify(home_soup.find("h1",class_="fontRB").text)
            home_price = beautify(home_soup.find("p",class_="fontRB fz24 price").text)
            home_properties = (home_soup.find("ul",class_="short-info-list")).find_all("li")
            home_info = home_soup.findAll("li",class_="spec-item")

            home_city = beautify(home_properties[0].text)
            home_town = beautify(home_properties[1].text)
            home_district = beautify(home_properties[2].text)
            home_type = beautify(home_properties[3].text+" "+home_properties[4].text)
            home_room = beautify(home_properties[5].text)
            home_size = beautify(home_properties[6].text)
            home_no = ""
            home_date = ""
            home_floor = ""
            home_age = ""
            home_heating = ""
            home_top_floor = ""
            home_credit = ""
            home_furniture = ""
            home_bathrooms = ""
            home_building = ""
            home_state = ""
            home_status = ""
            home_deed = ""
            home_dues = ""
            home_trade = ""
            home_front = ""
            home_rent = ""
            home_fuel = ""
            home_office = ""
            home_site = ""
            home_council = ""
            home_island = ""
            home_parcel = ""

            row_index = 2
            while True:
                if(sheet[f'A{row_index}'].value!=None):
                    row_index += 1
                else:
                    break
            sheet[f'A{row_index}'].value = home_title
            sheet[f'B{row_index}'].value = home_price
            sheet[f'C{row_index}'].value = home_city
            sheet[f'D{row_index}'].value = home_town
            sheet[f'E{row_index}'].value = home_district
            sheet[f'F{row_index}'].value = home_type
            sheet[f'G{row_index}'].value = home_room
            sheet[f'H{row_index}'].value = home_size

            print("Başlık:", home_title)
            print("Ücret:", home_price)
            print("Şehir:", home_city)
            print("İlçe:", home_town)
            print("Mahalle:", home_district)
            print("Konut Tipi:", home_type)
            print("Oda Sayısı:", home_room)
            print("Büyüklük:", home_size)

            for i in home_info: 
                if "İlan no" in i.text :
                    home_no = (i.text)[8:]
                    sheet[f'I{row_index}'].value = home_no
                    print("İlan no:", home_no)
                elif "Son Güncelleme Tarihi" in i.text:
                    home_date = (i.text)[22:]
                    sheet[f'J{row_index}'].value = home_date
                    print("Son Güncelleme Tarihi:", home_date)
                elif "Bulunduğu Kat" in i.text:
                    home_floor = (i.text)[14:]
                    sheet[f'K{row_index}'].value = home_floor
                    print("Bulunduğu Kat:", home_floor)
                elif "Bina Yaşı" in i.text:
                    home_age = (i.text)[10:]
                    sheet[f'L{row_index}'].value = home_age
                    print("Bina Yaşı:", home_age)
                elif "Isınma Tipi" in i.text:
                    home_heating = (i.text)[12:]
                    sheet[f'M{row_index}'].value = home_heating
                    print("Isınma Tipi:", home_heating)
                elif "Kat Sayısı" in i.text:
                    home_top_floor = (i.text)[11:]
                    sheet[f'N{row_index}'].value = home_top_floor
                    print("Kat Sayısı:", home_top_floor)
                elif "Krediye Uygunluk" in i.text:
                    home_credit = (i.text)[17:]
                    sheet[f'O{row_index}'].value = home_credit
                    print("Krediye Uygunluk:", home_credit)
                elif "Eşya Durumu" in i.text:
                    home_furniture = (i.text)[12:]
                    sheet[f'P{row_index}'].value = home_furniture
                    print("Eşya Durumu:", home_furniture)
                elif "Banyo Sayısı" in i.text:
                    home_bathrooms = (i.text)[13:]
                    sheet[f'Q{row_index}'].value = home_bathrooms
                    print("Banyo Sayısı:", home_bathrooms)
                elif "Yapı Tipi" in i.text:
                    home_building = (i.text)[10:]
                    sheet[f'R{row_index}'].value = home_building
                    print("Yapı Tipi:", home_building)
                elif "Yapının Durumu" in i.text:
                    home_state = (i.text)[15:]
                    sheet[f'S{row_index}'].value = home_state
                    print("Yapının Durumu:", home_state)
                elif "Kullanım Durumu" in i.text:
                    home_status = (i.text)[16:]
                    sheet[f'T{row_index}'].value = home_status
                    print("Kullanım Durumu:", home_status)
                elif "Tapu Durumu" in i.text:
                    home_deed = (i.text)[12:]
                    sheet[f'U{row_index}'].value = home_deed
                    print("Tapu Durumu:", home_deed)
                elif "Takas" in i.text:
                    home_trade = (i.text)[6:]
                    sheet[f'V{row_index}'].value = home_trade
                    print("Takas:", home_trade)
                elif "Cephe" in i.text:
                    home_front = (i.text)[6:]
                    sheet[f'W{row_index}'].value = home_front
                    print("Cephe:", home_front)
                elif "Site İçerisinde" in i.text:
                    home_site = (i.text)[16:]
                    sheet[f'X{row_index}'].value = home_site
                    print("Site İçerisinde:", home_site)
                elif "Yakıt Tipi" in i.text:
                    home_fuel = (i.text)[11:]
                    sheet[f'Y{row_index}'].value = home_fuel
                    print("Yakıt Tipi:", home_fuel)
                elif "Yetkili Ofis" in i.text:
                    home_office = (i.text)[13:]
                    sheet[f'Z{row_index}'].value = home_office
                    print("Yetkili Ofis:", home_office)
                elif "Aidat" in i.text:
                    home_dues = (i.text)[6:]
                    sheet[f'AA{row_index}'].value = home_dues
                    print("Aidat:", home_dues)
                elif "Kira Getirisi" in i.text:
                    home_rent = (i.text)[14:]
                    sheet[f'AB{row_index}'].value = home_rent
                    print("Kira Getirisi:", home_rent)
                elif "Ada" in i.text:
                    home_island = (i.text)[3:]
                    sheet[f'AC{row_index}'].value = home_island
                    print("Ada:", home_island)
                elif "Parsel" in i.text:
                    home_parcel = (i.text)[6:]
                    sheet[f'AD{row_index}'].value = home_parcel
                    print("Parsel:", home_parcel)
                elif "İlgili Belediye" in i.text:
                    home_council = (i.text)[15:]
                    sheet[f'AE{row_index}'].value = home_council
                    print("İlgili Belediye:", home_council)
                
                
            print("------------------------------------------------")
        page_index +=1
        workbook.save("Sonuc.xlsx")
input("Bitti")


        
