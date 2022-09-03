from unicodedata import name
from openpyxl import load_workbook
import requests
import time

names,page_url = [],[]
counter,offset = 0,0
url = f"https://api.vk.com/method/wall.get?owner_id=-212637974&extended=1&count=100&access_token=vk1.a.EVGM5hYS6zzsfFYj5IuvZairWa2hM7oFfMWcWea3bRJZi7bcBS048d2xQUFSdqvcS6aLXb4tHG4XT20fCxXQ_1KKkCf-lSu676xGuKZc81zFIWrTPunpfp_GLGkqCejV9PH7gEMT2wFNd6ubpvlVAzbmVj0T62NftYJMPTWBoImLd-F15-vk597Vrj6zgtue&v=5.131"
response = requests.get(url).json()
total_count_post = response["response"]["count"]

def Post_Checker(post_id):
    url_post = f"https://api.vk.com/method/wall.getComments?owner_id=-212637974&extended=1&post_id={post_id}&count=100&access_token=vk1.a.EVGM5hYS6zzsfFYj5IuvZairWa2hM7oFfMWcWea3bRJZi7bcBS048d2xQUFSdqvcS6aLXb4tHG4XT20fCxXQ_1KKkCf-lSu676xGuKZc81zFIWrTPunpfp_GLGkqCejV9PH7gEMT2wFNd6ubpvlVAzbmVj0T62NftYJMPTWBoImLd-F15-vk597Vrj6zgtue&v=5.131"
    response = requests.get(url_post).json()

    for objects in response["response"]["profiles"]:
        name = objects["first_name"]+" "+objects["last_name"]
        name_id = "https://vk.com/id"+str(objects["id"])
        if name not in names:
            names.append(name)
            page_url.append(name_id)


def Wall_Checker():
    global total_count_post, offset, counter
    url_wall = f"https://api.vk.com/method/wall.get?offset={offset}&owner_id=-212637974&extended=1&count=100&access_token=vk1.a.EVGM5hYS6zzsfFYj5IuvZairWa2hM7oFfMWcWea3bRJZi7bcBS048d2xQUFSdqvcS6aLXb4tHG4XT20fCxXQ_1KKkCf-lSu676xGuKZc81zFIWrTPunpfp_GLGkqCejV9PH7gEMT2wFNd6ubpvlVAzbmVj0T62NftYJMPTWBoImLd-F15-vk597Vrj6zgtue&v=5.131"
    response = requests.get(url_wall).json()

    for item in response['response']['items']:
        print (item['id'])
        Post_Checker(int(item['id']))
        time.sleep(0.25)
        counter += 1
        print(counter)

    if total_count_post > 0:
        offset += 100
        total_count_post -= 100
        Wall_Checker()

Wall_Checker()

print(len(names))

exel_file = "Table.xlsx"
work_book = load_workbook(exel_file)
work_book.remove(work_book["Data"])
work_list = work_book.create_sheet("Data")
for i in range(1,len(names)+1):
    work_list[f"A{i}"] = names[i-1]
    work_list[f"B{i}"] = page_url[i-1]
work_book.save(exel_file)
work_book.close()
