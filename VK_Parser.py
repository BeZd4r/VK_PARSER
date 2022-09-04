from unicodedata import name
from openpyxl import load_workbook
import requests
import time

names,page_url = [],[]
offset = 0
total_count_post = 0

def Post_Checker(post_id):
    try:
        url_post = f"https://api.vk.com/method/wall.getComments?thread_items_count=10&owner_id=-212637974&extended=1&post_id={post_id}&count=100&access_token={token}&v=5.131"
        response = requests.get(url_post).json()

        for objects in response["response"]["profiles"]:
            name = objects["first_name"]+" "+objects["last_name"]
            name_id = "https://vk.com/id"+str(objects["id"])
            if name not in names:
                names.append(name)
                page_url.append(name_id)

        reply_id = ""
        for replies in response["response"]["items"]:
            for reply in replies["thread"]["items"]:
                if str(reply["from_id"]) not in reply_id:
                    reply_id += str(reply["from_id"])+", "
        else:
            time.sleep(0.1)
            url_users = f"https://api.vk.com/method/users.get?user_ids={reply_id}&access_token={token}&v=5.131"
            response = requests.get(url_users).json()
            for objectes in response["response"]:
                name = str(objectes["first_name"])+" "+str(objectes["last_name"])
                name_id = "https://vk.com/id"+str(objectes["id"])
                if name not in names:
                    names.append(name)
                    page_url.append(name_id)

        print(f"Post {post_id} is checked")
    except Exception:
        print(f"Post {post_id} is broke")

def Wall_Checker():
    global total_count_post, offset
    url_wall = f"https://api.vk.com/method/wall.get?offset={offset}&owner_id=-212637974&extended=1&count=100&access_token={token}&v=5.131"
    response = requests.get(url_wall).json()

    for item in response['response']['items']:
        Post_Checker(int(item['id']))
        time.sleep(0.25)

    if total_count_post > 0:
        offset += 100
        total_count_post -= 100
        Wall_Checker()

f = open("Server_Token_(VK_API).txt")
token = f.read()

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
