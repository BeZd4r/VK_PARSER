from openpyxl import load_workbook
import requests
import time

names,page_ids ,page_url = [],[], []
offset = 0
total_count_post = 0

def Post_Checker(post_id):
    try:
        url_post = f"https://api.vk.com/method/wall.getComments?thread_items_count=10&owner_id=-212637974&extended=1&post_id={post_id}&count=100&access_token={token}&v=5.131"
        response = requests.get(url_post).json()

        for objects in response["response"]["profiles"]:
            name = objects["first_name"]+" "+objects["last_name"]
            name_id = "https://vk.com/id"+str(objects["id"])
            page_id = str(objects["id"])
            if name not in names:
                names.append(name)
                page_ids.append(page_id)
                page_url.append(name_id)

        for objects in response["response"]["groups"]:
            name = objects["name"]
            name_id = "https://vk.com/id"+str(objects["id"])
            page_id = str(objects["id"])
            if name not in names:
                names.append(name)
                page_ids.append(page_id)
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

url_wall = f"https://api.vk.com/method/wall.get?offset=0&owner_id=-212637974&extended=1&count=100&access_token={token}&v=5.131"
response = requests.get(url_wall).json()
total_count_post = response["response"]["count"]

Wall_Checker()

exel_file = "Table.xlsx"
work_book = load_workbook(exel_file)
work_book.remove(work_book["Data"])
work_list = work_book.create_sheet("Data")
work_list.column_dimensions['A'].width = 10
work_list.column_dimensions['B'].width = 24
work_list.column_dimensions['C'].width = 10
work_list.column_dimensions['D'].width = 30
for i in range(1,len(names)+1):
    work_list[f"A{i}"] = i
    work_list[f"B{i}"] = names[i-1]
    work_list[f"C{i}"] = page_ids[i-1]
    work_list[f"D{i}"] = page_url[i-1]
work_book.save(exel_file)
work_book.close()
