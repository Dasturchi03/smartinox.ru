import os
import requests
from bs4 import BeautifulSoup
from bs4.element import Tag, NavigableString
from openpyxl import Workbook
from aiogram import Bot, Dispatcher, executor, types


TOKEN = "6430551619:AAG4F4yKgHah4mnrHcNHCzMa3dMqicGlLOw"

bot = Bot(TOKEN)
dp = Dispatcher(bot)


@dp.message_handler(commands=['start'])
async def start_bot(message: types.Message):
    await bot.send_message(message.chat.id, 'Привествие 👋')
    await bot.send_message(message.chat.id, 'Используйте команду /file, чтобы получить файл')


@dp.message_handler(commands=['file'])
async def send_file(message: types.Message):
    mess = await message.reply('Сбор данных...')
    main()
    await mess.delete()
    await bot.send_document(message.chat.id, types.InputFile('files/smartinox_ru.xlsx'))


def main():
    URLS = get_main_urls()
    data = []
    head = True
    for i in URLS:
        r = requests.get(i)
        soup = BeautifulSoup(r.text, 'lxml')
        div = soup.find('div', class_='catalog__sections items')
        urls = []
        for tag in div:
            if isinstance(tag, Tag) and tag.name == 'a':
                urls.append(tag.attrs['href'])
        
        # print(urls)
        for i in urls:
            r1 = requests.get(MAIN_URL + i)
            sp = BeautifulSoup(r1.text, 'lxml')
            d = parse_data(sp, head)
            head = False
            if d == -1:
                break
            data += d
            pages_cnt = check_pages(sp)
            for j in range(pages_cnt-1):
                r2 = requests.get(MAIN_URL + i + f'/?PAGEN_1={j+2}')
                sp1 = BeautifulSoup(r2.text, 'lxml')
                d1 = parse_data(sp1, head=False)
                if d1 != -1:
                    data += d1
    write_xlsx(data)


def parse_data(sp: BeautifulSoup, head=True):
    t = sp.find('div', class_='breds')
    name = t.text.split('->')[-1].strip()
    print(name)
    try:
        table_tag = sp.find('div', class_='catalog__table').table
    except:
        return -1
    q, k = 0, 0
    dt = []
    headers = []
    for tr in table_tag.tbody.contents:
        if isinstance(tr, Tag):
            q += 1
            if q % 2:
                if k < 2:
                    headers = ['Название категории']
                    k += 1
                    for td in tr.contents:
                        if isinstance(td, Tag):
                            headers.append(td.text.strip())
                continue
            s = [name]
            # print('headers', headers)
            for td in tr.contents:
                if isinstance(td, Tag):
                    try:
                        val = td.a.text.strip()
                        s.append(val)
                    except:
                        pass
            s = s[:-1]
            s[-2] = s[-2].removesuffix('руб.')
            dt.append(s)
    headers += ['Цена, руб']
    if head:
        dt.insert(0, headers)
    # print(dt)
    return dt


def get_main_urls():
    r = requests.get(MAIN_URL)
    soup = BeautifulSoup(r.text, 'lxml')
    URLS = []
    div = soup.find('div', class_='catalog__sections')
    for a in div.contents:
        if isinstance(a, Tag) and a.name == 'a':
            URL = a.attrs['href']
            URLS.append(MAIN_URL + URL)
    return URLS


def check_pages(soup: BeautifulSoup):
    div = soup.find('div', class_='bx-pagination')
    lis_cnt = 0
    if div:
        ul = div.div.ul
        for li in ul.contents:
            if isinstance(li, Tag):
                lis_cnt += 1
        lis_cnt -= 2
    return lis_cnt


def check_tag_is_exist(tag: Tag, name):
    for t in tag.children:
        if isinstance(t, Tag):
            if t.name == name:
                return True
    return False


def write_xlsx(data: list):
    wk = Workbook()
    sheet = wk.active
    for i in data:
        # print(i)
        sheet.append(i)
    
    try:
        os.mkdir('files')
    except Exception as err:
        print(err)
    
    wk.save('files/smartinox_ru.xlsx')


if __name__ == '__main__':
    MAIN_URL = 'https://smartinox.ru'
    executor.start_polling(dp, skip_updates=True)

