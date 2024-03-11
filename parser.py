import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import aiohttp
import asyncio
from bs4 import BeautifulSoup
import json
import bleach
import openpyxl
from threading import Thread

async def fetch(session, url):
    try:
        async with session.get(url) as response:
            return await response.text()
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return None

def sync_check_links():
    url = url_entry.get()
    link_pattern = link_pattern_entry.get()

    async def check_links():
        try:
            async with aiohttp.ClientSession() as session:
                html = await fetch(session, url)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    links = [a['href'] for a in soup.select(link_pattern) if 'href' in a.attrs]
                    unique_links = list(set(links))
                    status_label.config(text=f"Found links: {len(unique_links)}")
                else:
                    status_label.config(text="Failed to fetch page")
        except Exception as e:
            status_label.config(text=f"Error: {e}")
    
    # Run the asynchronous function in a separate thread
    Thread(target=lambda: asyncio.run(check_links())).start()

async def parse_product_page(session, url, title_selector, image_selector, description_selector, specs_selector):
    html = await fetch(session, url)
    if html is None:
        return None  # або поверніть значення за замовчуванням

    soup = BeautifulSoup(html, 'html.parser')
    title = 'Unavailable'
    image = 'Unavailable'
    description = 'Unavailable'
    specs = 'Unavailable'

    if title_selector:
        title_element = soup.select_one(title_selector)
        if title_element:
            title = title_element.get_text(strip=True)

    if image_selector:  # Додано обробку зображень
        image_elements = soup.select(image_selector + ' img')
        images = [img['src'] for img in image_elements if 'src' in img.attrs]
        if images:
            image = ', '.join(images)

    if description_selector:
        description_element = soup.select_one(description_selector)
        if description_element:
            description = clean_html(description_element.prettify())  

    if specs_selector:
        specs_element = soup.select_one(specs_selector)
        if specs_element:
            specs = clean_html(specs_element.prettify())

    return {
        'title': title,
        'image': image,
        'description': description,
        'specs': specs
    }





async def start_parsing_async():
    url = url_entry.get()
    link_pattern = link_pattern_entry.get()
    title_selector = product_title_entry.get()
    image_selector = product_image_entry.get()
    description_selector = product_description_entry.get()
    specs_selector = product_specs_entry.get()

    

    async with aiohttp.ClientSession() as session:
        html = await fetch(session, url)
        if html:
            soup = BeautifulSoup(html, 'html.parser')
            links = [a['href'] for a in soup.select(link_pattern) if 'href' in a.attrs]
            unique_links = list(set(links))

            total_links = len(unique_links)
            parsed_count = 0

            # Створити новий файл Excel та аркуш
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(['title', 'img', 'body', 'char'])

            for link in unique_links:
                product = await parse_product_page(session, link, title_selector, image_selector, description_selector, specs_selector)
                if product:
                    parsed_count += 1
                    status_label.config(text=f"Parsing... {parsed_count}/{total_links}")

                    # Додати дані в аркуш Excel
                    sheet.append([product['title'], product['image'], product['description'], product['specs']])

            # Зберегти Excel файл
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if filename:
                workbook.save(filename)

            status_label.config(text="Parsing complete")
        else:
            status_label.config(text="Failed to fetch initial page")


def save_profile():
    profile_data = {
        'url': url_entry.get(),
        'link_pattern': link_pattern_entry.get(),
        'product_title_selector': product_title_entry.get(),
        'product_image_selector': product_image_entry.get(),
        'product_description_selector': product_description_entry.get(),
        'product_specs_selector': product_specs_entry.get(),
    }
    filename = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('JSON files', '*.json')], title="Save Profile")
    if filename:
        with open(filename, 'w') as file:
            json.dump(profile_data, file)

def load_profile():
    filename = filedialog.askopenfilename(filetypes=[('JSON files', '*.json')], title="Load Profile")
    if filename:
        with open(filename, 'r') as file:
            profile_data = json.load(file)
            url_entry.insert(tk.END, profile_data['url'])
            link_pattern_entry.insert(tk.END, profile_data['link_pattern'])
            product_title_entry.insert(tk.END, profile_data['product_title_selector'])
            product_image_entry.insert(tk.END, profile_data['product_image_selector'])
            product_description_entry.insert(tk.END, profile_data['product_description_selector'])
            product_specs_entry.insert(tk.END, profile_data['product_specs_selector'])

# Wrapper function to run start_parsing_async in a separate thread
def start_parsing():
    Thread(target=lambda: asyncio.run(start_parsing_async())).start()


# Видалення атрибутів класу, ідентифікаторів та стилів перед очищенням HTML
def clean_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    for tag in soup.find_all():
        del tag['class']
        del tag['id']
        del tag['style']
    allowed_tags = ['p', 'ul', 'li', 'a', 'table', 'tr', 'td']  # Дозволені HTML-теги

    # Очищення HTML-коду від атрибутів <a> і залишення лише тексту
    for a_tag in soup.find_all('a'):
        a_tag.clear()  # Очистити вміст тегу <a>

    stripped_html = bleach.clean(str(soup), tags=allowed_tags, strip=True)
    return stripped_html



# Tkinter GUI setup
root = tk.Tk()
root.title("Web Page Parser")
root.geometry("600x500")

url_label = tk.Label(root, text="Page URL:")
url_label.pack()
url_entry = tk.Entry(root, width=50)
url_entry.pack()

link_pattern_label = tk.Label(root, text="Link Pattern (CSS Selector):")
link_pattern_label.pack()
link_pattern_entry = tk.Entry(root, width=50)
link_pattern_entry.pack()

check_button = tk.Button(root, text="Check Links", command=sync_check_links)
check_button.pack(pady=10)

product_title_label = tk.Label(root, text="Product Title CSS Selector:")
product_title_label.pack()
product_title_entry = tk.Entry(root, width=50)
product_title_entry.pack()

product_image_label = tk.Label(root, text="Product Image CSS Selector:")
product_image_label.pack()
product_image_entry = tk.Entry(root, width=50)
product_image_entry.pack()

product_description_label = tk.Label(root, text="Product Description CSS Selector:")
product_description_label.pack()
product_description_entry = tk.Entry(root, width=50)
product_description_entry.pack()

product_specs_label = tk.Label(root, text="Product Specs CSS Selector:")
product_specs_label.pack()
product_specs_entry = tk.Entry(root, width=50)
product_specs_entry.pack()

parse_button = tk.Button(root, text="Start Parsing", command=start_parsing)
parse_button.pack(pady=10)

save_button = tk.Button(root, text="Save Profile", command=save_profile)
save_button.pack(pady=5)

load_button = tk.Button(root, text="Load Profile", command=load_profile)
load_button.pack(pady=5)

status_label = tk.Label(root, text="")
status_label.pack()

root.mainloop()
