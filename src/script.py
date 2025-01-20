import aiohttp
import asyncio
import openpyxl
from xml.etree import ElementTree as ET


BASE_URL = "https://api.pbd.complexbar.ru"
HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
    "Authorization": "Basic cnhlc2lkZTpMb2VtYWt0djM5MTA=",
    "Cache-Control": "max-age=0",
    "Cookie": "PHPSESSID=b36cb5baded4d54141996cc0fd0c221c",
    "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Mobile Safari/537.36",
}


async def fetch(session, url, params=None):
    """Асинхронно получить данные по URL."""
    async with session.get(url, headers=HEADERS, params=params) as response:
        response.raise_for_status()
        return await response.text()


async def get_products(session, page):
    """Получить список товаров с указанной страницы."""
    url = f"{BASE_URL}/products"
    params = {"expand": "fileProduct", "page": page}
    response = await fetch(session, url, params)
    root = ET.fromstring(response)
    products = []

    for item in root.findall('item'):
        product_id = item.find('id').text
        article_number = item.find('article_number').text
        products.append({"id": product_id, "article_number": article_number})

    return products


async def get_product_details(session, product_id):
    """Получить количество штук в упаковке."""
    url = f"{BASE_URL}/prods"
    params = {"id": product_id}
    response = await fetch(session, url, params)
    root = ET.fromstring(response)

    quantity_element = root.find(".//props/item[title='Количество в упаковке (шт.)']/_value")
    return quantity_element.text if quantity_element is not None else None


async def process_products():
    """Обработать товары и сохранить в XLSX."""
    async with aiohttp.ClientSession() as session:
        page = 1
        seen_product_ids = set()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Товары"
        ws.append(["Артикул", "Штук в упаковке"])

        while True:
            print(f"Загружается страница {page}...")
            products = await get_products(session, page)
            if not products:
                print("Все страницы обработаны.")
                break

            new_products = [p for p in products if p["id"] not in seen_product_ids]
            if not new_products:
                print("Повторяющиеся товары. Завершаем обработку.")
                break

            seen_product_ids.update(p["id"] for p in new_products)

            tasks = [get_product_details(session, p["id"]) for p in new_products]
            details = await asyncio.gather(*tasks)

            for product, pieces_per_pack in zip(new_products, details):
                print(f"Товар {product['article_number']}, Штук в упаковке: {pieces_per_pack}")
                ws.append([product["article_number"], pieces_per_pack])

            page += 1

        wb.save("products.xlsx")


if __name__ == "__main__":
    asyncio.run(process_products())
