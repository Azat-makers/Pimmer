
import pandas as pd
from lxml import etree

# Загрузка данных из Excel
file_path = 'Книга2_test (4) (2).xlsx'
data = pd.read_excel(file_path, header=1)

# Приведение всех значений к строковому типу
data = data.astype(str)

# Функция для создания YML-документа
def create_yml(data):
    # Создание корневого элемента
    root = etree.Element("yml_catalog", date="2024-05-31")

    # Создание дочернего элемента <shop>
    shop = etree.SubElement(root, "shop")
    etree.SubElement(shop, "name").text = "MyName"
    etree.SubElement(shop, "company").text = "MyCompany"

    # Создание элемента <currencies>
    currencies = etree.SubElement(shop, "currencies")
    currency = etree.SubElement(currencies, "currency", id="KGS", rate="1")

    # Создание элемента <categories>
    categories = etree.SubElement(shop, "categories")
    category_ids = data['categoryid'].unique()
    for cat_id in category_ids:
        category = etree.SubElement(categories, "category", id=cat_id)
        category.text = "Категория " + cat_id  # Использование уникального идентификатора категории

    # Создание элемента <offers>
    offers = etree.SubElement(shop, "offers")

    # Добавление предложений из данных
    for index, row in data.iterrows():
        offer = etree.SubElement(offers, "offer", id=row['offer id'], available="true")
        etree.SubElement(offer, "url").text = "http://www.example.com/product/{}".format(row['offer id'])
        etree.SubElement(offer, "price").text = row['price']
        etree.SubElement(offer, "currencyId").text = row['currencyId']
        etree.SubElement(offer, "categoryId").text = row['categoryid']
        etree.SubElement(offer, "picture").text = row['picture']
        if pd.notna(row['picture.1']):
            etree.SubElement(offer, "picture").text = row['picture.1']

        # Добавление параметров из данных, исключая categoryid
        param_columns = data.columns[11:]  # все колонки, начиная с 11-й, это параметры
        for param_name in param_columns:
            if param_name != 'categoryid':
                etree.SubElement(offer, "param", name=param_name).text = row[param_name]

    # Преобразование в строку
    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8").decode("utf-8")

# Создание YML-документа
yml_string = create_yml(data)

# Сохранение в файл
output_file_path = 'generated_test_5.xml'
with open(output_file_path, 'w', encoding='utf-8') as file:
    file.write(yml_string)

print(f"YML файл сохранен в {output_file_path}")
