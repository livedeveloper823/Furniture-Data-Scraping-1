from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium import webdriver
from time import sleep
import openpyxl, json, re 
import func

urls = [
    "https://www.homecenter.com.co/homecenter-co/category/cat10334/sofas/?cid=cat10332ftg#Sofas",
    "https://www.homecenter.com.co/homecenter-co/category/cat10340/sofa-camas/?cid=cat10332cat#SofaCama",
    "https://www.homecenter.com.co/homecenter-co/category/cat10328/sillas-reclinables-y-descanso/?cid=cat10332cat#Reclinables",
    "https://www.homecenter.com.co/homecenter-co/category/cat3690001/juegos-de-sala/?cid=cat10332cat#JuegoSala",
    "https://www.homecenter.com.co/homecenter-co/category/cat80004/puffs/?cid=cat10332cat#Puffs",  
    "https://www.homecenter.com.co/homecenter-co/category/cat1670283/sillones-y-poltronas/?cid=cat10332cat#Poltronas",
    "https://www.homecenter.com.co/homecenter-co/category/cat1670036/centros-de-entretenimiento-y-muebles-para-tv/?cid=cat10332cat#MuebleTV",
    "https://www.homecenter.com.co/homecenter-co/category/cat10344/juegos-de-comedor/?cid=cat10344cat#JuegosComedor",
    "https://www.homecenter.com.co/homecenter-co/category/cat1670091/mesas-de-comedor/?cid=cat10344cat#MesasComedor",
    "https://www.homecenter.com.co/homecenter-co/category/cat4840042/consolas-y-bifes/?cid=cat10344cat#Consolas",
    "https://www.homecenter.com.co/homecenter-co/category/cat10338/sillas-para-bar/?cid=cat10344cat#SillasBar", 
    "https://www.homecenter.com.co/homecenter-co/category/cat1670118/muebles-para-bar/?cid=cat10344cat#Bares",
    "https://www.homecenter.com.co/homecenter-co/category/cat10310/closets-y-armarios/?cid=cat10308cat#Closet",
    "https://www.homecenter.com.co/homecenter-co/category/cat1670037/mesas-de-noche/?cid=cat10308cat#MesaNoche",
    "https://www.homecenter.com.co/homecenter-co/category/cat1660049/camas/?cid=cat10308cat#Camas",
    "https://www.homecenter.com.co/homecenter-co/category/cat6040003/comodas-y-tocadores/?cid=cat10308cat#Comodas",
    "https://www.homecenter.com.co/homecenter-co/category/cat1370016/cabeceros-para-cama/?cid=cat10308cat#Cabeceros",
    "https://www.homecenter.com.co/homecenter-co/category/cat1660051/camas-infantiles-y-cunas/?cid=cat10308cat#CamaInfantil",
    "https://www.homecenter.com.co/homecenter-co/category/cat90508/colchones-sencillos/?cid=cat210003ftg#Sencillos",
    "https://www.homecenter.com.co/homecenter-co/category/cat90509/colchones-semidobles/?cid=cat210003ftg#Semidobles",
    "https://www.homecenter.com.co/homecenter-co/category/cat90510/colchones-dobles/?cid=cat210003ftg#Dobles",
    "https://www.homecenter.com.co/homecenter-co/category/cat90268/colchones-queen/?cid=cat210003ftg#Queen",
    "https://www.homecenter.com.co/homecenter-co/category/cat90622/colchones-king/?cid=cat210003ftg#King",
    "https://www.homecenter.com.co/homecenter-co/category/cat740055/base-cama-y-colchon/?cid=cat210003ftg#BaseCol"

    ]

func.create_excel()
match_num = 0

driver = webdriver.Chrome()
driver.maximize_window()

for url in urls:

    driver.get(url)
    func.wait_url(driver, url)

    product_urls = []
    # while True:
        # try:
    driver.find_element(By.CLASS_NAME, "pagination-and-back-to-top").find_element(By.ID, "bottom-pagination-next-page")
    products = func.find_element(driver, By.CLASS_NAME, "search-results-products-container").find_elements(By.CLASS_NAME, "product-wrapper")

    for product in products:
        product_url = product.find_element(By.TAG_NAME, "a").get_attribute("href")
        print(product_url)
        product_urls.append(product_url)
            # func.find_element(driver, By.CLASS_NAME, "pagination-and-back-to-top").find_element(By.ID, "bottom-pagination-next-page").click()
    sleep(2)
        # except:
        #     break


    for each_product_url in product_urls:
        match_num += 1
        workbook = openpyxl.load_workbook("homecenter.xlsx")
        sheet = workbook['Sheet']
        driver.get(each_product_url)
        sleep(2)
        try:
            driver.find_element(By.CLASS_NAME, "out-of-stock-info")
            pass
        except:
            product_name = func.find_element(driver, By.CLASS_NAME, "product-title").text
            sheet[f'D{match_num + 2}'] = product_name
            product_price = re.findall(r'\d+', func.find_element(driver, By.CLASS_NAME, "regular-price").find_element(By.CLASS_NAME, "primary").text.replace(".", ""))[0]
            sheet[f'E{match_num + 2}'] = product_price

            try:
                product_detail_box = driver.find_element(By.ID, "Ficha técnica").find_element(By.CLASS_NAME, "group-container")
                product_details = product_detail_box.find_elements(By.CLASS_NAME, "sub-table-container")
                for properties in product_details:
                    classification = properties.find_element(By.CLASS_NAME, "group-title").text
                    print(classification)
                    if classification == "Dimensiones":
                        characteristics = properties.find_elements(By.CLASS_NAME, "attribute") 
                        for characteristic in characteristics:
                                if "Ancho" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                    width = characteristic.find_element(By.CLASS_NAME, "value").text
                                    sheet[f'F{match_num + 2}'] = width
                                elif "Alto" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                    height = characteristic.find_element(By.CLASS_NAME, "value").text
                                    sheet[f'G{match_num + 2}'] = height
                                elif characteristic.find_element(By.CLASS_NAME, "key").text == "Peso":
                                    weight = characteristic.find_element(By.CLASS_NAME, "value").text
                                    sheet[f'H{match_num + 2}'] = weight
                                elif "Fondo" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                    back = characteristic.find_element(By.CLASS_NAME, "value").text
                                    sheet[f'I{match_num + 2}'] = back
                                elif "Largo" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                    length = characteristic.find_element(By.CLASS_NAME, "value").text
                                    sheet[f'J{match_num + 2}'] = length
                                elif characteristic.find_element(By.CLASS_NAME, "key").text == "Profundidad":
                                    depth = characteristic.find_element(By.CLASS_NAME, "value").text
                                    sheet[f'K{match_num + 2}'] = depth
                                elif characteristic.find_element(By.CLASS_NAME, "key").text == "Peso máximo soportado":
                                    max_weight = characteristic.find_element(By.CLASS_NAME, "value").text
                                    sheet[f'L{match_num + 2}'] = max_weight
                                elif characteristic.find_element(By.CLASS_NAME, "key").text == "Dimensiones de la cama":
                                    bed_demensions = characteristic.find_element(By.CLASS_NAME, "value").text
                                    sheet[f'M{match_num + 2}'] = bed_demensions
                    elif classification == "Especificaciones":
                        characteristics = properties.find_elements(By.CLASS_NAME, "attribute")
                        for characteristic in characteristics: 
                            if characteristic.find_element(By.CLASS_NAME, "key").text == "Tipo":
                                tipo = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'N{match_num + 2}'] = tipo
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Número de cuerpos":
                                body_num = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'O{match_num + 2}'] = body_num
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Material del tapiz":
                                tapestry_material = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'P{match_num + 2}'] = tapestry_material
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Términos Garantía":
                                warranty_terms = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'Q{match_num + 2}'] = warranty_terms
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "País de Origen":
                                original = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'R{match_num + 2}'] = original
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Material de la estructura":
                                structure_material = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'S{match_num + 2}'] = structure_material
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Garantía Producto":
                                product_guarantee = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'T{match_num + 2}'] = product_guarantee
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Color del tapiz":
                                tapestry_color = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'U{match_num + 2}'] = tapestry_color
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Material":
                                material = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'V{match_num + 2}'] = material
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Material Tapiz":
                                tapiz_material = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'O{match_num + 2}'] = tapiz_material
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Color":
                                color = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'W{match_num + 2}'] = color
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Número de puestos":
                                position_num = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'X{match_num + 2}'] = position_num
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Pulgadas TV":
                                tv_inch = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'Y{match_num + 2}'] = tv_inch
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Diseño de la Mesa":
                                table_design = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'Z{match_num + 2}'] = table_design
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Plegable":
                                folding = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AA{match_num + 2}'] = folding
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Cuenta con ruedas":
                                wheels = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AB{match_num + 2}'] = wheels
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Modo de fijación":
                                fix_mode = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AC{match_num + 2}'] = fix_mode
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Espacio recomendado":
                                space = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AD{match_num + 2}'] = space
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Cantidad de repisas":
                                shelve_num = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AE{match_num + 2}'] = shelve_num
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Cantidad de cajones":
                                drawer_num = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AF{match_num + 2}'] = drawer_num
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Cantidad de puertas":
                                door_num = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AG{match_num + 2}'] = door_num
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Tipo de guadarropa":
                                wardrobe_type = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AH{match_num + 2}'] = wardrobe_type
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Sistema de apertura":
                                opening_sys = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AI{match_num + 2}'] = opening_sys
                            # elif characteristic.find_element(By.CLASS_NAME, "key").text == "Peso del producto":
                            #     product_weight = characteristic.find_element(By.CLASS_NAME, "value").text
                            #     sheet[f'AJ{match_num + 2}'] = product_weight
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Número de cajones":
                                drawers_num = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AJ{match_num + 2}'] = drawers_num
                            # elif characteristic.find_element(By.CLASS_NAME, "key").text == "Acabado":
                            #     finish = characteristic.find_element(By.CLASS_NAME, "value").text
                            #     sheet[f'AK{match_num + 2}'] = finish
                            # elif characteristic.find_element(By.CLASS_NAME, "key").text == "Edad recomendada":
                            #     recommand_age = characteristic.find_element(By.CLASS_NAME, "value").text
                            #     sheet[f'AL{match_num + 2}'] = recommand_age
                            # elif characteristic.find_element(By.CLASS_NAME, "key").text == "Alto (centímetros)":
                            #     specific_height = characteristic.find_element(By.CLASS_NAME, "value").text
                            #     sheet[f'AM{match_num + 2}'] = specific_height
                            # elif characteristic.find_element(By.CLASS_NAME, "key").text == "Ancho (centímetros)":
                            #     specific_width = characteristic.find_element(By.CLASS_NAME, "value").text
                            #     sheet[f'AN{match_num + 2}'] = specific_width
                            # elif characteristic.find_element(By.CLASS_NAME, "key").text == "Largo (centímetros)":
                            #     specific_length = characteristic.find_element(By.CLASS_NAME, "value").text
                            #     sheet[f'AO{match_num + 2}'] = specific_length
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Tamaño":
                                size = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AK{match_num + 2}'] = size
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Nivel de Confort":
                                comfort_level = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AL{match_num + 2}'] = comfort_level
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Composición Interna":
                                internal_composition = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AM{match_num + 2}'] = internal_composition
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Línea":
                                line = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AN{match_num + 2}'] = line
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Material tapiz del colchon":
                                material_tapiz = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'O{match_num + 2}'] = material_tapiz
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Color del Colchón":
                                mattress_color = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'AO{match_num + 2}'] = mattress_color
                            elif "Ancho" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                width = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'F{match_num + 2}'] = width
                            elif "Alto" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                height = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'G{match_num + 2}'] = height
                            elif characteristic.find_element(By.CLASS_NAME, "key").text == "Peso":
                                weight = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'H{match_num + 2}'] = weight
                            elif "Fondo" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                back = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'I{match_num + 2}'] = back
                            elif "Largo" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                length = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'J{match_num + 2}'] = length
                            elif "Profundidad" in characteristic.find_element(By.CLASS_NAME, "key").text:
                                depth = characteristic.find_element(By.CLASS_NAME, "value").text
                                sheet[f'K{match_num + 2}'] = depth
                        else:
                            pass
                    image_url = driver.find_element(By.CLASS_NAME, "product-image-holder").find_element(By.CLASS_NAME, "product-image").find_element(By.CLASS_NAME, "zoom-class").find_elements(By.TAG_NAME, "img")[1].get_attribute("src")
                    sheet[f'AU{match_num + 2}'] = image_url
                    workbook.save("homecenter.xlsx")
            except:
                pass
    match_num += 2