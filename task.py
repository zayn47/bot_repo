# +
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from time import sleep
from RPA.Dialogs import Dialogs
from RPA.Archive import Archive
from RPA.Robocorp.Vault import Vault

_secret = Vault().get_secret("credentials")
browser = Selenium()
excel = Files()
pdf = PDF()
dialogs = Dialogs()
archive = Archive()


# +
def login():
    browser.input_text('xpath://*[@id="username"]', _secret["username"])
    browser.input_text('xpath://*[@id="password"]', _secret["password"])
    browser.find_element('xpath://*[@id="root"]/div/div/div/div[1]/form/button').click()


def open_browser():
    browser.open_available_browser("https://robotsparebinindustries.com/", maximized=True)
    login()
    browser.go_to("https://robotsparebinindustries.com/#/robot-order")
    browser.click_button('OK')


def dynamic_form_filling():
    dialogs.add_file_input(name="orders")
    user = dialogs.run_dialog()
    excel.open_workbook(path=user.orders[0])
    order_data = excel.read_worksheet(header=True)

    for item in order_data:

        browser.select_from_list_by_value("head", str(item['Head']))
        browser.select_radio_button("body", str(item['Body']))
        browser.input_text('xpath://html/body/div/div/div[1]/div/div[1]/form/div[3]/input', str(item['Legs']))
        browser.input_text('xpath://*[@id="address"]', str(item['Address']))
        browser.find_element('xpath://*[@id="preview"]').click()
        sleep(2)
        browser.find_element('xpath://*[@id="order"]').click()
        while True:
                try:
                    browser.find_element("id:receipt")
                    break
                except:
                    browser.click_button("Order")

        browser.wait_until_page_contains_element('id:receipt')
        recipt_table = browser.get_element_attribute('id:receipt', 'outerHTML')
        pdf.html_to_pdf(recipt_table, f'output/recipts/recipt_robot_{item["Order number"]}.pdf')
        browser.screenshot('robot-preview-image', f'output/bot_{item["Order number"]}.png')
        pdf.add_watermark_image_to_pdf(
            image_path=f'output/bot_{item["Order number"]}.png',
            source_path=f'output/recipts/recipt_robot_{item["Order number"]}.pdf',
            output_path=f'output/recipts/recipt_robot_{item["Order number"]}.pdf')

        browser.click_button('Order another robot')
        browser.click_button('OK')
    browser.close_window()


def zip_recipt():
    archive.archive_folder_with_zip(folder='output/recipts',archive_name='output/recipts.zip')


# -

if __name__ == "__main__":
    open_browser()
#     dynamic_form_filling()
#     zip_recipt()







