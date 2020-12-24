import time
import xlsxwriter

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.remote.webdriver import WebDriver


def attach_to_session(executor_url, session_id):
    original_execute = WebDriver.execute

    def new_command_execute(self, command, params=None):
        if command == "newSession":
            # Mock the response
            return {'success': 0, 'value': None, 'sessionId': session_id}
        else:
            return original_execute(self, command, params)

    WebDriver.execute = new_command_execute
    driver = webdriver.Remote(command_executor=executor_url, desired_capabilities={})
    driver.session_id = session_id
    WebDriver.execute = original_execute
    return driver


def main():
    driver = webdriver.Chrome("chromedriver.exe")
    # driver = webdriver.Firefox()
    executor_url = driver.command_executor._url
    session_id = driver.session_id
    bro = attach_to_session(executor_url, session_id)
    bro.get('https://www.ifcg.ru/en/kb/tnved/')

    workbook = xlsxwriter.Workbook('Товарная позиция (6 символа).xlsx')
    bold = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet("Товарная позиция (6 сивола)")
    worksheet.write('A1', 'Код', bold)
    worksheet.write('B1', 'Название', bold)

    t = 3
    nums = 2
    time.sleep(3)
    links = []
    span = driver.find_elements_by_class_name('description')[80:]
    for i in span:
        href = i.find_element_by_tag_name('a').get_attribute('href')
        links.append(href)
    s2 = []
    while links:
        s = []
        for i in links:
            bro.get(i)
            time.sleep(0.5)
            span = bro.find_elements_by_class_name('description')
            for y in span:
                try:
                    a = y.find_element_by_tag_name('a')
                    href = a.get_attribute('href')
                    if href in s2:
                        pass
                    else:
                        code = href.split('https://www.ifcg.ru/en/kb/tnved/')[1].replace('/', '')
                        title = a.text
                        if len(code) == 6:
                            print(f"{code} {title}")
                            worksheet.write(f'A{nums}', code)
                            worksheet.write(f'B{nums}', title)
                            nums += 1
                        s.append(href)
                        s2.append(href)
                except NoSuchElementException:
                    pass
        links = [t for t in s]

    bro.close()
    workbook.close()


if __name__ == "__main__":
    main()
