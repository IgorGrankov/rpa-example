import datetime
import pathlib

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables

browser = Selenium()
xls = Files()
table = Tables()

# Move to devdata/env.json
#test_dept_name = "National Science Foundation"

# Constants
BASE_URL = "https://itdashboard.gov"
DEFAULT_XLS_PATH = "output/totals.xlsx"
DOWNLOADS_DIR = f"{pathlib.Path(__file__).parent.resolve()}/output"
TIMEOUT = datetime.timedelta(seconds=10)

# Locators
DIVE_IN_BTN = "//a[@href='#home-dive-in']"
DEPT_NAME = "//a/span[contains (@class, 'h4')]"
DEPT_VALUE = "//a/span[contains (@class, 'h1')]"
TEST_DEPT = f"//span[text()='{TEST_DEPT}']"
TABLE_SELECTOR = "investments-table-object_length"
SECOND_PAGE = "//a[text()='2']"
UII_COLUMN_URLS = "//table[@id='investments-table-object']//td[contains(@class, 'left')][1]/a"
UII_COLUMN = "//table[@id='investments-table-object']//td[contains(@class, 'left')][1]"
BUREAU_COLUMN = "//table[@id='investments-table-object']//td[contains(@class, 'left')][2]"
TOTAL_COLUMN = "//table[@id='investments-table-object']//td[contains(@class, 'left')][3]"
TITLE_COLUMN = "//table[@id='investments-table-object']//td[contains(@class, 'right')]"
TYPE_COLUMN = "//table[@id='investments-table-object']//td[contains(@class, 'left')][4]"
RATING_COLUMN = "//table[@id='investments-table-object']//td[contains(@class, 'center')][1]"
NUM_COLUMN = "//table[@id='investments-table-object']//td[contains(@class, 'center')][2]"
PDF_LINK = "//a[text()='Download Business Case PDF']"


def open_the_website():
    browser.open_available_browser(BASE_URL)


def click_dive_in_button():
    browser.click_element_when_visible(DIVE_IN_BTN)


def click_test_dept():
    browser.click_element_when_visible(TEST_DEPT)


def get_individual_investments():
    browser.wait_until_element_is_visible(TABLE_SELECTOR, TIMEOUT)
    browser.select_from_list_by_label(TABLE_SELECTOR, "All")
    browser.wait_until_page_does_not_contain_element(SECOND_PAGE, TIMEOUT)
    browser.execute_javascript("document.querySelector('.dataTables_scrollBody').scrollTop=1500")

    uii = [browser.get_text(elem) for elem in browser.get_webelements(UII_COLUMN)]
    bureau = [browser.get_text(elem) for elem in browser.get_webelements(BUREAU_COLUMN)]
    total = [browser.get_text(elem) for elem in browser.get_webelements(TOTAL_COLUMN)]
    title = [browser.get_text(elem) for elem in browser.get_webelements(TITLE_COLUMN)]
    type = [browser.get_text(elem) for elem in browser.get_webelements(TYPE_COLUMN)]
    rating = [browser.get_text(elem) for elem in browser.get_webelements(RATING_COLUMN)]
    num_of_projects = [browser.get_text(elem) for elem in browser.get_webelements(NUM_COLUMN)]

    return {"UII": uii, "Bureau": bureau, "Total": total, "Title": title, "Type": type, "Rating": rating,
            "# of Projects": num_of_projects}


def get_agencies_names():
    browser.wait_until_element_is_visible(DEPT_NAME)
    names = [browser.get_text(elem) for elem in browser.get_webelements(DEPT_NAME)]
    return names


def get_agencies_totals():
    browser.wait_until_element_is_visible(DEPT_VALUE)
    values = [browser.get_text(elem) for elem in browser.get_webelements(DEPT_VALUE)]
    return values


def write_depts_data_to_excel(agencies=None, totals=None, path=DEFAULT_XLS_PATH):
    try:
        xls.create_workbook(path)
        data = table.create_table({"Departments": agencies, "Totals": totals})
        xls.append_rows_to_worksheet(data, header=True)
        xls.save_workbook()
    finally:
        xls.close_workbook()


def download_available_pdfs():
    browser.set_download_directory(DOWNLOADS_DIR)
    urls = [browser.get_element_attribute(elem, "href") for elem in browser.get_webelements(UII_COLUMN_URLS)]
    for url in urls:
        browser.open_available_browser(url)
        browser.click_element_when_visible(PDF_LINK)


def write_individual_investments_data(data=None, path=DEFAULT_XLS_PATH):
    try:
        xls.open_workbook(path)
        xls.create_worksheet("Individual Investments")
        data = table.create_table(data)
        xls.append_rows_to_worksheet(data, header=True)
        xls.save_workbook()
    finally:
        xls.close_workbook()


def main():
    try:
        open_the_website()
        click_dive_in_button()
        write_depts_data_to_excel(get_agencies_names(), get_agencies_totals())
        click_test_dept()
        write_individual_investments_data(get_individual_investments())
        download_available_pdfs()
    finally:
        browser.close_all_browsers()


if __name__ == "__main__":
    main()
