"""
This Robot automates the process of extracting data from itdashboard.gov.
"""
from RPA.Browser.Selenium import Selenium
import time
import xlsxwriter
from RPA.PDF import PDF
import re


browser_lib = Selenium()
pdf = PDF()
workbook = xlsxwriter.Workbook('output/Agencies.xlsx')

browser_lib.set_download_directory('.output')

# Declare agency for detail report
DETAIL_REPORT_AGENCY = 'Department of Defense'
DOWNLOAD_DIRECTORY = '/home/yaroslav/Downloads/'
URL = "https://itdashboard.gov/"


def parse_agencies():
    """
    Parse Agency name and their spending

    :return: agencies
    """
    browser_lib.click_link('#home-dive-in')
    time.sleep(5)
    browser_lib.wait_until_page_contains_element('//div[@id="home-dive-in" and @aria-expanded="true"]')

    # Get agencies elements
    agencies_widget = browser_lib.find_element('id:agency-tiles-widget')
    agencies_elements = agencies_widget.find_elements_by_partial_link_text('Spending')

    # Parse agencies elements
    agencies = {}

    for agency in agencies_elements:
        data = agency.find_elements_by_tag_name('span')
        name = data[0]
        amounts = data[1]
        agencies[name.text] = {'amounts': amounts.text, 'element': agency}

    return agencies


def write_agencies_to_excel(agencies):
    worksheet = workbook.add_worksheet('agencies')

    for row, agency in enumerate(agencies.items()):
        worksheet.write(row, 0, agency[0])
        worksheet.write(row, 1, agency[1]['amounts'])

    return


def write_agency_detail_to_excel(investment_ighlights):
    worksheet = workbook.add_worksheet(DETAIL_REPORT_AGENCY)

    for row, item in enumerate(zip(investment_ighlights[0::2], investment_ighlights[1::2])):
        worksheet.write(row, 0, item[0])
        worksheet.write(row, 1, item[1])

    return


def parse_investments_of_agency():
    """
    Parse investment highlights of agency

    :return: investment_highlights
    """
    browser_lib.wait_until_page_contains_element('id:agency-quick-stats-widget')
    investment_highlights = browser_lib.find_element('id:agency-quick-stats-widget')
    investment_highlights = investment_highlights.text.split('\n')

    return investment_highlights


def compare_investment_title_and_uii(uii, uii_in_pdf, investment_title, investment_title_in_pdf):
    """
    Compare uii and investment title from agency page with pdf
    """
    print(f"{investment_title} is {'same' if investment_title == investment_title_in_pdf else 'different'}")
    print(f"{uii} is {'same' if uii == uii_in_pdf else 'different'}")


def parse_uii():
    """
    Parse investments from investments table on agency page
    """
    browser_lib.wait_until_page_contains_element('id:investments-table-widget', 30)
    browser_lib.select_from_list_by_value('investments-table-object_length', '-1')
    browser_lib.wait_until_page_contains_element("//a[@class='paginate_button next disabled']", 30)

    investments_rows = browser_lib.find_element('id:investments-table-object')
    investments_rows = investments_rows.find_elements_by_xpath('//tr[.//a]')

    for row in investments_rows:
        link = row.find_element_by_tag_name('a')

        investment_title = row.find_elements_by_tag_name('td')[2].text
        uii = link.text
        url = link.get_attribute('href')

        download_pdf(url)
        uii_in_pdf, investment_title_in_pdf = parse_pdf(uii)
        compare_investment_title_and_uii(uii, uii_in_pdf, investment_title, investment_title_in_pdf)


def download_pdf(url):
    browser_lib.driver.execute_script("window.open();")
    browser_lib.driver.switch_to.window(browser_lib.driver.window_handles[1])
    browser_lib.driver.get(url)
    browser_lib.wait_until_page_contains_element('id:business-case-pdf')
    browser_lib.find_element('link:Download Business Case PDF').click()
    time.sleep(5)
    browser_lib.driver.close()
    browser_lib.driver.switch_to.window(browser_lib.driver.window_handles[0])


def parse_pdf(uii):
    """ Parse Section A"""
    pdf_file = uii + '.pdf'

    pdf.open_pdf(DOWNLOAD_DIRECTORY + pdf_file)
    pdf.convert(DOWNLOAD_DIRECTORY + pdf_file)

    pages = pdf.get_text_from_pdf()

    # Parse Section A
    for page in pages.values():
        start = re.search(r'Section A: ', page).end()
        end = re.search(r'Section B:', page).start()
        if start and end:
            break

    section_a = page[start: end]

    separator_1_idx = section_a.find('Agency:')
    separator_2_idx = section_a.find('Bureau:')
    separator_3_idx = section_a.find('1.')
    separator_4_idx = section_a.find('2.')

    list_of_values = section_a[:separator_1_idx].split('\n')
    list_of_values += [section_a[separator_1_idx:separator_2_idx]]
    list_of_values += [section_a[separator_2_idx:separator_3_idx]]
    list_of_values += [section_a[separator_3_idx + 3:separator_4_idx]]
    list_of_values += [section_a[separator_4_idx + 3:]]

    # Store values
    section_a_values = {}

    for item in list_of_values:
        try:
            name, value = item.split(':')
        except ValueError:
            continue
        section_a_values[name.strip()] = value.strip()

    investment_title_in_pdf = section_a_values['Name of this Investment'].replace("\n", " ")
    uii_in_pdf = section_a_values['Unique Investment Identifier (UII)']

    pdf.close_pdf(DOWNLOAD_DIRECTORY + pdf_file)

    return uii_in_pdf, investment_title_in_pdf


def detail_agency_report(agencies):
    """Detail report for DETAIL_REPORT_AGENCY"""
    element = agencies.get(DETAIL_REPORT_AGENCY)['element']
    element.click()

    investment_highlights = parse_investments_of_agency()
    write_agency_detail_to_excel(investment_highlights)

    parse_uii()


def main():
    try:
        browser_lib.open_chrome_browser(URL)
        agencies = parse_agencies()
        write_agencies_to_excel(agencies)

        if agencies.get(DETAIL_REPORT_AGENCY):
            detail_agency_report(agencies)

    finally:
        browser_lib.close_all_browsers()
        workbook.close()


if __name__ == "__main__":
    main()
