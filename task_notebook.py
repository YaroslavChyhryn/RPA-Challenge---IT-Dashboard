from RPA.Browser.Selenium import Selenium
import time
import xlsxwriter
from RPA.PDF import PDF
import re

# Declare agency for detail report
detail_report_agency = 'Department of Defense'
# detail_report_agency = ''
download_directory = '/home/yaroslav/Downloads/'


browser_lib = Selenium()
pdf = PDF()
# TODO set download direvtory
browser_lib.set_download_directory('.output')

browser_lib.open_chrome_browser("https://itdashboard.gov/")
browser_lib.click_link('#home-dive-in')
time.sleep(5)
browser_lib.wait_until_page_contains_element('//div[@id="home-dive-in" and @aria-expanded="true"]')

# Get agencies elements
agencies_widget = browser_lib.find_element('id:agency-tiles-widget')
agencies_elements = agencies_widget.find_elements_by_partial_link_text('Spending')

# +
# Parse agencies elements
agencies = {}

for agency in agencies_elements:
    data = agency.find_elements_by_tag_name('span')
    name = data[0]
    amounts = data[1]
    agencies[name.text] = {'amounts': amounts.text, 'element': agency}

# +
# Write agencies spendings to spreadsheet
workbook = xlsxwriter.Workbook('output/Agencies.xlsx')
worksheet = workbook.add_worksheet('agencies')

for row, agency in enumerate(agencies.items()):
    worksheet.write(row, 0, agency[0])
    worksheet.write(row, 1, agency[1]['amounts'])
# -

# Deteil report for specific agency
if agencies.get(detail_report_agency):
    element = agencies.get(detail_report_agency)['element']
    element.click()

    # Parse Investment Highlights
    browser_lib.wait_until_page_contains_element('id:agency-quick-stats-widget')
    investment_ighlights = browser_lib.find_element('id:agency-quick-stats-widget')
    investment_ighlights = investment_ighlights.text.split('\n')

    # Write Investment Highlights to new worksheet
    worksheet = workbook.add_worksheet(detail_report_agency)
    for row, item in enumerate(zip(investment_ighlights[0::2], investment_ighlights[1::2])):
        worksheet.write(row, 0, item[0])
        worksheet.write(row, 1, item[1])

    # Get list of UII
    browser_lib.wait_until_page_contains_element('id:investments-table-widget', 30)
    browser_lib.select_from_list_by_value('investments-table-object_length', '-1')
    browser_lib.wait_until_page_contains_element("//a[@class='paginate_button next disabled']", 30)

    investments_rows = browser_lib.find_element('id:investments-table-object')
    investments_rows = investments_rows.find_elements_by_xpath('//tr[.//a]')

    # Download pdf's
    for row in investments_rows[:2]:
        link = row.find_element_by_tag_name('a')

        investment_title = row.find_elements_by_tag_name('td')[2].text
        uii = link.text
        url = link.get_attribute('href')

        # Download pdf in new tab
        pdf_file = uii + '.pdf'
        browser_lib.driver.execute_script("window.open();")
        browser_lib.driver.switch_to.window(browser_lib.driver.window_handles[1])
        browser_lib.driver.get(url)
        browser_lib.wait_until_page_contains_element('id:business-case-pdf')
        download_link = browser_lib.find_element('link:Download Business Case PDF').click()
        time.sleep(5)
        browser_lib.driver.close()
        browser_lib.driver.switch_to.window(browser_lib.driver.window_handles[0])

        # Open PDF
        pdf.open_pdf(download_directory+pdf_file)
        pdf.convert(download_directory+pdf_file)

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
        list_of_values += [section_a[separator_3_idx+3:separator_4_idx]]
        list_of_values += [section_a[separator_4_idx+3:]]

        # Store values
        section_a_values = {}

        for item in list_of_values:
            name, value = item.split(':')
            section_a_values[name.strip()] = value.strip()

        # Compare values from pdf with table on agency page
        investment_title_in_pdf = section_a_values['Name of this Investment'].replace("\n", " ")
        uii_in_pdf = section_a_values['Unique Investment Identifier (UII)']

        print(f"{investment_title} is {'same' if investment_title==investment_title_in_pdf else 'different'}")
        print(f"{uii} is {'same' if uii==uii_in_pdf else 'different'}")
        pdf.close_pdf(download_directory+pdf_file)

workbook.close()

# +
pdf_file='007-000100123.pdf'
pdf.open_pdf(download_directory+pdf_file)
pdf.convert(download_directory+pdf_file)

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
list_of_values += [section_a[separator_3_idx+3:separator_4_idx]]
list_of_values += [section_a[separator_4_idx+3:]]

print(section_a[:separator_1_idx].split('\n'))
print(section_a[separator_1_idx:separator_2_idx])
print(section_a[separator_2_idx:separator_3_idx])
print(section_a[separator_3_idx+3:separator_4_idx])
print(section_a[separator_4_idx+3:])

# Store values
section_a_values = {}

# for item in list_of_values:
#     name, value = item.split(':')
#     section_a_values[name.strip()] = value.strip()
pdf.close_pdf(download_directory+pdf_file)
