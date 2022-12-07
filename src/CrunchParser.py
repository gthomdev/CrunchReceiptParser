from bs4 import BeautifulSoup
from collections import OrderedDict
import pyexcel
import xlsxwriter
import os

def get_soup_from_file_path(file_path):
    try:
        return BeautifulSoup(open(file_path), 'html.parser')
    except Exception as e:
        print(e)
        raise


def get_crunch_table_nodes_rows_from_soup(soup):
    try:
        return soup.find_all('div', role='row', class_='table-row table-has-shadow')
    except Exception as e:
        print(e)
        raise


def get_total_from_table_row_node(table_row_node):
    return table_row_node.find('div', class_='money').string[1:]


def get_date_from_table_row_node(table_row_node):
    return table_row_node.find('span', class_='text-style u-margin--none text--minor').string


def get_supplier_from_table_row_node(table_row_node):
    return table_row_node.find('span', class_='text-style u-margin--none text--bold').string


def get_attachment_name_from_table_row_node(table_row_node):
    return table_row_node.find('div', class_='drop-zone--file--name').string


def get_paid_status_from_table_row_node(table_row_node):
    return table_row_node.find('div', class_='status-pill status-pill--bordered').string


def is_advertisement_table_row(table_row_node):
    return table_row_node.find('span', class_='button__icon-label', string='Upgrade to Pro') is not None


def get_parsed_crunch_table_row(table_row_node):
    return OrderedDict({
        'Date': get_date_from_table_row_node(table_row_node),
        'Supplier': get_supplier_from_table_row_node(table_row_node),
        'Total': get_total_from_table_row_node(table_row_node),
        'Paid Status': get_paid_status_from_table_row_node(table_row_node),
        'Attachment Name': get_attachment_name_from_table_row_node(table_row_node)
    })


def get_parsed_crunch_table(table_row_collection):
    parsed_table = []
    for row in table_row_collection:
        parsed_table.append(get_parsed_crunch_table_row(row))
    return parsed_table


def get_parsed_crunch_table_from_file_path(file_path):
    crunch_soup = get_soup_from_file_path(file_path)
    table_rows = get_crunch_table_nodes_rows_from_soup(crunch_soup)
    remove_advertisement_table_rows(table_rows)
    return get_parsed_crunch_table(table_rows)


def remove_advertisement_table_rows(table_row_collection):
    for row in table_row_collection:
        if is_advertisement_table_row(row):
            table_row_collection.remove(row)


def export_parsed_crunch_table_to_xls(parsed_table, file_path):
    pyexcel.save_as(records=parsed_table, dest_file_name=file_path)


def write_xlswriter_row(worksheet, row, row_data):
    for col, data in enumerate(row_data):
        worksheet.write(row, col, data)


def write_xlswriter_column_headers(worksheet, column_headers):
    for col, header in enumerate(column_headers):
        worksheet.write(0, col, header)


def main():
    # Example HTML file path: C:\Users\Downloads\crunch.html
    # Example Export Path: C:\Users\Downloads\crunch.xlsx
    html_path = r'C:\Users\George\Documents\gthomdev\CrunchReceiptParser\samples\Expenses _ Crunch.html'
    export_path = r'C:\Users\George\Documents\gthomdev\CrunchReceiptParser\samples\Expenses _ Crunch.xlsx'
    receipt_directory = 'C:\\Users\\Luis\\Desktop\\crunch\\Receipts\\'
    parsed_table = get_parsed_crunch_table_from_file_path(html_path)
    workbook = xlsxwriter.Workbook(export_path)
    worksheet = workbook.add_worksheet()
    write_xlswriter_column_headers(worksheet, parsed_table[0].keys())
    worksheet.write_url('F1', receipt_directory)

    for row, row_data in enumerate(parsed_table):
        write_xlswriter_row(worksheet, row + 1, row_data.values())
        file_path = os.path.join(receipt_directory, "01_" + row_data['Attachment Name'])
        worksheet.write_url(row + 1, 5, file_path)
    workbook.close()


if __name__ == '__main__':
    main()



