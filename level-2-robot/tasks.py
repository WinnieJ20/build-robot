from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive


import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100
    )

    open_robot_order_website()
    download_order_file()
    fill_orders_and_save()
    archive_receipts()


def open_robot_order_website():
    """
    Open the robot order window
    """
    browser.goto('https://robotsparebinindustries.com/#/robot-order')

def download_order_file():
    """Downloads order excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """
    Extract data from downloaded csv file
    """
    csv = Tables()
    orders = csv.read_table_from_csv("orders.csv", columns=['Order number', 'Head','Body','Legs','Address'])
    return orders

def fill_orders_and_save():
    """
    Loop the data to get the orders in
    """
    orders = get_orders()
    for row in orders:
        close_annoying_modal()
        fill_the_form_and_preview(row)
        submit_order()
        store_receipt_as_pdf(row['Order number'])
        new_order()
        break

def fill_the_form_and_preview(row):
    """
    Fill in one single row of data and submit the order
    """
    page = browser.page()

    page.select_option("#head", row['Head'])
    page.check(f'//*[@id="id-body-{row["Body"]}"]')
    page.fill('.form-control', row['Legs'])
    page.fill("#address", row['Address'])

    # Preview order
    page.click("text=Preview")

def close_annoying_modal():
    """
    Get rid of the constitutional rights pop up
    """
    page = browser.page()
    page.click("text=OK")

def submit_order():
    page = browser.page()

    page.click("button:text('Order')")
    while not page.query_selector("#order-another"):
        page.click("button:text('Order')")


def new_order():
    page = browser.page()

    page.click("text=Order another robot")

def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    if not os.path.exists('output/receipts'):
        os.makedirs('output/receipts')
    
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, f"output/receipts/{order_number}_receipt.pdf")
    screenshot_name = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot_name + ':align=center', f"output/receipts/{order_number}_receipt.pdf", pdf)


def screenshot_robot(order_number) -> str:
    page = browser.page()
    element = page.locator("#robot-preview-image")
    element.screenshot(path=f"output/receipts/{order_number}_robot.png")
    return f"output/receipts/{order_number}_robot.png"

def embed_screenshot_to_receipt(screenshot: str, target_path: str, pdf_file: PDF):
    pdf_file.add_files_to_pdf(
        files=[screenshot, ], 
        target_document=target_path, 
        append=True)

def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'output/receipts.zip', recursive=True)