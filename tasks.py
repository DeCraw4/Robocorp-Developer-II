from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive

import os 
import logging


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    close_annoying_modal()
    download_csv_file()
    orders = read_csv_file()
    fill_forms(orders)
    fill_and_submit
    handle_submission
    cleanup()


def open_robot_order_website():
    # opens robot order website
    browser.configure(slowmo=200)
    page = browser.page()
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def close_annoying_modal():
    """closes pop-up, forgos all your rights"""
    page = browser.page()
    page.click("text=I guess so...")


def download_csv_file():
    """ dowloads excel file from given URL """
    http = HTTP()
    http.download(
        "https://robotsparebinindustries.com/orders.csv",
    "output/orders.csv", overwrite=True
    )


def read_csv_file():
    # reads content of csv file without editing
    data = Tables().read_table_from_csv("output/orders.csv")
    return data


def fill_forms(orders):
    """completes order form(s) from csv file(s) when 
    combined with fill_and_submit function
    """
    page = browser.page()

    for order in orders:
        fill_and_submit(page, order)


def fill_and_submit(page, order):
    page = browser.page()

    page.select_option("#head", str(order["Head"]))
    page.set_checked("#id-body-" + str(order["Body"]), True)
    page.fill('input[type="number"]', str(order["Legs"]))
    page.fill('input#address', order["Address"])

    page.click("text=Preview")
    page.wait_for_selector("#robot-preview-image")
    
    # completes the order
    page.click("#order")
    
    # handle potential errors in submission
    handle_submission(page)

    # wait for the receipt to be visible and screenshot
    page.wait_for_selector("#receipt", timeout=30000)
    order_number = page.locator("#receipt .badge-success").inner_text()
    
    # works with later function to create a PDF file of HTML receipts 
    pdf_path = store_receipt_as_pdf(page, order_number)

    # order progression with error handling to complete core bot functionality 
    try:
        page.click("#order-another", timeout=5000)
    except:
        print(f"Couldn't find 'Order another robot' button for order {order_number}")
        # page should reload in errors during order process and proceed to complete orders
        page.reload()

    close_annoying_modal()
    

def handle_submission(page):
    is_alert_visible = page.locator("//div[@class='alert alert-danger']").is_visible(timeout=5000)

    while is_alert_visible:
        page.click("#order")
        is_alert_visible = page.locator("//div[@class='alert alert-danger']").is_visible(timeout=5000)

        if not is_alert_visible:
            break


def store_receipt_as_pdf(page, order_number):
    """ 
    stores recept as PDF outside of core bot functionality 
    and creates path for HTML to PDF conversion
    """
    page = browser.page()
    pdf = PDF()
    pdf_path = f"output/receipts/order_{order_number}.pdf"
    
    # creates directory path for html to pdf conversion 
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(order_receipt_html, pdf_path)

    # screenshot of completed robot(s)
    screenshot_path = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot_path, pdf_path)
    
    return pdf_path


def screenshot_robot(order_number):
    """Takes a screenshot of the robot and saves it."""
    page = browser.page()
    screenshot_path = f"output/receipts/order_{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path


def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the robot screenshot into the PDF receipt."""
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )


def archive_receipts():
    """Creates a ZIP archive of all PDF receipts."""
    lib = Archive()
    lib.archive_folder_with_zip("output/receipts", 'receipts.zip', recursive=True)


def cleanup(): 
    """Add any functions necessary to clear processed orders and prevent errors"""
    #TODO crete functions
 