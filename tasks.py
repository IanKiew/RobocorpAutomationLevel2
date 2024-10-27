from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
import os
from RPA.Archive import Archive
import shutil

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
        slowmo=200,
    )
    open_robot_order_website()
    fill_the_form()
    archive_receipts()
    clean_up()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Downloads order CSV file from the given URL and reads it into a table"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    tables = Tables()
    table = tables.read_table_from_csv("orders.csv")
    
    return table

def close_annoying_modal():
    """Close the annoying modal that pops up when you open the robot order website"""
    page = browser.page()  
    page.click("text=ok")

def proceed_to_next_order():
    """Proceed to next order"""
    page = browser.page()  
    page.click("#order-another")

def fill_the_form():
    """Fill in the order data"""
    orders = get_orders()
    for row in orders:
        page = browser.page()
        close_annoying_modal()

        page.select_option("#head", str(row['Head']))
        page.click(f"input[name='body'][value='{str(row['Body'])}']")
        page.fill("input[type='number'][placeholder='Enter the part number for the legs']", str(row['Legs']))
        page.fill("#address", row['Address'])
        page.click("#preview")

        # Attempt to submit the order with retries
        max_attempts = 3
        attempts = 0
        success = False
        
        while attempts < max_attempts and not success:
            page.click("#order")
            page.wait_for_timeout(2000) 

            # Check if error message appears
            if page.is_visible("div.alert.alert-danger"):
                print("Error encountered. Retrying...")
                attempts += 1
            else:
                success = True 

        if success:
            order_number = str(row['Order number'])
            pdf_path = store_receipt_as_pdf(order_number)
            screenshot_path = screenshot_robot(order_number)
            embed_screenshot_to_receipt(screenshot_path, pdf_path)
        else:
            print(f"Order submission failed for order number {row['Order number']} after {max_attempts} attempts.")

        proceed_to_next_order()

def store_receipt_as_pdf(order_number):
    """Export the receipt as a pdf file"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    # Define the output directory
    output_dir = "output/Receipts"
    os.makedirs(output_dir, exist_ok=True)  # Create the directory if it doesn't exist

    # Create the filename based on the order number
    filename = f"Receipt_{order_number}.pdf"
    filepath = os.path.join(output_dir, filename)

    # Convert HTML to PDF and save it to the specified path
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, filepath)
    return filepath

def screenshot_robot(order_number):
    """Take a screenshot of the robot"""
    page = browser.page()
    robot = page.locator("#robot-preview-image")

    # Define the output directory
    output_dir = "output/Robots"
    os.makedirs(output_dir, exist_ok=True)  # Create the directory if it doesn't exist

    # Create the filename based on the order number
    filename = f"Robot_{order_number}.png"
    filepath = os.path.join(output_dir, filename)
    robot.screenshot(path=filepath)
    return filepath

def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    """Embed the screenshot into the receipt PDF"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path, 
                                   source_path=pdf_path, 
                                   output_path=pdf_path)

def archive_receipts():
    """Archive all receipt PDF files into a single ZIP file"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/Receipts", "./output/receipts.zip")

def clean_up():
    """Cleans up the folders where receipts and screenshots are saved."""
    shutil.rmtree("./output/Receipts")
    shutil.rmtree("./output/Robots")