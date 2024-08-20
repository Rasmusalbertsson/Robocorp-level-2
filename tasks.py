from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import zipfile

@task
def order_robots_from_RobotSpareBin():
    """
    Main task to order robots from RobotSpareBin Industries Inc.
    - Opens the website and closes any modals.
    - Fills out the form to order robots.
    - Saves receipts as PDFs and screenshots of the ordered robots.
    - Embeds the screenshots into the PDF receipts.
    - Archives all the receipts and images into a ZIP file.
    """
    open_robot_order_website()
    close_annoying_modal()
    fill_the_form()
    archive_receipts()

def open_robot_order_website():
    """Navigates to the robot order page."""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Downloads the orders.csv file and returns the result."""
    http = HTTP()
    result = http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    return result

def close_annoying_modal():
    """Closes any annoying modal that might appear on the website."""
    page = browser.page()
    page.click("text=Yep")

def fill_the_form():
    """
    Fills the robot order form for each entry in the CSV.
    - Selects options and fills in the required fields.
    - Takes screenshots and stores PDF receipts.
    - Embeds screenshots into receipts and handles multiple orders.
    """
    page = browser.page()
    orders = get_orders()
    library = Tables()
    order = library.read_table_from_csv(
        "orders.csv", columns=['Order number', 'Head', 'Body', 'Legs', 'Address']
    )
    
    for row in order:
        # Fill in the form fields with data from the CSV
        page.select_option("#head.custom-select", row['Head'])
        page.check(f"input[name='body'][value='{row['Body']}']")
        page.fill("input[placeholder='Enter the part number for the legs']", str(row['Legs']))
        page.fill("input[placeholder='Shipping address']", str(row['Address']))
        page.click("id=preview")
        
        # Take a screenshot of the robot preview
        screenshot_robot(row["Order number"], page)

        # Try to click the ORDER button until it disappears
        while True: 
            order_button = page.locator("#order")
            if order_button.is_visible():
                order_button.click()
            else:
                print("ORDER button not found, moving to the next step.")
                break

        # Handle the receipt and embed the screenshot
        try: 
            pdf = store_receipt_as_pdf(row['Order number'], page)
            print(type(pdf))
            page.click("text=ORDER ANOTHER ROBOT")
            print("Clicked on 'ORDER ANOTHER ROBOT'")
            page.click("text=Yep")
            print("Clicked on 'Yep'")
            embed_screenshot_to_receipt(
                screenshot=f"output/robot_preview_{row['Order number']}.png",
                pdf_file=f"output/order_receipt_{row['Order number']}.pdf"
            )
        except Exception:
            print("No error detected, order successful")
            pdf = store_receipt_as_pdf(row['Order number'], page)
            print(type(pdf))
            page.click("text=ORDER ANOTHER ROBOT")
            page.click("text=Yep")
            embed_screenshot_to_receipt(
                screenshot=f"output/robot_preview_{row['Order number']}.png",
                pdf_file=f"output/order_receipt_{row['Order number']}.pdf"
            )
            break

def store_receipt_as_pdf(order_number, page):
    """
    Converts the HTML receipt on the page into a PDF.
    - Captures the receipt HTML.
    - Saves the receipt as a PDF file.
    """
    receipt_html = page.locator("id=receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, f"output/order_receipt_{order_number}.pdf")

def screenshot_robot(order_number, page):
    """Takes a screenshot of the robot preview and saves it as a PNG file."""
    page.locator("id=robot-preview-image").screenshot(path=f"output/robot_preview_{order_number}.png")

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """
    Embeds the screenshot into the PDF receipt.
    - Takes a screenshot and the corresponding PDF file.
    - Appends the screenshot to the PDF.
    """
    pdf = PDF()
    list_of_files = [screenshot, pdf_file]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=pdf_file, append=True
    )

def archive_receipts():
    """
    Archives all the order receipt PDFs into a single ZIP file.
    - Collects the list of PDF files generated for each order.
    - Creates a ZIP file containing all the receipts.
    """
    filename = "output/order_receipt_"
    listofiles = []
    library = Tables()
    
    # Read the CSV and get order numbers
    order = library.read_table_from_csv(
        "orders.csv", columns=['Order number', 'Head', 'Body', 'Legs', 'Address']
    )
    length = len(order)
    
    # Collect the list of files to archive
    for i in range(length):
        i = i + 1  # Adjusting index to match order numbers
        file_path = f"{filename}{i}.pdf"
        listofiles.append(file_path)

    # Create a ZIP file using Python's zipfile module
    with zipfile.ZipFile('output/orders.zip', 'w') as zipf:
        for file in listofiles:
            zipf.write(file, arcname=file.split('/')[-1])
    
    print("Zip archive created successfully!")




