from robocorp.tasks import task
from robocorp import browser, http
from RPA import Assistant, Archive
from RPA.Tables import Tables
from RPA.PDF import PDF
import os

@task
def minimal_task():
    """Order robots from RobotSpareBin Industries Inc"""
    open_robot_order_website()
    download_csv_file()
    get_dataset_from_csv()
    add_to_zip("output/receipts", "order_receipts.zip")
    success_dialog("Finish!")

def success_dialog(message):
    win_dialog = Assistant.Assistant()
    win_dialog.add_text(message)
    win_dialog.ask_user(timeout=180, height=300, width=300)

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.configure_context()
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads csv file from the given URL"""
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_dataset_from_csv():
    """Get dataset from CSV"""
    whole_table = Tables()
    csv_dataset = whole_table.read_table_from_csv(path= "orders.csv", header= True)

    for row in csv_dataset:
        fill_webpage(row)

def close_popup():
    '''Close popup'''
    page = browser.page()
    page.click("button:text('OK')")

def fill_webpage(orders):
    '''Fill webpage with values retrived from csv'''
    close_popup()

    page = browser.page()
    
    #Head
    page.select_option("#head",str(orders["Head"]))

    #Body
    selector = "#id-body-" + str(orders["Body"])
    page.check("#id-body-2")

    #Legs
    ### TEST PURPOSE ONLY - START ###
    #page.fill("xpath=html/body/div/div/div[1]/div/div[1]/form/div[3]/input","2")
    #page.fill("xpath=html/body/div/div/div[1]/div/div[1]/form/div[3] >> css=.form-control","2")
    #page.fill("//*/input[@placeholder='Enter the part number for the legs']","2")
    #page.fill(".form-control[placeholder='Enter the part number for the legs']","2")
    #page.fill("input[class=form-control]","2")
    ### TEST PURPOSE ONLY - END ###
    page.fill(".form-control",str(orders["Legs"]))

    #Address
    page.fill("#address",str(orders["Address"]))

    #Preview
    page.click("#preview")

    #Submit
    is_submit_visible = True
    num_retry = 0

    while is_submit_visible == True and num_retry <= 10:
        page.click("#order")
        is_submit_visible = page.is_visible("#order", timeout=1000)
        num_retry += 1

    path_pdf = export_as_pdf(str(orders["Order number"]))
    robot_image_path = get_robot_image()
    append_to_pdf(robot_image_path, path_pdf)

    #Order another robot
    page.click("#order-another")

def export_as_pdf(order_number):
    """Export order receipt to a pdf file"""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    path = "output/receipts/" + str(order_number) + "__order_receipt.pdf"
    pdf.html_to_pdf(order_receipt_html, path)

    return path

def get_robot_image():
    '''Get robot image'''
    page = browser.page()
    robot_image_path = "screenshot.png"
    page.locator("#robot-preview-image").screenshot(path=robot_image_path)

    return robot_image_path

def append_to_pdf(path_element_to_add, path_pdf):
    '''Append html element to pdf'''
    pdf = PDF()
    list_of_files = [
        path_element_to_add
    ]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=path_pdf,
        append=True
    )

def add_to_zip(folder_to_archive, archive_name):
    '''Add files to zip'''
    zip_init = Archive.Archive()
    zip_init.archive_folder_with_zip(folder= folder_to_archive, archive_name= archive_name)


'''
def minimal_task():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=100
    )
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()
    success_dialog()

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()
    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")
    
def success_dialog():
    win_dialog = Assistant.Assistant()
    win_dialog.add_text("Click on a button to close popup")
    win_dialog.ask_user(timeout=180)

def download_excel_file():
    """Downloads excel file from the given URL"""
    http.download("https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    workbook = excel.open_workbook("SalesData.xlsx")
    worksheet = workbook.worksheet("data").as_table(header=True)

    for row in worksheet:
        fill_and_submit_sales_form(row)

def collect_results():
    """Take a screenshot of the page"""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def log_out():
    """Presses the 'Log out' button"""
    page = browser.page()  
    page.click("text=Log out")

def export_as_pdf():
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")
    
'''