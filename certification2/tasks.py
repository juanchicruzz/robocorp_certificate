from robocorp.tasks import task
from robocorp import browser, http, vault
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem
import pandas as pd
import time




@task
def place_all_orders():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    init_folders()
    browser.configure(slowmo=100)
    download_excel_file()
    db = get_excel_data()
    open_the_intranet_website()
    for index,row in db.iterrows():
        fill_one_order(row["Head"],row["Body"],row["Legs"],row["Address"])
        if is_order_successful():
            get_results(str(row["Order number"]))
            another_robot()
        else:
            open_the_intranet_website()
    archive_recipes()
    clear_folders()
    

def init_folders():
    '''Init the folders.'''
    fs = FileSystem()
    fs.create_directory("output/receipts",exist_ok=True)

def clear_folders():
    '''Clear the folders.'''
    fs = FileSystem()
    fs.remove_directory("output/receipts",recursive=True)
    fs.remove_file("output/orders.csv")
    fs.remove_file("output/robot.png")

def open_the_intranet_website():
    '''Navegate to the given URL.'''
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_modal()


def close_modal():
    page = browser.page()
    try:
        page.wait_for_selector("//div[@role='dialog']", timeout=10000)
        page.click("//button[normalize-space()='OK']")
    except:
        pass
     
def another_robot():
    '''Click the another robot button.'''
    page = browser.page()
    page.click("//button[@id='order-another']")
    close_modal()

def fill_one_order(head,body,legs,address):
    '''Fill the sales data for the week.'''
    page = browser.page()
    page.select_option("//select[@id='head']",str(head))
    page.check("//input[@id='id-body-"+str(body)+"']")
    page.fill("//input[@placeholder='Enter the part number for the legs']",str(legs))
    page.fill("//input[@id='address']",str(address))
    page.click("//button[@id='order']")

def is_order_successful():
    '''Check if the order is successful.'''
    page = browser.page()
    try:
        page.wait_for_selector("//div[@id='receipt']", timeout=5000)
        return True
    except:
        page.wait_for_selector("//div[@class='alert alert-danger']", timeout=10000)
        return False

def get_results(order_number):
    '''Get the results.'''
    page = browser.page()
    time.sleep(0.5)
    page.query_selector("//div[@id='robot-preview']").screenshot(path="output/robot.png",)
    receipt = page.inner_html("//div[@id='receipt']")
    pdf = PDF()
    pathPdf = "output/receipts/receipt_order_"+str(order_number)+".pdf"
    pdf.html_to_pdf(receipt, pathPdf)
    imgList = ["output/robot.png:align=center"]
    pdf.add_files_to_pdf(imgList, pathPdf,append=True)


def archive_recipes():
    ar = Archive()
    ar.archive_folder_with_zip(folder="output/receipts", archive_name="output/receipts.zip") 


def download_excel_file():
    '''Download the excel file.'''
    http.download("https://robotsparebinindustries.com/orders.csv", "output/orders.csv", overwrite=True)

def get_excel_data():
    '''Get the excel data.'''
    db = pd.read_csv("output/orders.csv", delimiter=",")
    
    return db
    