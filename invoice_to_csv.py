# importing required classes 
from pypdf import PdfReader 
import sys
from functools import reduce
from datetime import datetime

CSV_HEADER = "Item Name,Category,Location,Purchase Date,Purchase Cost,Model Name,Model Number,Serial Number,Order Number,Requestable,Last Audit\n"
LINE_BREAK = "\n--------------------------------------------------------------------------------\n"
COMMON_CATEGORIES = [
    "Laptop",
    "Desktop",
    "Access Point",
    "Smartphone",
    "Router",
    "Switch"
]

class Asset:
    """
    Represents one type of asset in the invoice.

    If there are multiple of the same asset, they are all stored in one object

    Fields:
        model_no: Model Number
        name: Asset Name
        model: Model
        price: Price of one of the assets
        serial_nos: List of the serial numbers
        date: Date Ordered
        order_no: Order Number
        category: Asset Category
    """
    def __init__(self, model_no, name, model, price, serial_nos, date, order_no, category):
        self.model_no = model_no
        self.name = name
        self.model = model
        self.price = price
        self.serial_nos = serial_nos
        self.location = "Naples Office"
        self.date = date
        self.order_no = order_no
        self.category = category
        self.requestable = "true"
        self.last_audit_date = self.date
    
    def to_string(self):
        """
        Outputs each of the assets as rows for a csv file.
        """
        # Item name, Category, Location, Purchase Date, Purchase Cost, Model Name, Model No, Serial No, Order no
        assets_str = ""
        for serial_no in self.serial_nos:
            asset_str = self.name + "," + self.category + "," + self.location + "," + self.date
            asset_str = asset_str + "," +  self.price + "," + self.model + "," + self.model_no
            asset_str = asset_str + "," + serial_no + "," + self.order_no + "," + self.requestable
            asset_str = asset_str + "," + self.last_audit_date + "\n"
            assets_str = assets_str + asset_str
        return assets_str

def get_pdf_names():
    """
    Gets the PDF names to process.

    If there are command line arguments, treats those as the pdf filenames. Otherwise asks the user to input filenames.

    Returns:
        pdf_names: List of all valid pdf names
        invalid_names: List of all names that don't have a .pdf file extention.
    """
    if len(sys.argv) > 1:
        pdf_names, invalid_names = check_file_types(sys.argv[1:])
    else:
        input_names = input("Input the name (or names, seperated by spaces) of an invoice pdf: ").split(" ")
        pdf_names, invalid_names = check_file_types(input_names)
    return pdf_names, invalid_names


def check_file_types(file_names):
    """
    Helper method to check if the file names have the proper .pdf extension.

    Parameters:
        file_names: a list of the filenames to check.
    Returns:
        pdf_names: List of all valid pdf names
        invalid_names: List of all names that don't have a .pdf file extention.
    """
    pdf_names = []
    invalid_names = []
    for name in file_names:
        if name[-4:] == ".pdf":
            pdf_names.append(name)
        else:
            invalid_names.append(name)
    return pdf_names, invalid_names

def parse_pdf(page_text):
    """
    Parses a pdf and gets the data ready to be output into a csv.

    Parses the data based on the standard invoice format of Softrim.

    Parameters:
        page_text: The text of the pdf page to parse.
                   Note: This text should be formatted by pypdf's layout extraction mode. This ensures that there is
                   whitespace in all of the places where whitespace appears on the pdf, and this whitespace is used 
                   to parse the data.
    Returns: 
        An instance of the Asset object with all of the necessary data to form a csv.
    """
    START_STRING = "Billable  Products & Other Charges"
    END_STRING = "Total"
    start_i = page_text.find(START_STRING) + len(START_STRING)
    end_i = page_text.find(END_STRING)
    # Just the item info from the invoice
    item = page_text[start_i:end_i].strip()
    model_no = item.split(": ")[0]

    info = item.split(": ")[1]
    # Remove Serial Number(s): from the end
    info = info[:info.find("Serial")]
    # Split on large spaces
    info_list = list(filter(lambda x: x != "", info.split("  ")))
    # Split on new lines
    info_list = list(map(str.splitlines, info_list))
    # Expand inner list into elements of outer list
    info_list = reduce(lambda x, y: x + y, info_list)
    # Remove quantity 
    info_list.pop(1)
    # Get value of price and remove from list
    price = info_list.pop(1)
    # Remove $ and commas from price
    price = price.replace(",", "")
    price = price.replace("$", "")
    price = price.strip()
    # Remove total price
    info_list.pop(1)
    model = info_list[0].split(", ")[0]
    name = " ".join(info_list).replace(",", "")

    # Split the serial numbers into a list and strips them
    serial_nos = list(map(str.strip, item.split(": ")[-1].split(",")))

    # Get the order date and changes it to the correct format
    date = get_next_item(page_text, "BCB Homes")
    date = date.replace("/", "-")
    # Change from MM-DD-YYYY to YYYY-MM-DD
    date = date[6:] + date[5] + date[:5]

    # Get the order number
    order_no = get_next_item(page_text, "Service Request Number")

    # Request the user for the asset category
    category = get_category(model_no, model, name)

    return Asset(model_no, name, model, price, serial_nos, date, order_no, category)

def get_category(model_no, name, model):
    """
    Requests the user to specify the asset category for a particular asset.

    Common asset categories include: Laptop, Desktop, Smartphone, Monitor, etc.

    Parameters;
        model_no: Model Number of the asset
        name: Asset name
        model: Asset model
    Returns:
        The model category capitalized
    """
    category = check_for_common_categories(name)
    if category == None:
        print("\nModel Number:", model_no)
        print("Model:", model)
        print("Name:", name)
        category = input("What is the category for this item? ").capitalize()
    return category

def check_for_common_categories(name):
    """
    Checks the name of the asset to determine if it includes a common category name.

    Returns:
        The category name found, or None if no category is found.
    """
    found_category = None
    for category in COMMON_CATEGORIES:
        if name.lower().find(category.lower()) >= 0:
            found_category = category 
    return found_category


def get_next_item(text, word_before):
    """
    Gets the next single alphanumeric or symbol string in the file after a given word, spliting on spaces and new line. 

    Usage Example:
        To find the Order Number, word_before would be "Service Request Number"
        Text:
        ...
        Service Request Number             531363 
        ...

    Parameters:
        text: The text of the pdf page
        word_before: The word before the field you are trying to retrieve.
    Returns: 
        The next alphanumeric word or symbols after the word before
    """
    return text[text.find(word_before) + len(word_before):].strip().split(" ")[0].split("\n")[0]

def create_csv(csv_name, header):
    try:
        csv = open(csv_name, "x")
        print(csv_name, " created")
        csv.write(header)
        return csv 
    except FileExistsError:
        print("\nThe file", csv_name, "already exists.")
        selection = input("Would you like to overwrite it? (y/n) ")
        if selection.lower() == "y" or selection.lower() == "yes":
            csv = open(csv_name, "w")
            csv.write(header)
            print(csv_name, " overwritten")
            return csv
        else:
            exit()
    except OSError:
        print("Invalid filename. Contact Cade (404-219-5157) for assistance.") 
        exit()
    

def add_to_csv(csv, data):
    csv.write(data)

def process_pdfs(pdf_names):
    """
    Takes a list of pdf filenames and attempts to process and produce a csv for each one.

    Parameters:
        pdf_names: A list of pdf filenames to process
    """
    print("Files to process:")
    [print("\t" + name) for name in pdf_names]

    csv_name = str(datetime.now()).replace(" ", "@").replace(":", ";") + ".csv"
    csv = create_csv(csv_name, CSV_HEADER) 

    print(LINE_BREAK)

    for pdf_name in pdf_names:
        print("Processing", pdf_name)
        try:
            reader = PdfReader(pdf_name)
            page_text = reader.pages[0].extract_text(extraction_mode="layout")
            add_to_csv(csv, parse_pdf(page_text).to_string())
            print("Finished processing", pdf_name, "\n", LINE_BREAK)

        except FileNotFoundError: 
            print(pdf_name, "not found.\n")

def check_for_flags():
    """
    Checks the arguments for flags and handles them.

    Flags:
        -h | --help: Prints the usage message
    """
    flags = []
    if len(sys.argv) > 1:
        for i in range(1, len(sys.argv)):
            if sys.argv[i][0] == "-":
                flags.append(sys.argv.pop(i))
    for flag in flags:
        if flag == "-h" or flag == "--help":
            print_help()
            exit()

def print_help():
    print("Usage: python invoice_to_csv.py [optional: invoice filenames, space seperated]")

def main():
    """
    Main function. 
    """
    check_for_flags()
    pdf_names, invalid_names = get_pdf_names()
    process_pdfs(pdf_names)
    if len(invalid_names) > 0:
        print("\nInvalid filenames:")
        [print("\t" + name) for name in invalid_names]

# Run main
if __name__ == "__main__":
    main()
