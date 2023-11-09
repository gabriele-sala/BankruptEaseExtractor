import openpyxl
import requests
from pdfminer.high_level import extract_text

def extract_data_from_pdf(pdf_path):
    # Extract text from the PDF
    text = extract_text(pdf_path)
    
    # Split the text into lines
    lines = text.split("\n")
    
    # Find the start of the desired section
    start_index = lines.index("Part 11, Question 77: Other property of any kind not already listed")
    
    # Extract the relevant lines
    relevant_lines = lines[start_index:]
    
    # Process lines to handle multi-line cells
    processed_lines = []
    temp_line = ""
    for i, line in enumerate(relevant_lines):
        if "Crypto Assets:" in line:
            if temp_line:  # If there's content in temp_line, append it to processed_lines
                processed_lines.append(temp_line)
            temp_line = line.strip()  # Start a new entry
        else:
            temp_line += " " + line.strip()  # Concatenate lines

        # If the line contains a $ sign or it's the last line, finalize the entry
        if "$" in line or i == len(relevant_lines) - 1:
            processed_lines.append(temp_line)
            temp_line = ""

    return processed_lines

def get_current_prices(tickers):
    url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest"
    headers = {
        "Accepts": "application/json",
        "X-CMC_PRO_API_KEY": "APY KEY",  # Replace with your API key
    }

    response = requests.get(url, headers=headers)
    data = response.json()

    prices = {}
    for crypto in data["data"]:
        symbol = crypto["symbol"]
        if symbol in tickers:
            prices[symbol] = crypto["quote"]["USD"]["price"]

    return prices

def write_to_excel(data, excel_path):
    # Fetch current prices
    all_tickers = [process_line(line)[0] for line in data if process_line(line)[0]]
    current_prices = get_current_prices(all_tickers)

    wb = openpyxl.Workbook()
    
    # Function to write data to a worksheet with current prices and current value
    def write_data_to_sheet(ws, tickers=None):
        ws.append(["Crypto Asset", "Quantity", "USD Spot Price Value", "Current Price", "Current Value"])
        for line in data:
            asset, qty, value = process_line(line)
            if asset and (not tickers or asset in tickers):
                current_price = current_prices.get(asset, "N/A")
                try:
                    current_value = float(qty.replace(',', '')) * float(current_price)
                except ValueError:
                    current_value = "N/A"
                ws.append([asset, qty, value, current_price, current_value])

    # Write main data to the first sheet
    ws_main = wb.active
    ws_main.title = "Main Data"
    write_data_to_sheet(ws_main)

    # Write specific tickers to separate sheets
    tickers_sets = [
        ["BTC", "ETH", "BNB", "XRP", "ADA", "DOGE", "SOL", "TRX", "TON", "DOT", "MATIC", "LTC", "WBTC", "AVAX"],
        ["USDC", "USDT", "FIAT", "BUSD", "TUSD", "FRAX", "USDP", "GUSD", "EUROC"]
    ]
    
    for tickers in tickers_sets:
        ws = wb.create_sheet(title=tickers[0])  # Name the sheet after the first ticker for simplicity
        write_data_to_sheet(ws, tickers)

    wb.save(excel_path)


def process_line(line):
    parts = line.split(";")
    if len(parts) != 3:
        return None, None, None

    asset_parts = parts[0].split(":")
    qty_parts = parts[1].split(":")
    value_parts = parts[2].split(":")

    if len(asset_parts) < 2 or len(qty_parts) < 2 or len(value_parts) < 2:
        return None, None, None

    asset = asset_parts[1].strip()
    qty = qty_parts[1].strip()
    value = value_parts[1].strip().replace("$", "")  # Remove the $ symbol

    return asset, qty, value

if __name__ == "__main__":
    pdf_path = "path to file" # insert path to file
    excel_path = "path to output file" # name of the output xlsx file
    
    data = extract_data_from_pdf(pdf_path)
    write_to_excel(data, excel_path)
