from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import os
from datetime import datetime

def read_phone_numbers_from_excel(file_path):
    """
    Read phone numbers from Excel file
    """
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        
        # Look for phone number column (case-insensitive)
        phone_columns = [col for col in df.columns if 'phone' in col.lower() or 'number' in col.lower()]
        
        if phone_columns:
            # Use the first matching column
            phone_numbers = df[phone_columns[0]].dropna().astype(str).tolist()
        else:
            # If no phone column found, use the first column
            phone_numbers = df.iloc[:, 0].dropna().astype(str).tolist()
        
        # Clean phone numbers
        cleaned_numbers = []
        for num in phone_numbers:
            # Remove any whitespace
            num = str(num).strip()
            # Skip empty strings
            if num and num != 'nan':
                cleaned_numbers.append(num)
        
        return cleaned_numbers
    
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def send_whatsapp_bulk(excel_file_path, message):
    """
    Send WhatsApp messages to all numbers in Excel file with 1 minute delay
    """
    # Read phone numbers from Excel
    phone_numbers = read_phone_numbers_from_excel(excel_file_path)
    
    if not phone_numbers:
        print("❌ No phone numbers found in the Excel file!")
        return
    
    print(f"\n✅ Found {len(phone_numbers)} phone numbers in Excel file")
    print("Phone numbers to message:")
    for i, num in enumerate(phone_numbers[:5], 1):  # Show first 5
        print(f"  {i}. {num}")
    if len(phone_numbers) > 5:
        print(f"  ... and {len(phone_numbers) - 5} more")
    
    # Initialize Edge driver
    edge_options = EdgeOptions()
    edge_options.add_argument("--disable-blink-features=AutomationControlled")
    
    driver = webdriver.Edge(options=edge_options)
    
    # Track results
    successful_sends = 0
    failed_sends = []
    
    try:
        # Open WhatsApp Web
        print("\nOpening WhatsApp Web in Edge...")
        driver.get("https://web.whatsapp.com")
        driver.maximize_window()
        
        print("\n" + "="*70)
        print("WHATSAPP WEB LOGIN")
        print("="*70)
        
        # Wait for login
        logged_in = False
        max_wait_time = 300
        check_interval = 5
        elapsed_time = 0
        
        print(f"You have {max_wait_time} seconds (5 minutes) to scan the QR code")
        print("Steps:")
        print("1. Open WhatsApp on your phone")
        print("2. Go to Settings > Linked Devices")
        print("3. Tap 'Link a Device'")
        print("4. Scan the QR code on the screen")
        print("\n" + "-"*70)
        
        while elapsed_time < max_wait_time and not logged_in:
            try:
                # Check if logged in
                driver.find_element(By.XPATH, '//div[@id="side"]')
                logged_in = True
                print("\n✓ Successfully logged in to WhatsApp Web!")
                break
            except:
                remaining_time = max_wait_time - elapsed_time
                print(f"\rWaiting for QR scan... {remaining_time} seconds remaining", end='', flush=True)
                time.sleep(check_interval)
                elapsed_time += check_interval
        
        if not logged_in:
            print("\n✗ Login timeout")
            return
        
        # Wait for WhatsApp to fully load
        print("\nWaiting for WhatsApp to fully load...")
        time.sleep(5)
        
        # Send messages to each number
        print("\n" + "="*70)
        print("SENDING MESSAGES")
        print("="*70)
        
        total_numbers = len(phone_numbers)
        
        for index, phone_number in enumerate(phone_numbers, 1):
            print(f"\n[{index}/{total_numbers}] Processing: {phone_number}")
            
            try:
                # Clean phone number
                phone_number_clean = phone_number.replace('+', '').replace(' ', '').replace('-', '')
                
                # Navigate to the phone number
                print(f"   Opening chat...")
                url = f"https://web.whatsapp.com/send?phone={phone_number_clean}&text="
                driver.get(url)
                
                # Wait for page to load
                time.sleep(8)
                
                # Handle "Continue to Chat" button
                try:
                    continue_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//div[@role="button"]//span[contains(text(), "Continue to chat")]'))
                    )
                    continue_button.click()
                    print("   ✓ Clicked 'Continue to chat'")
                    time.sleep(3)
                except:
                    pass
                
                # Try to send message
                message_sent = False
                
                # Method 1: Footer method
                try:
                    message_box = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//footer//div[@contenteditable="true"]'))
                    )
                    message_box.click()
                    time.sleep(1)
                    message_box.send_keys(message)
                    time.sleep(1)
                    message_box.send_keys(Keys.ENTER)
                    message_sent = True
                except:
                    pass
                
                # Method 2: Action chains
                if not message_sent:
                    try:
                        actions = webdriver.ActionChains(driver)
                        actions.send_keys(message)
                        actions.send_keys(Keys.ENTER)
                        actions.perform()
                        message_sent = True
                    except:
                        pass
                
                if message_sent:
                    print(f"   ✓ Message sent to {phone_number}")
                    successful_sends += 1
                else:
                    print(f"   ✗ Failed to send to {phone_number}")
                    failed_sends.append(phone_number)
                
                # Wait 1 minute before next message (except for the last one)
                if index < total_numbers:
                    print("\n⏰ Waiting 1 minute before next message...")
                    for i in range(60, 0, -10):
                        print(f"   {i} seconds remaining...", end='\r')
                        time.sleep(10)
                    print("   Ready for next message!        ")
                
            except Exception as e:
                print(f"   ✗ Error with {phone_number}: {str(e)[:50]}")
                failed_sends.append(phone_number)
                continue
        
        # Final summary
        print("\n" + "="*70)
        print("SUMMARY")
        print("="*70)
        print(f"Total numbers: {total_numbers}")
        print(f"✓ Successful: {successful_sends}")
        print(f"✗ Failed: {len(failed_sends)}")
        
        if failed_sends:
            print("\nFailed numbers:")
            for num in failed_sends:
                print(f"  - {num}")
        
        # Save results to log file
        save_results_log(phone_numbers, successful_sends, failed_sends)
        
    except Exception as e:
        print(f"\n✗ Critical error: {str(e)}")
    
    finally:
        input("\nPress Enter to close Edge...")
        driver.quit()

def save_results_log(all_numbers, successful, failed_list):
    """
    Save sending results to a log file
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = f'whatsapp_sending_log_{timestamp}.txt'
    
    with open(log_file, 'w', encoding='utf-8') as f:
        f.write("WhatsApp Message Sending Log\n")
        f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("="*50 + "\n\n")
        
        f.write(f"Total numbers: {len(all_numbers)}\n")
        f.write(f"Successful: {successful}\n")
        f.write(f"Failed: {len(failed_list)}\n\n")
        
        f.write("Failed Numbers:\n")
        for num in failed_list:
            f.write(f"  - {num}\n")
    
    print(f"\n📄 Log file saved: {log_file}")

def get_excel_file():
    """
    Get Excel file path from user
    """
    while True:
        print("\n📁 Enter the Excel file path:")
        print("   (You can drag and drop the file here)")
        file_path = input("   File path: ").strip().strip('"').strip("'")
        
        if os.path.exists(file_path):
            if file_path.endswith(('.xlsx', '.xls')):
                return file_path
            else:
                print("❌ File must be an Excel file (.xlsx or .xls)")
        else:
            print("❌ File not found! Please check the path.")

def main():
    """
    Main function
    """
    print("="*70)
    print("📱 WhatsApp Bulk Message Sender")
    print("="*70)
    print("This tool will:")
    print("1. Read phone numbers from an Excel file")
    print("2. Send a message to each number")
    print("3. Wait 1 minute between messages")
    print("="*70)
    
    # Get Excel file
    excel_file = get_excel_file()
    
    # Get message
    print("\n📝 Enter the message to send:")
    message = input("   Message: ").strip()
    
    if not message:
        print("❌ Message cannot be empty!")
        return
    
    # Confirm before starting
    phone_numbers = read_phone_numbers_from_excel(excel_file)
    if phone_numbers:
        print(f"\n⚠️  This will send messages to {len(phone_numbers)} numbers")
        print(f"⏱️  Estimated time: ~{len(phone_numbers)} minutes")
        
        confirm = input("\n🚀 Start sending? (yes/no): ").strip().lower()
        if confirm in ['yes', 'y']:
            send_whatsapp_bulk(excel_file, message)
        else:
            print("❌ Cancelled.")
    else:
        print("❌ No valid phone numbers found in the Excel file!")

if __name__ == "__main__":
    main()