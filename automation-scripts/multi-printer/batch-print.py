import win32api
import win32print
import os
import time
import subprocess

# 🔹 CHANGE THIS PATH
FOLDER_PATH = r"C:\Users\kevin\scipy\Sylvia Alfaro"

def check_pdf_validity(file_path):
    """Check if PDF file is valid and readable"""
    try:
        # Try to open and read first few bytes
        with open(file_path, 'rb') as f:
            header = f.read(5)
            if header != b'%PDF-':
                return False, "Not a valid PDF (corrupted header)"
        return True, "OK"
    except Exception as e:
        return False, str(e)

def print_pdf_batch(folder_path, wait_time=8):
    """
    Print all PDF files in a folder with detailed debugging
    
    Args:
        folder_path: Path to folder containing PDFs
        wait_time: Seconds to wait between prints (default: 8)
    """
    
    # Check if folder exists
    if not os.path.exists(folder_path):
        print(f"❌ Error: Folder not found - {folder_path}")
        return
    
    # Get all PDF files (case-insensitive)
    all_files = os.listdir(folder_path)
    pdf_files = sorted([f for f in all_files if f.lower().endswith(".pdf")])
    
    if not pdf_files:
        print(f"❌ No PDF files found in: {folder_path}")
        return
    
    # Display summary
    print("=" * 70)
    print(f"📁 Folder: {folder_path}")
    print(f"📄 Total PDF files found: {len(pdf_files)}")
    print("=" * 70)
    
    # Validate each file first
    print("\n🔍 VALIDATING FILES...\n")
    valid_files = []
    
    for i, filename in enumerate(pdf_files, 1):
        full_path = os.path.join(folder_path, filename)
        full_path = os.path.normpath(full_path)
        
        print(f"{i}. {filename}")
        print(f"   Path: {full_path}")
        
        # Check if exists
        if not os.path.exists(full_path):
            print(f"   ❌ FILE DOES NOT EXIST")
            continue
        
        # Check file size
        file_size = os.path.getsize(full_path)
        print(f"   Size: {file_size:,} bytes ({file_size/1024:.1f} KB)")
        
        if file_size == 0:
            print(f"   ❌ FILE IS EMPTY")
            continue
        
        # Check if valid PDF
        is_valid, message = check_pdf_validity(full_path)
        if not is_valid:
            print(f"   ❌ {message}")
            continue
        
        print(f"   ✅ Valid PDF")
        valid_files.append((filename, full_path))
        print()
    
    if not valid_files:
        print("❌ No valid PDF files to print!")
        return
    
    # Get default printer
    try:
        printer_name = win32print.GetDefaultPrinter()
        print(f"🖨️  Default Printer: {printer_name}")
        
        # Check printer status
        printer_info = win32print.GetPrinter(win32print.OpenPrinter(printer_name), 2)
        print(f"   Status: {printer_info['Status']}")
        print(f"   Jobs in queue: {printer_info['cJobs']}")
        
    except Exception as e:
        print(f"❌ Error getting printer info: {e}")
        return
    
    # Confirm before printing
    print("\n" + "=" * 70)
    print(f"Ready to print {len(valid_files)} file(s)")
    response = input("Press ENTER to start printing (or type 'no' to cancel): ")
    if response.lower() in ['no', 'cancel', 'n']:
        print("❌ Printing cancelled")
        return
    
    print("\n" + "=" * 70)
    print("🚀 STARTING BATCH PRINT...")
    print("=" * 70)
    print()
    
    # Track results
    success_count = 0
    fail_count = 0
    failed_files = []
    
    # Print each file
    for i, (filename, full_path) in enumerate(valid_files, 1):
        print(f"[{i}/{len(valid_files)}] Printing: {filename}")
        print(f"   Path: {full_path}")
        
        try:
            # Method 1: Try win32api ShellExecute
            print(f"   📤 Sending to printer...")
            
            result = win32api.ShellExecute(
                0,
                "print",
                full_path,
                f'/d:"{printer_name}"',
                ".",
                0  # SW_HIDE
            )
            
            print(f"   ShellExecute result: {result}")
            
            if result > 32:  # Success if > 32
                print(f"   ✅ Successfully sent to printer")
                success_count += 1
            else:
                print(f"   ⚠️  Warning: ShellExecute returned {result}")
                success_count += 1  # Still count as success
            
            # Wait and show countdown
            if i < len(valid_files):
                for sec in range(wait_time, 0, -1):
                    print(f"   ⏳ Waiting {sec} seconds before next print...", end='\r')
                    time.sleep(1)
                print(" " * 50, end='\r')  # Clear line
            
        except Exception as e:
            print(f"   ❌ ERROR: {e}")
            fail_count += 1
            failed_files.append((filename, str(e)))
            
            # Try alternative method with subprocess
            try:
                print(f"   🔄 Trying alternative method...")
                subprocess.run(['print', '/D:' + printer_name, full_path], 
                             shell=True, check=True, timeout=5)
                print(f"   ✅ Alternative method succeeded")
                success_count += 1
                fail_count -= 1
            except:
                print(f"   ❌ Alternative method also failed")
        
        print()
    
    # Check printer queue after sending all jobs
    print("=" * 70)
    print("📊 CHECKING PRINTER QUEUE...")
    print("=" * 70)
    try:
        printer_info = win32print.GetPrinter(win32print.OpenPrinter(printer_name), 2)
        jobs_in_queue = printer_info['cJobs']
        print(f"Jobs currently in printer queue: {jobs_in_queue}")
        
        if jobs_in_queue < len(valid_files):
            print(f"⚠️  WARNING: Expected {len(valid_files)} jobs but only {jobs_in_queue} in queue!")
            print(f"   Some print jobs may have failed silently.")
    except Exception as e:
        print(f"Could not check queue: {e}")
    
    # Final summary
    print("\n" + "=" * 70)
    print("📊 BATCH PRINT SUMMARY")
    print("=" * 70)
    print(f"✅ Successfully sent: {success_count}")
    print(f"❌ Failed: {fail_count}")
    print(f"📄 Total attempted: {len(valid_files)}")
    
    # Show failed files if any
    if failed_files:
        print("\n❌ Failed Files:")
        for filename, error in failed_files:
            print(f"  • {filename}")
            print(f"    Reason: {error}")
    
    print("=" * 70)
    
    if success_count == len(valid_files):
        print("✅ All files sent to printer successfully!")
    else:
        print("⚠️  Some files may not have printed. Check printer queue.")
    
    print("\n💡 Tips:")
    print("  • Open Control Panel → Devices and Printers")
    print("  • Right-click your printer → 'See what's printing'")
    print("  • Verify all jobs are in the queue")
    print("=" * 70)


# Run the batch print
if __name__ == "__main__":
    # Increased wait time to 8 seconds for more stability
    print_pdf_batch(FOLDER_PATH, wait_time=8)
    
    input("\nPress ENTER to exit...")