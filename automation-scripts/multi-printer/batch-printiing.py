import win32api
import win32print
import os
import time

# 🔹 CHANGE THIS PATH
FOLDER_PATH = r"C:\Users\kevin\scipy\Sylvia Alfaro"


def print_pdf_batch(folder_path, wait_time=5):
    """
    Print all PDF files in a folder
    
    Args:
        folder_path: Path to folder containing PDFs
        wait_time: Seconds to wait between prints (default: 5)
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
    print("=" * 60)
    print(f"📁 Folder: {folder_path}")
    print(f"📄 Total PDF files: {len(pdf_files)}")
    print("=" * 60)
    
    # List all files to be printed
    for i, filename in enumerate(pdf_files, 1):
        print(f"  {i}. {filename}")
    
    # Get default printer
    try:
        printer_name = win32print.GetDefaultPrinter()
        print(f"\n🖨️  Using printer: {printer_name}")
    except Exception as e:
        print(f"❌ Error: Cannot get default printer - {e}")
        return
    
    # Confirm before printing
    print("\n" + "=" * 60)
    response = input("Press ENTER to start printing (or type 'cancel' to stop): ")
    if response.lower() == 'cancel':
        print("❌ Printing cancelled")
        return
    
    print("\n🚀 Starting batch print...\n")
    
    # Track results
    success_count = 0
    fail_count = 0
    failed_files = []
    
    # Print each file
    for i, filename in enumerate(pdf_files, 1):
        full_path = os.path.join(folder_path, filename)
        
        # Normalize path (handles spaces and special chars)
        full_path = os.path.normpath(full_path)
        
        print(f"[{i}/{len(pdf_files)}] {filename}")
        
        # Check if file exists and is accessible
        if not os.path.exists(full_path):
            print(f"  ❌ File not found!")
            fail_count += 1
            failed_files.append((filename, "File not found"))
            continue
        
        if not os.path.isfile(full_path):
            print(f"  ❌ Not a valid file!")
            fail_count += 1
            failed_files.append((filename, "Not a valid file"))
            continue
        
        # Attempt to print
        try:
            win32api.ShellExecute(
                0,
                "print",
                full_path,
                f'/d:"{printer_name}"',
                ".",
                0
            )
            print(f"  ✅ Sent to printer")
            success_count += 1
            
            # Wait before next print
            if i < len(pdf_files):  # Don't wait after last file
                print(f"  ⏳ Waiting {wait_time} seconds...")
                time.sleep(wait_time)
            
        except Exception as e:
            print(f"  ❌ Error: {e}")
            fail_count += 1
            failed_files.append((filename, str(e)))
        
        print()
    
    # Final summary
    print("=" * 60)
    print("📊 BATCH PRINT SUMMARY")
    print("=" * 60)
    print(f"✅ Successful: {success_count}")
    print(f"❌ Failed: {fail_count}")
    print(f"📄 Total: {len(pdf_files)}")
    
    # Show failed files if any
    if failed_files:
        print("\n❌ Failed Files:")
        for filename, error in failed_files:
            print(f"  • {filename}")
            print(f"    Reason: {error}")
    
    print("=" * 60)
    print("✅ Batch printing completed!")
    print("\n💡 Tip: Check your printer queue to monitor progress")


# Run the batch print
if __name__ == "__main__":
    # You can adjust the wait time (in seconds) between prints
    print_pdf_batch(FOLDER_PATH, wait_time=5)