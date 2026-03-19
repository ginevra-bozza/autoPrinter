import os
import time
import win32api
import smtplib
from email.message import EmailMessage
from imap_tools import MailBox, AND
from PIL import Image
import img2pdf
import img2pdf
from dotenv import load_dotenv

# --- CONFIGURATION ---
load_dotenv()
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
APP_PASSWORD = os.getenv("APP_PASSWORD")

if not EMAIL_ADDRESS or not APP_PASSWORD:
    print("CRITICAL ERROR: Missing .env file or credentials! Please make sure the .env file is in the same folder.")
    time.sleep(10)
    exit()

# SECURITY: Only allow these safe file types. Everything else is ignored.
ALLOWED_EXTENSIONS = ('.pdf', '.png', '.jpg', '.jpeg', '.txt', '.docx', '.doc')

# Create a temporary folder on the PC to hold files while processing
TEMP_DIR = os.path.join(os.environ['USERPROFILE'], "Documents", "PrintSpool_Temp")
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)


def send_confirmation_reply(to_email, ignored_files=None):
    """Sends an automatic success email back to the sender."""
    try:
        msg = EmailMessage()

        body = "Documento inviato con successo\n"
        if ignored_files:
            body += f"\nNota: I seguenti allegati sono stati ignorati per motivi di sicurezza {', '.join(ignored_files)}"

        msg.set_content(body)
        msg['Subject'] = 'Re: Stampa ricevuta'
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = to_email

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, APP_PASSWORD)
            server.send_message(msg)

        print(f"Confirmation email successfully sent to {to_email}")
    except Exception as e:
        print(f"Failed to send confirmation reply: {e}")


def process_and_merge_to_a4(image_paths, output_pdf):
    """Converts images to RGB, maps them to A4 size, and merges to one PDF."""
    print(f"Formatting {len(image_paths)} image(s) to A4 paper...")
    processed_paths = []

    for img_path in image_paths:
        with Image.open(img_path) as img:
            img = img.convert('RGB')
            temp_path = img_path + "_temp.jpg"
            img.save(temp_path, "JPEG", quality=95)
            processed_paths.append(temp_path)

    a4_layout = img2pdf.get_layout_fun((img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297)))

    with open(output_pdf, "wb") as f:
        f.write(img2pdf.convert(processed_paths, layout_fun=a4_layout))

    for p in processed_paths:
        if os.path.exists(p): os.remove(p)

    return output_pdf


def monitor_email():
    print(f"Printer Script Running. Listening to {EMAIL_ADDRESS}...")

    while True:
        try:
            with MailBox('imap.gmail.com').login(EMAIL_ADDRESS, APP_PASSWORD) as mailbox:
                for msg in mailbox.fetch(AND(seen=False)):
                    print(f"New email received from: {msg.from_}")

                    images_to_print = []
                    direct_print_files = []
                    files_to_cleanup = []
                    ignored_files = []
                    printed_something = False

                    # 1. Download, Filter, and Sort attachments
                    for att in msg.attachments:
                        # SECURITY CHECK
                        if not att.filename.lower().endswith(ALLOWED_EXTENSIONS):
                            print(f"SECURITY ALERT: Ignored unsafe file -> {att.filename}")
                            ignored_files.append(att.filename)
                            continue  # Skip to the next attachment

                        filepath = os.path.join(TEMP_DIR, att.filename)
                        with open(filepath, 'wb') as f:
                            f.write(att.payload)

                        if att.filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                            images_to_print.append(filepath)
                        else:
                            direct_print_files.append(filepath)

                    # 2. Handle Images (Merge & Print)
                    if images_to_print:
                        final_pdf = os.path.join(TEMP_DIR, f"PrintJob_{int(time.time())}.pdf")
                        process_and_merge_to_a4(images_to_print, final_pdf)

                        print("Sending fitted images to printer...")
                        win32api.ShellExecute(0, "print", final_pdf, None, ".", 0)
                        printed_something = True

                        files_to_cleanup.extend(images_to_print)
                        files_to_cleanup.append(final_pdf)

                    # 3. Handle Documents (Print Directly)
                    for doc in direct_print_files:
                        print(f"Sending document directly to printer: {os.path.basename(doc)}")
                        win32api.ShellExecute(0, "print", doc, None, ".", 0)
                        printed_something = True
                        files_to_cleanup.append(doc)

                    # 4. Confirm and Clean Up
                    if printed_something or ignored_files:
                        send_confirmation_reply(msg.from_, ignored_files)

                        time.sleep(10)
                        for file in files_to_cleanup:
                            if os.path.exists(file):
                                try:
                                    os.remove(file)
                                except Exception as e:
                                    pass

        except Exception as e:
            print(f"Connection check failed, retrying in 30s. Error: {e}")

        time.sleep(30)


if __name__ == "__main__":
    monitor_email()