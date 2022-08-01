import imaplib
import email
import os
import sys
import getpass
import time
import argparse
from colorama import init, Fore, Back, Style

def main():
    parser = argparse.ArgumentParser(description='Run executable attachments of new incoming mails')
    parser.add_argument('-H', '--host', required=True)
    parser.add_argument('-P', '--port', required=True)
    parser.add_argument('-U', '--user', required=True)
    args = parser.parse_args()
    host = args.host
    port = args.port
    user = args.user
    password = getpass.getpass("Password: ")

    init()
    try:
        mail = imaplib.IMAP4_SSL(host, port)
        mail.login(user, password)
        mail.select('inbox')
        print(Fore.GREEN + "[+] Successful connection" + Style.RESET_ALL)
    except:
        print(Fore.RED + "[-] Connection failed" + Style.RESET_ALL)
        sys.exit()

    extensions = ['.exe', '.ps1', '.sh']
    attachments_dir = 'attachments'

    while True:
        try:
            result, data = mail.uid('search', None, "ALL")
            latest_email_uid = data[0].split()[-1]
            result, data = mail.uid('fetch', latest_email_uid, '(RFC822)')
            raw_email = data[0][1]
            email_message = email.message_from_bytes(raw_email)
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                filename = part.get_filename()
                if not filename:
                    continue
                for ext in extensions:
                    if filename.endswith(ext):
                        print(Fore.GREEN + "[+] File found: " + filename + Style.RESET_ALL)
                        if not os.path.exists(attachments_dir):
                            os.makedirs(attachments_dir)
                        path = attachments_dir + "/" + time.strftime("%d-%m-%Y-%H-%M-%S") + "_" + filename
                        with open(path, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        if os.name == 'posix':
                            os.system("chmod +x {}".format(path))
                        os.system(path)
                        mail.uid('store', latest_email_uid, '+FLAGS', '\\Deleted')
                        mail.expunge()
                        break
        except Exception as e:
            print(Fore.RED + "[-] Error: " + str(e) + Style.RESET_ALL)
            sys.exit()
        time.sleep(5)

if __name__ == "__main__":
    main()