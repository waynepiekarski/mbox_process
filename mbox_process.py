import mailbox
import os
import csv
import argparse
import sys


def write_output(folder_name, in_file_name, write_data, symlink=False):
    # Remove bad characters from the attachment file name
    file_name = ''.join([s for s in in_file_name if s in 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.,-_ ~`!@#$%^&+=<>'])
    # Some email filters would replace a file extension like .doc with DEFANGED-doc so fix these up
    file_name = file_name.replace("DEFANGED-", ".")

    dupstr = ""
    dup = 0
    while True:
        if os.path.exists(os.path.join(folder_name, dupstr + file_name)):
            dup = dup + 1
            dupstr = str(dup) + "-"
        else:
            break
        if dup == 100:
            sys.exit(f"FATAL: Reached 100 duplicates for [{folder_name}] attachment [{attachment_name}]")
    if symlink:
        os.symlink(os.path.join("..", f"msg-{idx:04}", dupstr + file_name), f"{dir_name}/attachments/msg-{idx:04}-{dupstr}{file_name}")
    # Need to convert any strings to UTF-8 so they can be written in "wb" mode
    if isinstance(write_data, str):
        write_data = write_data.encode('utf-8')
    open(os.path.join(folder_name, dupstr + file_name), "wb").write(write_data)


def safe_decode(payload, charset):
    try:
        body = payload.decode(encoding=charset)
    except UnicodeDecodeError as ude:
        # Some emails are indicated as utf-8 but are actually a different charset, so lets try an alternative to see if it can be made to work
        test_charset = charset.lower().replace("-","")
        if test_charset == "utf8":
            alt_charset = "iso-8859-1"
        # Chinese emails seem to indicate one charset but don't decode, so try an alternative
        elif test_charset == "gb2312" or test_charset == "big5":
            alt_charset = "iso-8859-1"
        # Japanese charsets can fail to decode
        elif test_charset == "iso2022jp":
            alt_charset = "iso-8859-1"
        else:
            sys.exit(f"No alternative charset for {charset}/{test_charset}: Exception decoding {idx:04}: {type(ude).__name__}=[{ude}]")
        try:
            print(f"Msg {idx:04}: Using alternative charset {alt_charset} due to exception using {charset}, specified encoding was probably wrong")
            body = payload.decode(encoding=alt_charset)
        except:
            sys.exit(f"Alternative exception decoding {idx:04}: {type(e).__name__}=[{e}]")
    except Exception as e:
        sys.exit(f"Exception decoding {idx:04}: {type(e).__name__}=[{e}]")
    return body


def safe_charset(part):
    content_charset = part.get_content_charset()
    if content_charset is None:
        # Python needs a charset, so pick a suitable default if none provided
        content_charset = "iso-8859-1"
    elif content_charset == "us-ascii" or content_charset == "ascii":
        # Some emails in us-ascii actually contain non-ascii data, so pick a more useful charset to handle this
        content_charset = "iso-8859-1"
    elif content_charset == "unknown-8bit" or content_charset == "x-unknown":
        # Some emails use "unknown-8bit" but not sure why, lets try UTF-8 and if it fails it will try ISO-8859-1
        content_charset = "utf-8"
    elif content_charset == "utf-8,iso-8859-1":
        # Some emails include two charsets which is not valid, so fix it up
        content_charset = "utf-8"
    elif content_charset == "windows-874":
        # Could be a Thai encoding, which is -11 and not -1
        content_charset = "iso-8859-11"

    return content_charset


############################
### PARSE CLI ARGUEMENTS ###
############################

# Create the parser
my_parser = argparse.ArgumentParser(description='List the content of a folder')

# Add the arguments
my_parser.add_argument('file_name',
                       action='store',
                       help='Name of file to be processed.')

my_parser.add_argument('dir_name',
                       action='store',
                       help='Name of directory to store messages.')

# Execute the parse_args() method
args = my_parser.parse_args()

###################################
### PREPARE FOR FILE PROCESSING ###
###################################

# get filename from user
file_name = args.file_name

# get directory to store messages in from user
dir_name = args.dir_name

# if the output directory exists then we should exit
if os.path.isdir(dir_name):
    sys.exit(f"Output directory {dir_name} already exists")
if dir_name.endswith("/"):
    sys.exit(f"Output directory {dir_name} ends with /, provide only the directory name")

# Create output directory as well as subdirectories for symlink farms
os.mkdir(dir_name)
os.mkdir(os.path.join(dir_name, "attachments"))

# create a list to store information on each message to be written to csv
csv_headers = ['Message', 'From', 'To', 'Date', 'Subject', 'Attachment', 'PNG', 'JPG']
email_list = []

#########################
### PROCESS MBOX FILE ###
#########################

# iterate over messages
count = 0
for idx, message in enumerate(mailbox.mbox(file_name)):
    count = count + 1

    # create a folder for the message
    folder_name = f"{dir_name}/msg-{idx:04}"
    if not os.path.isdir(folder_name):
        # make a folder for this email (named after the subject)
        os.mkdir(folder_name)

    print(f"== Msg {idx:04}: from=[{message['from']}], to=[{message['to']}], subject=[{message['subject']}] date=[{message['date']}]")

    # add message to summary list for csv
    msg_dict_temp = {
        'Message': idx,
        'From': message['from'],
        'To': message['to'],
        'Date': message['date'],
        'Subject': message['subject'],
        'Attachment': "N",
        'PNG': "N",
        'JPG': "N"
    }


    # add header info to full message
    full_message = f'''### ### ### Start Message  ### ### ### \n
TO: {message['to']}
FROM: {message['from']}
DATE: {message['date']}
SUBJECT: {message['subject']}

CONTENT:
    '''

# add html header info to full message
    html_header = f'''### ### ### Start Message  ### ### ### \n </br>
TO: {message['to']}</br>
FROM: {message['from']}</br>
DATE: {message['date']}</br>
SUBJECT: {message['subject']}</br>
</br>
    '''

    # Messages can be either multi-part or single, so make an array so we can use the same code for both cases
    if message.is_multipart():
        message_walk = message.walk()
    else:
        message_walk = [ message ]
    if True:
        # iterate over the message parts
        for part in message.walk():
            # get email content
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition")).replace('\n',' ')
            content_charset = safe_charset(part)
            print(f"Msg {idx:04}: get_content_type={content_type}, Content-Disposition={content_disposition}, get_content_charset={part.get_content_charset()}-->{content_charset}")
            payload = part.get_payload(decode=True)
            # https://docs.python.org/3/library/email.compat32-message.html#email.message.Message.get_payload
            # "If the message is a multipart and the decode flag is True, then None is returned."
            if payload is None:
                # Skip this part because payload.decode() will fail otherwise
                # print(f"Skipping None multipart payload")
                continue

            # save plain text of email
            if content_type == "text/plain":
                # add the body of the message
                body = safe_decode(payload, content_charset)
                full_message_text = full_message + f"{body}"

                # write the message to a file
                filename = f"msg-{idx:04}.txt"
                # Previously was writing {body} if the file existed already
                write_output(folder_name, filename, full_message_text)

            # save html text of email
            elif content_type == "text/html":
                # add the body of the message
                body = safe_decode(payload, content_charset)
                full_message_html = html_header + f"CONTENT: \n{body}"

                # write the message to a file
                filename = f"msg-{idx:04}.html"
                write_output(folder_name, filename, full_message_html)

            # multipart/mixed messages
            elif content_type == "multipart/mixed":
                # add the body of the message
                body = safe_decode(payload, content_charset)
                full_message_mixed = full_message + f"CONTENT: \n{body}"

                # write the message to a file
                filename = f"msg-{idx:04}-mxd.html"
                write_output(folder_name, filename, full_message_mixed)

            # save png attachments
            elif content_type == "image/png":
                # update email summary list
                msg_dict_temp['PNG'] = "Y"
                attachment_name = part.get_filename()
                if attachment_name:
                    write_output(folder_name, attachment_name, part.get_payload(decode=True), symlink=True)

            # save jpg attachments
            elif content_type == "image/jpeg":
                # update email summary list
                msg_dict_temp['JPG'] = "Y"
                attachment_name = part.get_filename()
                if attachment_name:
                    write_output(folder_name, attachment_name, part.get_payload(decode=True), symlink=True)

            # save email attachment
            elif "attachment" in content_disposition:
                # update email summary list
                msg_dict_temp['Attachment'] = "Y"
                attachment_name = part.get_filename()
                if attachment_name:
                    write_output(folder_name, attachment_name, part.get_payload(decode=True), symlink=True)

            # save any other attachments that have a filename
            elif part.get_filename() != None and part.get_filename() != "":
                write_output(folder_name, part.get_filename(), part.get_payload(decode=True), symlink=True)

            else:
                print(f"Ignoring item which had no match in if statements, attachment has no filename")

        # append final message dict to email summary list
        email_list.append(msg_dict_temp)


with open(f'{dir_name}/summary.csv', 'w') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=csv_headers)
    writer.writeheader()
    writer.writerows(email_list)


if count == 0:
    sys.exit(f"Failed to find any emails in mailbox {file_name}")
