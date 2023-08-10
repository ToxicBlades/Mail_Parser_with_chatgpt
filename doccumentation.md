# Mail_Parser_with_chatgpt
The main objective of my project was to efficiently manage the diverse range of offers received through our email platform. Initially, I considered creating numerous cases to handle each type of offer, but I realized that this approach would be time-consuming and inefficient. Therefore, I decided to integrate the assistance of Chat_GPT to simplify and enhance the analysis process. 

## Python code
Here i will tell everything you need to know about each functions i used in my programm

***
### Db connection
```
db_connection = mysql.connector.connect(
    host='localhost',
    user='root',
    password='',
    database= DBNAME
) 
```
Here we use 'mysql.connector' which allows us to create connection with our database, host here is IP address of server on which
stays Database, user and password its your credentials for database, database = name of your database 

***
### cursor object

 ` cursor = db_connection.cursor() `
this code snippet is used to establish a connection to a database and obtain a cursor object, which allows you to execute SQL queries and fetch results from the database.

***
### pattern

` pattern = r"\s-\s(?![\w\d]+-[а-яА-ЯёЁ]+) `

\s matches any whitespace character.
'-' matches the hyphen character.
\s matches another whitespace character.
(?![\w\d]+-[а-яА-ЯёЁ]+) is a negative lookahead assertion that specifies a pattern not to be matched. It checks if the following characters after the hyphen are not a combination of word characters (letters, digits, underscores) followed by a hyphen and Cyrillic characters (letters or ёЁ). 

***
### logging

` logging.basicConfig(filename='email_processing.log', level=logging.DEBUG) `

set up basic configuration for logging in Python. It configures the logging module to write log messages to a file named "email_processing.log" and sets the logging level to DEBUG, which is the lowest level. \
level=logging.DEBUG: This sets the logging level to DEBUG, which is the lowest level. It means that all log messages, including those with DEBUG, INFO, WARNING, ERROR, and CRITICAL levels, will be written to the log file. 

***
### Api key

` openai.api_key = APIKEY `
here we just set our apikey 

***
### Output direction

` OUTPUT_DIR = DIR `
here we create variable which will be with path to our folder for downloading attachemnets(for reading them)

***
### Remove text and words function

``` 
    def remove_text_and_words(input_string):
    words_to_remove = ["Forwarded", "message","messages", "Date", "Subject", "To", "Cc", "www","Adress","Mail","email","e-mail","Cell","Skype","Whatsapp","Warehouse","Manager","https","Copyright","Director","Web","Mobile","Phone","Mob","Ustr","Address","Tel","web-chat","T:","F:","regards","http","Trade"]
    russian_pattern = re.compile("[а-яА-ЯёЁ]+")  # find Russian words

    output_string = input_string
    for word in words_to_remove:
        pattern_n = r"(?i)\b" + word + r"\b.*?[\r\n]"  # case-insensitive match for the whole word until newline or carriage return
        output_string = re.sub(pattern_n, "", output_string)

    output_string = re.sub(russian_pattern, "", output_string)  # delete Russian words
    output_string = output_string.replace('\r', '')  # remove carriage return characters

    return output_string 
```

In this code we create list with words which we indicate as useless info for AI,creating russian pattern because all russian text 
is not needed also. After that The code iterates over each word in the list words_to_remove.
It creates a regular expression pattern pattern_n to match the word as a whole word using the \b boundary markers. The (?i) at the beginning makes the pattern case-insensitive. The .*?[\r\n] part matches any characters until a newline or carriage return is encountered.
It uses the re.sub() function to replace all occurrences of pattern_n in the output_string with an empty string, effectively removing those words and everything after them until a newline or carriage return.
It uses a russian_pattern (presumably a regular expression pattern) to delete Russian words from the output_string.
It uses the replace() function to remove carriage return characters (\r) from the output_string.

***
### Connect to gmail

``` 
    def connect_to_gmail(username, password):

    M = imaplib.IMAP4_SSL('imap.gmail.com')
    M.login(username, password)
    M.select('inbox')
    return M 
```


The function connect_to_gmail takes two parameters: username and password, which represent the Gmail account credentials.
It creates an instance of the IMAP4_SSL class from the imaplib module and assigns it to the variable M. The IMAP4_SSL class establishes a secure SSL connection to the IMAP server.
It uses the login method of the IMAP4_SSL instance (M) to authenticate with the Gmail server using the provided username and password.
It selects the 'inbox' mailbox using the select method of the IMAP4_SSL instance. This means that subsequent operations will be performed on the 'inbox' mailbox.
Finally, the function returns the IMAP4_SSL instance (M), allowing the caller to interact with the Gmail server using the returned connection object.

***
### fetch mails

``` 
    def fetch_all_emails(M):
    """Fetches all emails from the inbox"""
    #Change UNSEEN TO ALL for first time if needed
    result, data = M.search(None, 'UNSEEN')
    email_ids = data[0].split()
    return email_ids
```

    The function fetch_all_emails takes a single parameter M, which represents the connection object to the IMAP server.
It uses the search method of the IMAP4_SSL instance (M) to search for emails in the selected mailbox. The search criteria are specified as 'UNSEEN', which retrieves all unread emails. If you want to fetch all emails regardless of their read/unread status, you can change 'UNSEEN' to 'ALL'.
The search method returns a result and a list of email IDs (data) that match the search criteria.
The email IDs are extracted from data[0] using the split method, which splits the string into a list of individual email IDs.
Finally, the function returns the list of email IDs.

***
### send mails
``` 
    def send_email(subject, sender, date, to_email):
    """Send email using a Gmail account"""
    
    from_email = EMAIL  # your email here
    from_password = PASSWORD  

    try:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = "Этот офер нужно обработать вручную" # you need to add data from this message by your self

        body = f"Subject: {subject}\nSender: {sender}\nDate: {date}"
        msg.attach(MIMEText(body, 'plain'))
        time.sleep(20) 

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()

    except Exception as e:
        logging.error("This email havent been send for some reason ") 
```

The function send_email takes four parameters: subject, sender, date, and to_email, which represent the email subject, sender information, email date, and the recipient's email address, respectively.
The from_email variable is set to the value of EMAIL, which presumably holds the Gmail account email address.
The from_password variable is set to the value of PASSWORD, which presumably holds the password for the Gmail account.
The code enters a try block to handle potential exceptions that may occur during the email sending process.
An instance of MIMEMultipart is created to represent the email message.
The From, To, and Subject headers of the email are set using the respective properties of the msg object.
The body of the email is constructed using the provided subject, sender, and date values, and attached to the msg object using MIMEText.
A 20-second delay is added using time.sleep(20). This is presumably to avoid spamming if there are a lot of messages being sent in rapid succession.
An SMTP server connection is established with Gmail's SMTP server using smtplib.SMTP and the appropriate host and port.
The connection is secured using starttls().
The Gmail account is authenticated by calling login on the server object with the from_email and from_password.
The email message is converted to a string using msg.as_string().
The sendmail method of the server object is called to send the email, specifying the from_email, to_email, and the email text.
Finally, the SMTP server connection is closed using server.quit(). 

***
### Extract email fields

``` 
    def extract_email_fields(mail):
    """Extracts relevant fields from the email"""
    subject = str(mail.subject)
    sender = ''.join(map(str, mail.from_))
    date = str(mail.date)
    text = ''.join(map(str, mail.text_plain))
    text_html = ''
    #this thing exist for cases when for some reason we cant recive plain text from mail, so we get info with html
    if len(text) < 5:
        code_html = str(mail.text_html)
        soup = BeautifulSoup(code_html, 'html.parser')
        text_html = soup.get_text(separator=' ').strip()
    return subject, sender, date, text, text_html 
```

The function extract_email_fields(mail) takes an email object as input and extracts relevant fields from the email. Here's a breakdown of what the function does:
It initializes variables for subject, sender, date, text, and text_html.
It converts the mail.subject to a string and assigns it to the subject variable.
It converts the mail.from_ (sender) to a string and assigns it to the sender variable. The mail.from_ field is usually a list, so the map function is used to convert each item in the list to a string, and then the join function concatenates the strings.
It converts the mail.date to a string and assigns it to the date variable.
It converts the mail.text_plain (plain text body) to a string and assigns it to the text variable. The mail.text_plain field is also usually a list, so the map and join functions are used similar to the sender variable.
It checks if the length of text is less than 5 characters. If it is, it means that the plain text body was not properly received, so it tries to extract the text from the HTML body.
If the length of text is less than 5, it assigns the mail.text_html (HTML body) to the code_html variable as a string.
It uses the BeautifulSoup library to parse the HTML code in code_html and extract the text from it.
It assigns the extracted text from the HTML body to the text_html variable.
Finally, it returns the subject, sender, date, text, and text_html as a tuple. 

***
### def procces attachments


``` 
    def process_attachments(mail, attached_data):
    """Processes attachments in the email"""
    for attachment in mail.attachments:
        filename = attachment['filename']
        extension = os.path.splitext(filename)[1].lower()

        if extension == '.xlsx':
            file_data = attachment['payload']
            file_content = base64.b64decode(file_data)
            temp_filename = os.path.join(OUTPUT_DIR, filename)
            with open(temp_filename, 'wb') as temp_file:
                temp_file.write(file_content)
            xls_data = openpyxl.load_workbook(temp_filename)
            for sheet_name in xls_data.sheetnames:
                xls_sheet = xls_data[sheet_name]
                for row in xls_sheet.iter_rows(values_only=True):
                    row_data = '|'.join(map(str, row))
                    attached_data += row_data + '\n'
            os.remove(temp_filename)

        elif extension == '.xls':
            file_data = attachment['payload']
            file_content = base64.b64decode(file_data)
            temp_filename = os.path.join(OUTPUT_DIR, filename)
            with open(temp_filename, 'wb') as temp_file:
                temp_file.write(file_content)
            xls_data = xlrd.open_workbook(temp_filename)
            for sheet_name in xls_data.sheet_names():
                xls_sheet = xls_data.sheet_by_name(sheet_name)
                for row_index in range(xls_sheet.nrows):
                    row_data = '|'.join(map(str, xls_sheet.row_values(row_index)))
                    attached_data += row_data + '\n'
            os.remove(temp_filename)

        elif extension == '.pdf':
            file_data = attachment['payload']
            file_content = base64.b64decode(file_data)
            temp_filename = os.path.join(OUTPUT_DIR, filename)
            with open(temp_filename, 'wb') as temp_file:
                temp_file.write(file_content)
            
            pdf_text = extract_text(temp_filename)
            pdf_text = pdf_text.replace("\f","")
            attached_data += pdf_text
            os.remove(temp_filename)
        
        elif extension == '.docx':
            file_data = attachment['payload']
            file_content = base64.b64decode(file_data)
            temp_filename = os.path.join(OUTPUT_DIR, filename)
            with open(temp_filename, 'wb') as temp_file:
                temp_file.write(file_content)

            docx_text = ''
            doc = Document(temp_filename)
            for para in doc.paragraphs:
                docx_text += para.text + '\n'
            attached_data += docx_text

            os.remove(temp_filename)
        
        elif extension == '.csv':
            file_data = attachment['payload']
            file_content = base64.b64decode(file_data)
            temp_filename = os.path.join(OUTPUT_DIR, filename)
            with open(temp_filename, 'wb') as temp_file:
                temp_file.write(file_content)
            with open(temp_filename, 'r') as csv_file:
                csv_reader = csv.reader(csv_file)
                for row in csv_reader:
                    row_data = '|'.join(map(str, row))
                    attached_data += row_data + '\n'
            os.remove(temp_filename)

    return attached_data 
```

    The function takes two parameters: mail, which represents the email object, and attached_data, which is a string that accumulates the data from the attachments.

The code iterates through the attachments in the email using a for loop. For each attachment, it extracts the filename and determines the file extension using os.path.splitext and converts it to lowercase. Based on the file extension, different processing steps are performed:

If the extension is .xlsx, the attachment's payload is decoded from base64 and saved to a temporary file. The openpyxl library is used to load the workbook from the temporary file. Then, for each sheet in the workbook, the data from each row is extracted and concatenated with '|' as a separator. The resulting row data is appended to the attached_data string. Finally, the temporary file is removed.

If the extension is .xls, a similar process as above is followed, but the xlrd library is used to load the workbook and extract the data from each sheet.

If the extension is .pdf, the attachment's payload is decoded from base64 and saved to a temporary file. The extract_text function is called to extract the text from the PDF file. Any newline characters created by form feeds ("\f") are removed, and the resulting text is appended to the attached_data string. The temporary file is then removed.

If the extension is .docx, the attachment's payload is decoded from base64 and saved to a temporary file. The python-docx library is used to read the document and extract the text from each paragraph. The extracted text is appended to the attached_data string. The temporary file is then removed.

If the extension is .csv, the attachment's payload is decoded from base64 and saved to a temporary file. The file is opened in read mode, and the csv module is used to read the file. For each row in the CSV file, the values are joined with '|' as a separator, and the resulting row data is appended to the attached_data string. The temporary file is then removed.

Finally, the function returns the accumulated attached_data string, which should contain the extracted data from all the attachments in the email. 

***
### Parse emails

``` 
    def parse_emails(username, password):
    """Main function to parse emails"""
    try:
        M = connect_to_gmail(username, password)
        email_ids = fetch_all_emails(M)

        email_data = []
        for email_id in email_ids:
            try:
                email_data.append(process_email(M, email_id))
            except Exception as e:
                logging.error(f"Failed to process email {email_id}: {str(e)}")
    except Exception as e:
        logging.error(f"Failed to connect or fetch emails: {str(e)}")
        return

    process_ai(email_data)

    M.close()
    M.logout() 
```

Function called parse_emails that serves as the main function to parse emails. It takes two parameters: username and password
which represent the credentials to connect to a Gmail account.

Inside the function, there is a try-except block. The code attempts to connect to the Gmail account using the connect_to_gmail function and fetches all the email IDs using the fetch_all_emails function . These functions are assumed to exist and handle the respective operations.

If the connection and email fetching are successful, an empty list called email_data is created. Then, a for loop iterates over each email ID in the email_ids list. For each email ID, the code attempts to process the email using the process_email function . If an exception occurs during the email processing, an error message is logged, indicating the failed email ID and the exception details.

If an exception occurs during the connection or email fetching process, an error message is logged, indicating the exception details. The function then returns without proceeding further.

After processing all the emails, the email_data list is passed to a function called process_ai

***
### process mail
``` 
    def process_email(M, email_id):
    """Function to process a single email"""
    result, data = M.fetch(email_id, '(RFC822)')
    mail = mailparser.parse_from_bytes(data[0][1])

    subject, sender, date, text, text_html = extract_email_fields(mail)

    attached_data = process_attachments(mail, '')

    if text_html:
        return [subject, sender, date, text, text_html, attached_data]
    else:
        return [subject, sender, date, text, '', attached_data] 
```

The function takes two parameters: M and email_id. M is assumed to be an instance of an email client (e.g., an IMAP client) that provides the fetch method to retrieve email data. email_id represents the unique identifier of the email to be processed.

The M.fetch(email_id, '(RFC822)') line fetches the email with the specified email_id using the fetch method of the email client M. The argument '(RFC822)' indicates that the email should be fetched in RFC822 format, which is the standard format for email messages.

The fetched email data is then parsed using the mailparser.parse_from_bytes(data[0][1]) function. The data variable is a tuple returned by the fetch method, and data[0][1] represents the email data itself. The mailparser library is used to parse the email data and create a mail object.

The extract_email_fields function is called to extract various fields from the parsed email object mail. These fields include the subject, sender, date, plain text body (text), and HTML body (text_html). The extract_email_fields function is not shown in the provided code snippet and may be defined elsewhere.

The process_attachments function is called to process any attachments in the email. The mail object and an empty string ('') are passed as arguments. The purpose and implementation of the process_attachments function are not shown in the provided code and may be defined elsewhere.

Depending on the presence of text_html, the function returns a list containing the subject, sender, date, plain text body, HTML body, and the result of processing attachments (attached_data). If text_html is not available, an empty string ('') is used in place of the HTML body.

Overall, the process_email function fetches an email, parses its content, extracts relevant fields, processes attachments, and returns the extracted information as a list.

***
### process ai

` ai_responses = []  # List to save AI responses `


``` 
    def process_ai(email_data):
    """Function to perform AI processing on email data"""
    for data in email_data:
        subject = data[0]
        text = data[3]
        if "Fwd:" in subject:
            sender = ""
            date = ""
            subject = subject.replace("Fwd:","") #delete useless info
            lines = text.split("\n")  # Split the email text into lines

            for line in lines:
                if line.startswith("From:") or line.startswith("От:"):
                    sender = line.replace("From:", "").replace("От:", "").strip()
                    sender = sender.replace("<","").replace(">","")
                elif line.startswith("Date:"):
                    date = line.replace("Date:", "").strip()
        else: 
            sender = data[1]
            sender = sender.replace("<","").replace(">","")
            date = data[2]

        text_html = data[4]
        attached_data = data[5]
        response_data = ''

        if attached_data:
            if text_html:
                combined_text = remove_text_and_words(text_html + "\n" + attached_data)
            else:
                combined_text = remove_text_and_words(text + "\n" + attached_data)
        else:
            combined_text = remove_text_and_words(text)

        retries = 0
        while retries < 3:
            try:
                completion = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages = [{"role": "user", "content": f" I need your help to analyze the following text. Can you read through the content and provide me with the following details in a structured format: Name of the Products, Weight/Volume, type of packaging,their respective Prices,type of price(per case/box/case), and Incoterms. The expected output format is as follows: 'Product - Name of the product, Weight/volume - weigtht/volume of the product,Type - type of packaging, Price - Price of the product,Type of price - its per case/box/pcs,Incoterms - respective Incoterm'. Please extract the required information from the text below:  {combined_text} "}],
                max_tokens = 500,
                temperature = 0.8)
                # Extract the required information from the completion
                response_data = extract_info_from_ai_completion(completion)
                time.sleep(20)
                if response_data:
                    #leave while because we got data 
                    break

                retries += 1 #dont ask me why this exist,but without it it doesnt works for some reason, base of programing be like :P

                # Pause for 20 seconds (Api has limit for requsts in a minute,if we procces more then 3 mail it give us api error)
            except Exception:
                retries += 1
                time.sleep(20)

        if response_data:
            # Add the response to the list
            for response in response_data:
                combined_list = response + [sender,subject,date]
                ai_responses.append(combined_list)

        else:
            send_email(subject, sender, date, EMAIL)

    save_ai_responses()  # Save AI responses after all emails are processed 
```

It takes in email_data as input, which is a list of email data.
It iterates over each item in email_data and extracts the subject, text, sender, date, text_html, and attached_data from each email.
If the subject of the email contains "Fwd:", it removes the "Fwd:" prefix from the subject and extracts the sender and date from the email lines.
If the email doesn't contain "Fwd:", it extracts the sender and date directly from the email data.
It checks if there is attached data in the email. If so, it combines the text_html and attached_data or text and attached_data, and removes any unwanted text and words.
If there is no attached data, it removes unwanted text and words from the text.
The function then attempts to make a chat completion API call using OpenAI's GPT-3.5-turbo model. It sends a message to the model, asking for help in analyzing the provided text and extracting specific details related to products.
The response from the AI model is processed to extract the required information.
If the response_data is not empty, it appends the extracted information, along with the sender, subject, and date, to a list called ai_responses.
If the response_data is empty, it calls a function send_email to send an email with the subject, sender, date, and the EMAIL variable.
Finally, the function calls another function save_ai_responses to save the AI responses after processing all the emails.

***
### Extract info from Ai

``` 
    def extract_info_from_ai_completion(completion):
    """Function to extract the required information from the AI completion"""
    response = completion.choices[0].message['content']
    logging.debug('Chat answer is %s', response)
    #this thing needed if response contains only 1 answer(means that only 1 product) we dont need to separate message
    word = "Product"
    repeated = check_word_repetition(response,word)
    last_price = None  # Variable to hold the last price seen

    if repeated:
        #some answers sperated with \n some without so we need handle both cases
        if "\n\n" in response:
            products = response.split("\n\n")
        else:
            products = response.split("\n")
        extracted_info = []
        for product in products:
        # Split the product info by space-dash-space
            product_info = re.split(  pattern, product)
            try:
                product_name = product_info[1].strip()  # The product name is the 1 element
                product_name = product_name.replace("Weight/Volume","").replace("Weight/volume","")

                Weigth = product_info[2].strip() # product weigth is 2 element
                Weigth = Weigth.replace("Type","")

                pack_type = product_info[3].strip() #product pack type is 3 elemnt
                pack_type = pack_type.replace("Price","")

                price = product_info[4].strip()  # The product price is the 4 element
                price = price.replace("Type of price","") #delete useless info
                price = remove_last_comma(price)
                
                price_type = product_info[5].strip() # price_type is the 5 element
                price_type = price_type.replace("Incoterms","")


                if not any(char.isdigit() for char in price):
                    if last_price is not None:
                        price = last_price
                    else:
                        logging.error("No price available for product. Skipping this product.")
                        last_price = "0"
                        continue
                else:
                    last_price = price
                


                incoterm = product_info[6].strip()  # The incoterm is the 6 element

                
                extracted_info.append([product_name, Weigth,pack_type, price, price_type, incoterm])
            except IndexError:
                logging.error("Error while extracting information from ai response. Skipping this product.")
                continue

    else:
        products = response
        # Extract the required information for each product
        extracted_info = []
        product_info = products.split(" - ")
        try:
            product_name = product_info[1].strip()  # The product name is the 1 element
            product_name = product_name.replace("Weight/Volume","").replace("Weight/volume","").replace(",","")

            Weigth = product_info[2].strip() # product weigth is 2 element
            Weigth = Weigth.replace("Type","").replace(",","")

            pack_type = product_info[3].strip() # product pack_type is 3 element
            pack_type = pack_type.replace("Price","").replace(",","")

            price = product_info[4].strip()  # The product price is the 4 element
            price = price.replace("Type of price","") #delete useless info
            price = remove_last_comma(price)
                
            price_type = product_info[5].strip() # price_type is the 5 element
            price_type = price_type.replace("Incoterms","").replace(",","")

            incoterm = product_info[6].strip()  # The incoterm is the 6 element


            extracted_info.append([product_name, Weigth,pack_type, price,price_type, incoterm])
        except IndexError:
            logging.error("Error while extracting information from product. Skipping this product.")
    
    return extracted_info 
```

    It takes the completion object as input, which represents the response received from the AI model.
It extracts the content of the response from the completion object and assigns it to the response variable.
If the word "Product" appears multiple times in the response, it indicates that there are multiple product entries. The function checks if there is repetition of the word "Product" using the check_word_repetition function.
If there is repetition, it splits the response into individual product entries. The splitting is done based on the presence of double newlines (\n\n) or single newlines (\n).
It initializes an empty list called extracted_info to store the extracted information.
It iterates over each product entry and splits it into different components using the pattern defined by the pattern variable. The pattern seems to be a regular expression pattern used for splitting the product information.
It attempts to extract the product name, weight/volume, package type, price, price type, and incoterm from each product entry.
It performs some string manipulation to remove unwanted text and characters from the extracted information.
If a product entry is missing any required information (e.g., IndexError occurs during extraction), it logs an error and skips that product entry.
The extracted information for each product is appended as a list to the extracted_info list.
If there is no repetition of the word "Product" in the response, it assumes there is only one product entry.
It splits the response based on the " - " separator and attempts to extract the required information for that single product entry.
Similar to the previous case, it performs string manipulation and error handling for the extraction process.
The extracted information for the single product is appended to the extracted_info list.
Finally, the function returns the extracted_info list containing the extracted information for all products.

***
### save ai responses


``` 
    def save_ai_responses():
    """Saves AI responses to the database"""
    sql = "INSERT INTO offers (product, weight, pack_type, price,price_type, incoterm, sender, subject, date) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"
    for data in ai_responses:
        product, weight, pack_type, price, price_type, incoterm,sender,subject,date = data
        values = (product,weight, pack_type, price, price_type, incoterm,sender,subject,date)
        cursor.execute(sql, values)

    db_connection.commit()

    print("AI responses saved to the database") 
```

    The function begins by defining an SQL query as a string. The query is an INSERT statement that inserts data into a table named "offers" with columns named "product," "weight," "pack_type," "price," "price_type," "incoterm," "sender," "subject," and "date."
The function iterates over each item in the ai_responses list, which contains the AI responses and associated information.
For each item, it unpacks the values from the item into variables: product, weight, pack_type, price, price_type, incoterm, sender, subject, date.
It creates a tuple called values that contains the unpacked values.
The SQL query is executed using the cursor.execute() method, passing the query string and the values tuple as parameters.
The changes made to the database are committed using the db_connection.commit() method.
Finally, a message "AI responses saved to the database" is printed to indicate that the saving process is complete

***
### count word occurrences

``` 
    def count_word_occurrences(text, word):
    """" Count how much words Product in string"""
    words = text.split()
    count = 0
    for w in words:
        if w.lower() == word.lower():
            count += 1
    return count 
```

    
The code you provided includes a function called count_word_occurrences that counts the number of occurrences of a specific word in a given text. Here's how the function works:

The function takes two parameters: text and word. text represents the string in which the word occurrences need to be counted, and word represents the word to be counted.
It splits the text into individual words using the split() method, which splits the string based on whitespace characters (spaces, tabs, newlines, etc.).
It initializes a variable count to keep track of the number of occurrences.
It iterates over each word in the words list.
For each word, it compares it with the word parameter in a case-insensitive manner by converting both words to lowercase using the lower() method.
If the comparison is true, indicating a match, it increments the count by 1.
After iterating through all the words in the words list, it returns the final count.

***
### check word repetition

``` 
    def check_word_repetition(text, word):
    """ Check if words Prouct more then 2"""
    count = count_word_occurrences(text, word)
    return count >= 2 
```

The function takes two parameters: text and word. text represents the string in which the word repetitions need to be checked, and word represents the word to be checked for repetition.
It calls the count_word_occurrences function, passing the text and word parameters, to get the count of occurrences of the word in the text.
It assigns the count to the variable count.
It checks if the count is greater than or equal to 2, indicating that the word appears more than twice.
If the condition is true, it returns True to indicate that there is repetition of the word.
If the condition is false, it returns False to indicate that there is no repetition of the word.

***
### remove last coma

``` 
    def remove_last_comma(string):
    if string.endswith(" "):
        string = string[:-2]  # Remove the last character
    return string 
```

The function takes a parameter called string, which represents the string from which the last comma needs to be removed.
It checks if the string ends with a space character. This is done using the endswith() method.
If the condition is true, indicating that the last character is a space, it removes the last two characters from the string. This is done using slicing with the [:-2] syntax, which excludes the last two characters from the string.
Finally, it returns the modified string.

***
### run program

To run program you just need to run parse_emails function \

``` 
username = EMAIL  #your mail here
password = PASSWORD  #your password for apps here
parse_emails(username, password) 
```
