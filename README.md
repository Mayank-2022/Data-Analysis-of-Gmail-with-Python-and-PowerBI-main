Here's a **step-by-step guide** to implement the **Gmail Data Analysis with Python and Power BI** project.  

---

## **Step 1: Enable IMAP Access & Generate App Password**
Before extracting emails from Gmail, you need to enable **IMAP access** and generate an **app password** for authentication.  

### **1.1 Enable IMAP in Gmail**  
1. Go to [Gmail IMAP Settings](https://mail.google.com/mail/u/0/#settings/fwdandpop/).
2. Scroll down to the **IMAP access** section.
3. Select **Enable IMAP**.
4. Click **Save Changes**.  

### **1.2 Generate an App Password**  
1. Go to [Google App Passwords](https://myaccount.google.com/apppasswords).  
2. Sign in to your Google account (if required).  
3. Click on **Select App** → Choose **Mail**.  
4. Click on **Select Device** → Choose **Your Device**.  
5. Click **Generate**, and copy the password shown (you’ll need this later).  

---

## **Step 2: Install Required Python Libraries**
To extract, process, and analyze emails, install the required Python libraries:

```bash
pip install imapclient email pandas numpy tqdm wordcloud pillow matplotlib
```

---

## **Step 3: Connect to Gmail and Fetch Emails**
Create a Python script to connect to Gmail using IMAP, authenticate using the **app password**, and fetch emails.

### **3.1 Import Required Libraries**
```python
import imaplib
import email
from email.header import decode_header
import pandas as pd
from tqdm import tqdm
import numpy as np
from wordcloud import WordCloud, STOPWORDS
import matplotlib.pyplot as plt
from PIL import Image
from matplotlib.colors import LinearSegmentedColormap
```

### **3.2 Establish Connection to Gmail**
```python
# Gmail credentials
user = "your_email@gmail.com"
password = "your_app_password"

# Connect to Gmail IMAP server
imap_url = 'imap.gmail.com'
mail = imaplib.IMAP4_SSL(imap_url)
mail.login(user, password)

# Select the inbox
mail.select("Inbox")
```

### **3.3 Fetch Emails from Inbox**
```python
# Get the total number of emails in the inbox
status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()

# Create a DataFrame to store email details
email_df = pd.DataFrame(columns=['Date', 'From', 'Subject'])

# Loop through emails and extract details
for i in tqdm(email_ids[::-1], desc="Fetching Emails"):
    res, msg_data = mail.fetch(i, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            subject = msg["subject"]
            sender = msg["from"]
            date = msg["date"]

            # Decode subject if needed
            if subject:
                subject, encoding = decode_header(subject)[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")

            # Append data to DataFrame
            email_df = email_df.append({'Date': date, 'From': sender, 'Subject': subject}, ignore_index=True)

# Save DataFrame to CSV
email_df.to_csv("emails.csv", index=False, encoding="utf-8")
```

---

## **Step 4: Clean and Transform Data**
### **4.1 Format Date and Time Columns**
```python
# Convert date to datetime format
email_df['Date'] = pd.to_datetime(email_df['Date'], errors='coerce')

# Extract useful time details
email_df['H_M_S'] = email_df['Date'].dt.strftime('%H:%M:%S')
email_df['Hour'] = email_df['Date'].dt.strftime('%H') + 'h-' + (email_df['Date'].dt.hour + 1).astype(str) + 'h'
email_df['WeekDay'] = email_df['Date'].dt.strftime('%A')
email_df['Date'] = email_df['Date'].dt.date

# Save cleaned data
email_df.to_csv("emails_cleaned.csv", index=False, encoding="utf-8")
```

### **4.2 Extract Sender Email and Name**
```python
# Extract email addresses from "From" column
def extract_email(sender):
    if '<' in sender:
        return sender.split('<')[-1].split('>')[0]
    return sender

# Extract sender name
def extract_name(sender):
    name, encoding = decode_header(sender)[0]
    if isinstance(name, bytes):
        return name.decode(encoding if encoding else "utf-8")
    return name

email_df["Mail"] = email_df["From"].apply(extract_email)
email_df["Name"] = email_df["From"].apply(extract_name)
email_df.drop(columns=['From'], inplace=True)

# Save final cleaned dataset
email_df.to_csv("emails_final.csv", sep="|", index=False, encoding="utf-8")
```

---

## **Step 5: Create a Word Cloud from Email Subjects**
### **5.1 Define Color Palette**
```python
# Function to normalize colors
def inter_from_256(x):
    return np.interp(x=x, xp=[0, 255], fp=[0, 1])

# Gmail color theme
cdict = {
    'red': ((0.0, inter_from_256(197), inter_from_256(197)), (1.0, inter_from_256(251), inter_from_256(251))),
    'green': ((0.0, inter_from_256(34), inter_from_256(34)), (1.0, inter_from_256(188), inter_from_256(188))),
    'blue': ((0.0, inter_from_256(31), inter_from_256(31)), (1.0, inter_from_256(4), inter_from_256(4))),
}
gmail_colormap = LinearSegmentedColormap('gmail', segmentdata=cdict)
```

### **5.2 Generate and Plot Word Cloud**
```python
# Concatenate all email subjects into a single string
text_data = " ".join(str(subject).lower() for subject in email_df["Subject"].dropna())

# Define stopwords to remove common words
stopwords = set(STOPWORDS).union({"re", "fw", "fwd", "subject"})

# Generate Word Cloud
wordcloud = WordCloud(
    width=800, height=400, background_color="black",
    colormap=gmail_colormap, stopwords=stopwords
).generate(text_data)

# Display the word cloud
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud, interpolation="bilinear")
plt.axis("off")
plt.show()
```

---

## **Step 6: Data Visualization with Power BI**
### **6.1 Load Data into Power BI**
1. Open **Power BI Desktop**.
2. Click on **Get Data → Text/CSV**.
3. Select **emails_final.csv** and load the data.

### **6.2 Create Visualizations**
- **Top 10 Senders** → Bar Chart showing the top senders with the most emails.
- **Emails by Date & Time** → Line Chart to analyze email trends over time.
- **Word Cloud** → Use the **Word Cloud Custom Visual** in Power BI to show the most common words in email subjects.
- **Emails by Weekday & Hour** → Heatmap to understand email frequency distribution.

### **6.3 Publish Report**
1. Click **Publish** to upload the report to **Power BI Service**.
2. Share insights with stakeholders.

---

## **Final Output**
- **CSV files**: `emails.csv`, `emails_cleaned.csv`, `emails_final.csv`
- **Power BI Dashboard** with email insights
- **Word Cloud** visualization of most common email subject words

---

## **Conclusion**
This project demonstrates how to extract, clean, and analyze **Gmail email data** using **Python and Power BI**. The insights include the **most frequent senders, email trends, and commonly used words** in email subjects.