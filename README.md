# 📧 Email Automation with Python

This project helps automate the process of sending personalized feedback emails to multiple recipients using data collected from a Google Form (stored in an Excel sheet). It’s ideal for anyone looking to save time on bulk email tasks using Python.

---

## 🚀 Features

* Load data from Excel (name + email)
* Compose personalized feedback emails
* Send emails automatically via Gmail
* Simple setup and customization

---

## 📁 Project Structure

```
email-automation/
│
├── data/          # Contains data.xlsx with email addresses and names
├── script.py      # Main Python script to run automation
└── README.md      # Project documentation
```

---

## 💻 Requirements

* Python 3.8 or later
* Gmail account (app password recommended)

Install dependencies with:

```bash
pip install -r requirements.txt
```

Modules used:

* `pandas` – Read Excel file
* `smtplib` – Handle Gmail authentication and sending
* `email.mime` – Build email message structure
* `message` (custom or simple string formatting) – Compose personalized messages

---

## 🔧 How It Works

1. Open and read `data.xlsx` from the `data/` folder.
2. Loop through each name-email row.
3. Format a custom message for each recipient.
4. Authenticate and connect to Gmail SMTP.
5. Send out each email automatically.

---

## ▶️ Usage

Run the script with:

```bash
python script.py
```

Make sure to allow access for less secure apps (or use Gmail app passwords) if using a Gmail account.

---

## 🤝 Contributing

Contributions are welcome! If you’d like to add features (e.g., HTML email, attachments, logging, Google Sheets integration), feel free to fork this repo, create a new branch, and submit a pull request.

Steps:

1. Fork the project
2. Create your feature branch: `git checkout -b feature-name`
3. Commit your changes: `git commit -m 'Add new feature'`
4. Push to the branch: `git push origin feature-name`
5. Open a Pull Request

---

## 🧠 Ideas for Improvement

* Add support for reading data from Google Sheets
* Attach PDFs or images with each email
* Use HTML email templates
* Add a simple GUI using Tkinter or Streamlit

---

## 🙋‍♂️ Author

**Mohit Kumhar**
🔗 [GitHub Profile](https://github.com/mohitkumhar)

---

