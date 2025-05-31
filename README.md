# 📄 Compare Two Versions of a Document – Office.js Word Add-in

This Office Add-in allows users to upload two `.docx` files, compare their contents, and insert highlighted differences directly into a Word document. Useful for editors, legal teams, or anyone reviewing document changes.

---

## 🚀 Features

- Upload **two versions** of a Word document
- Compare content using **inline diff**
- Visual differences:
  - 🟢 Additions highlighted in green
  - 🔴 Deletions highlighted in red with strikethrough
- Inserts comparison directly into the open Word document

---

## 🧰 Technologies Used

- [Office.js](https://learn.microsoft.com/office/dev/add-ins/)
- [Mammoth.js](https://github.com/mwilliamson/mammoth.js) – Extracts text from `.docx`
- [jsdiff](https://github.com/kpdecker/jsdiff) – Text comparison
- Vanilla JS & DOM APIs

---

## 📦 Project Structure

```plaintext
compare-docx/
│
├── manifest.xml       # Office Add-in manifest
├── taskpane.html      # UI for the add-in
├── taskpane.js        # Logic for file upload, diff, and Word insertion
├── styles.css         # Optional custom styles
├── assets/            # Icons, images, etc.
├── README.md          # This file
```


---

## 🧪 How It Works

1. Load the add-in in Word.
2. Use the **task pane** to upload two `.docx` files.
3. The add-in converts both files to HTML/text using `Mammoth.js`.
4. Text is compared using `jsdiff.diffWords()`.
5. The differences are formatted into a styled HTML block.
6. This HTML is inserted into the **Word document body**.

---

## 💡 Requirements

- Word (Windows, Mac, or Web with Office.js support)
- Node.js (for sideloading setup)
- Office Add-in host (e.g., sideload with Yeoman or manually install the manifest)

---

## 🛠️ Installation & Setup

### 1. Clone the Repo

```bash
git clone https://github.com/your-username/compare-docx-office-addin.git
cd compare-docx-office-addin
```

### 📦 Install Prerequisites

To get started, install the required global tools:

```bash
npm install -g yo generator-office
```

### Sideload in Word
```bash
npm install
npm start
```

## 📄 Comparison Result Inserted in Word
Here’s how the differences look inside the Word document.

![image](https://github.com/user-attachments/assets/82f998f8-0f0a-48b9-ae6e-11f30e3473aa)

Video url : https://drive.google.com/file/d/1nEbv0Hbo5cNZXfnF-4K3AGl5IeuIggmH/view?usp=share_link

