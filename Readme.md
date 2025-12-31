# Excel-Driven Document Automation with Python

A reusable automation pattern for generating Microsoft Word documents directly from structured Excel data using Python.

This project demonstrates how Excel can act as a data source, Python as the processing engine, and Word templates as dynamic outputs — eliminating repetitive manual document creation.

---

## 🚀 Project Overview

Instead of hardcoding content into documents, this system follows a **data-driven approach**:

- Data is maintained in an Excel (.xlsx) file
- Python reads and processes the data
- Word documents are generated dynamically using templates

The same logic can be reused for multiple real-world use cases by changing only the Excel schema and the Word template.

---

## ⚙️ Tech Stack

- **Python**
- **openpyxl** – Reading, writing, and iterating Excel files
- **docxtpl (python-docx-template)** – Word document templating using Jinja2
- **python-docx** – Handling Word document structure (runs, paragraphs, tables)
- **Jinja2** – Dynamic placeholders and looping logic

---

## 🧠 Core Automation Logic

This project focuses on the following reusable concepts:

- Converting Excel rows into Python dictionaries
- Mapping structured data to document templates
- Iterating over datasets to generate multiple documents
- Separating **data**, **template**, and **business logic**
- Automating bulk document creation using templates

These concepts are applicable to any document automation workflow.

---

## 🎥 Demo Video

Watch the LinkedIn demo showing Excel-driven document automation in action:  
👉 https://www.linkedin.com/posts/akshaya-sen_python-pythonautomation-documentautomation-activity-7406307855647084546-4wS2

---

## 📦 Example Use Cases

Using the same Excel → Python → Template pipeline, you can generate:

- 🎓 Certificates (participation, merit, completion)
- 🧾 Invoices and billing documents
- 📄 Offer / appointment letters
- 📊 Student mark sheets and grade cards
- 🧑‍💼 HR onboarding documents
- 📋 Workshop or event participant letters
- 🏫 Academic documents
- 📑 Legal or administrative forms

Only the **Excel structure** and **Word template** change — the automation logic remains the same.

---

## 📂 Project Structure

excel-document-automation/
├── src/
│ └── generate_documents.py
├── templates/
│ └── certificate.docx
├── sample_data/
│ └── wshop.xlsx
├── output/
│ └── generated_docs/
├── README.md

