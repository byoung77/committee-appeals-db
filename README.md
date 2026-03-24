# ASC Database Interact

A Python-based graphical application for managing and reporting on committee records using a MySQL database and LaTeX-generated PDF reports.

---

## 📌 Overview

This project provides a complete workflow for storing, retrieving, and reporting on structured committee records.

It combines:

* **Python (Tkinter GUI)** for user interaction
* **MySQL** for structured data storage
* **LaTeX** for generating professional PDF reports

The system enables both detailed record retrieval and high-level statistical analysis, supporting real-world administrative decision-making.

This project was originally developed for the Academic Standards Committee (ASC) at Wilkes University to replace fragmented record-keeping across multiple documents with a structured, searchable system for managing student appeal records.

---

## ✨ Features

* Tkinter-based GUI for database interaction
* MySQL backend for structured data storage
* LaTeX-powered PDF report generation
* Date-range reporting
* Targeted reporting by student identifier (WIN)
* Cover-page summaries for multi-student reports
* Aggregate statistical reporting (approval rates, repeat cases, etc.)
* Single-entry creation, editing, and deletion
* Bulk data upload via CSV
* Full database backup and restore via SQL dumps

---

## ⚙️ Requirements

* Python 3.x
* MySQL (local installation)
* LaTeX (TeX Live recommended)

### Python Packages

```bash
pip install mysql-connector-python
```

---

## 🚀 Setup Instructions

### 1. Clone or Download the Repository

Download from GitHub or clone locally:

```bash
git clone https://github.com/<username>/committee-appeals-db.git
cd committee-appeals-db
```

---

### 2. Set Up MySQL

Start MySQL and log in:

```bash
mysql -u root -p
```

Create the database:

```sql
CREATE DATABASE Appeals_DB;
```

---

### 3. Install LaTeX

On Ubuntu / Pop!_OS:

```bash
sudo apt install texlive-full
```

---

### 4. Run the Application

```bash
python ASC_DB_Interact.py
```

---

## 🔑 Database Login

When the application launches, you will be prompted for a password.

**Important:**

* This is the **MySQL root password** for your local installation
* The application connects using:

  * user: `root`
  * host: `localhost`
  * database: `Appeals_DB`

Example:

```bash
mysql -u root -p
```

Enter the same password in the GUI.

---

## ⚠️ First-Time Setup Behavior

* **Incorrect password** → error dialog displayed
* **Database does not exist** → created automatically
* **Empty database** → you must restore from a backup or load test data

---

## 🧪 Included Test Data

The repository includes a complete test dataset:

* `examples/Test_DB_CSV.csv` – for bulk upload
* `examples/Test_DB_SQL_FILE.sql` – for full database restoration

All data is synthetic:

* Randomly generated student names and IDs
* Sample “minutes” sourced from public domain texts

---

## 📄 Example Outputs

The application generates several types of PDF reports using LaTeX.

---

### 🗓️ Date Range Report

Returns all records within a specified date range.

📄 `examples/test_report_1.pdf`

* Structured record-level output
* Includes full meeting minutes and decisions

---

### 🔍 Specific WINs Report

Returns all records associated with selected student IDs.

📄 `examples/test_report_2.pdf`

Includes a **cover page summary** that:

* Lists only students with matching records
* Omits students with no entries
* Displays the number of records per student

Useful for reviewing repeat appearances and prior decisions.

---

### 📊 Detailed Stats Report

Generates an aggregate summary of the database.

📄 `examples/AppealsDB_Stats_2026_03_24_0813.pdf`

Includes:

* Total number of records
* Number of distinct students
* Date range coverage
* Approval/denial counts and percentages
* Breakdown by first-time, second-time, third-time, and repeat appeals
* Outcome percentages within each category

---

## 📁 Project Structure

```
committee-appeals-db/
│
├── README.md
├── LICENSE
├── ASC_DB_Interact.py
├── ASC_Database_User_Manual.pdf
├──
```
