# 💰 StudioOS: Automated Payroll Engine

A highly precise, low-code hourly-fee calculator and payroll management tool built for boutique studios, operating entirely on Google Sheets and Google Apps Script.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 📖 About The Project

This is the **Financial Operations** module of **StudioOS**. Calculating instructor payroll in a boutique studio often involves complex variables: different hourly rates, co-teaching split payments, and varying class lengths. Manual calculation is prone to errors and takes hours of administrative time.

This automated payroll engine solves data integrity issues by utilizing strict rounding-off mechanisms and automated script triggers to ensure 100% accurate instructor compensation without any manual calculation.

## ✨ Core Features

- **Split-Payment Logic:** Seamlessly handles co-teaching scenarios where class fees need to be split accurately between multiple instructors.
- **Strict Rounding Algorithms:** Built-in financial rounding mechanisms to prevent decimal calculation errors and ensure accounting data integrity.
- **Zero Frontend Maintenance:** Operates entirely within the familiar interface of Google Sheets, empowered by Google Apps Script automations.
- **One-Click Execution:** Studio managers can trigger monthly or weekly payroll calculations with a simple custom menu button in the spreadsheet.

## 🛠️ Tech Stack

- **Backend / Automations:** Google Apps Script (GAS)
- **Database / UI:** Google Sheets

## 🚀 Getting Started

To use this payroll engine for your studio, follow these simple steps:

### 1. Make a Copy of the Database Template
- Click the link below to copy the pre-formatted Google Sheet to your own Google Drive:
  👉 [**Click Here to Copy the Payroll Template**]([https://docs.google.com/spreadsheets/d/15wBhd_rBD_qW7ntqU1bDy1WhYhhvtQZL/edit?usp=sharing&ouid=110827027590017422909&rtpof=true&sd=true]

### 2. Setup the Automation Logic (Google Apps Script)
- Open your newly copied Google Sheet.
- Navigate to `Extensions > Apps Script`.
- Copy the backend logic provided in the `Code.gs` file of this repository.
- Paste it into your Apps Script editor and save.
- Refresh your Google Sheet. You will now see a new custom menu to execute payroll calculations!

## 🔒 Security Note
This is an open-source template. The provided Google Sheet template contains dummy data (`Instructor A`, `Class B`). Do not commit your studio's actual financial data or instructor salaries to any public repository.

## 📄 License
Distributed under the MIT License. See `LICENSE` for more information.
