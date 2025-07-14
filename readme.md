# Gym Member Tracking System

## Overview

This is a low-cost, efficient gym member tracking system built using **Google Sheets**, **Google Forms**, and **Google Apps Script**. Designed for a local gym, it automates data handling, tracks payments, membership plans, and ensures monthly organization of entries.

## Objectives

* Digitize and streamline gym member tracking
* Eliminate manual clutter from paper logs
* Automate calculations, duplicate checks, and monthly sheet management
* Provide insights into member stats, payments, and plan durations

## 🔗 Live Project Links

- 📄 [View Google Sheet (View-only)](https://docs.google.com/spreadsheets/d/1rBj-08L24_2klMxvUzsnxb_tSCLlhoWDW_dlt25I3SQ/edit?usp=sharing)
- 📝 [View Google Form](https://forms.gle/98CtVCvTxbQ2bz11A)

## Motivation

This project was created to solve a real-world problem: a gym owner who is my friend was managing all member records on paper, leading to clutter, mismanagement, and lack of insight. This system provides a clean, automated alternative.

## Key Features

* **Google Form** for user/member data entry
* **Automated monthly sheet creation** with protected structure
* **Formulas** for calculating plan duration, payment totals, and expiry dates
* **Duplicate checks** to avoid redundant entries
* **Script-based logic** for data transfer, validation, and monthly automation

## Tools & Tech Stack

* **Google Forms** (for member input)
* **Google Sheets** (as a database and dashboard)
* **Apps Script** (for automation and monthly logic)

## How It Works

1. Member submits data via **Google Form**
2. Response is stored in `_RawResponses` sheet
3. Apps Script:

   * Transfers entries to the corresponding **monthly sheet**
   * Performs **duplicate detection** based on name & phone
   * Applies calculated columns for **fees, expiry, start dates**, etc.
   * If a monthly sheet doesn't exist (e.g., August 2025), it creates it by copying a base template

## Challenges Faced

* Parsing dynamic form data into consistent monthly views
* Keeping sheets organized without breaching Apps Script quotas
* Handling calculated columns and protected ranges within script automation

## Use Case & Impact

* Helps gym owners stay organized
* Minimizes errors and redundancy
* No cost, zero maintenance once deployed
* Adaptable to other small businesses with monthly client tracking needs
