# Allowance Manager

A private admin dashboard for tracking your student allowance, savings, and money held by your mother.

## Setup

1. Create a Google Sheet with the required sheets (see below).
2. Set up a Google Cloud service account, enable Sheets API, download JSON key.
3. Share your Google Sheet with the service account email.
4. Copy `.env.example` to `.env` and fill in the values.
5. Deploy to Vercel (connect your GitHub repo).
6. Set the environment variables in Vercel's project settings.

## Google Sheet Structure

- `transactions`: Date, Type, Amount, Note, ID
- `allowance_stats`: Month, Date Received, Allowance Amount, Usage, Savings, Balance
- `savings_stats`: Month, Savings, Usage, Balance
- `manual_offsets`: (row 1: wallet_offset, row 2: savings_offset) – values in column B
- `admin`: A1 = username, B1 = password

## Usage

Open the app, login with the credentials from the `admin` sheet, and manage your allowance.
