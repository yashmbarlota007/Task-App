# ðŸ“‹ Yash Marketing - Task Operations & Staff Portal

Live production web application for task management, recurring routines, multi-photo assignment, store manager order fulfillment, and comprehensive 7-period staff performance scorecards.

---

## âš¡ Instant Live Deployment with Vercel (Recommended)

This repository is configured identically to Yash-Marketing-Backend-App to deploy instantly on **Vercel** with automatic live updates and zero caching issues.

### How to Deploy on Vercel:
1. Log in to [Vercel](https://vercel.com).
2. Click **Add New...** -> **Project**.
3. Import your GitHub repository: yashmbarlota007/Task-App.
4. Keep the default settings (**Framework Preset: Other**, Root Directory: ./).
5. Click **Deploy**.

Every time you or an agent push code to the main branch, Vercel will automatically build and deploy the update live in 5-10 seconds!

---

## âš™ï¸ Connecting Backend API (Google Apps Script)

1. Open your Google Spreadsheet for Task Operations.
2. Go to **Extensions > Apps Script**.
3. Copy the code from Code.gs in this repository and paste it into your Apps Script editor.
4. Click **Deploy > New Deployment** (or **Manage Deployments**):
   - **Type:** Web app
   - **Execute as:** Me (your Google account)
   - **Who has access:** Anyone
5. Copy the Web App URL (https://script.google.com/macros/s/.../exec).
6. In your live Vercel App (or local browser):
   - Click **âš™ï¸ Backend API Settings** on the login screen.
   - Paste your Apps Script Web App URL.
   - Click **Test Connection** (turns green âœ…).
   - Click **Save & Reload**.

---

## ðŸŒŸ Key Features

1. **Bug Fix: Recurring Tasks (Daily, Weekly, Monthly) Auto-Reset & Visibility:**
   - Tasks assigned to "All Staff" or specific staff appear in the relevant tabs (Daily, Assigned, etc.).
   - Cycle tracking automatically resets recurring tasks for the next day/week/month upon completion.

2. **7-Period Performance Scorecard Engine:**
   - Supports 7 distinct timeframes: *Daily (Today)*, *Weekly*, *Monthly*, *Last Month*, *Current Quarter*, *Last Quarter*, *Yearly*.
   - Available on Staff Dashboard ("My Score / Performance") and Owner Dashboard ("Staff Scores" leaderboard).

3. **Multiple Reference Photo Uploads & Zoom Gallery:**
   - Owner can upload multiple photos during task creation.
   - Thumbnails with remove option and click-to-zoom modal gallery.

4. **Universal Dual-Environment Hybrid Bridge:**
   - Seamlessly runs both inside Google Apps Script (HtmlService) and on Vercel / mobile browsers via REST API (doPost).