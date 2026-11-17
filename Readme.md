# VehicleTM

A Google Apps Script suite designed to automate vehicle rental fleet operations, financial calculations, maintenance intervals, reservation conflict detection, cash register tracking, and automated daily email/Telegram alerts directly inside Google Sheets.

---

## Overview & Core Capabilities

VehicleTM turns a standard Google Spreadsheet into an operational fleet management system across four synchronized modules:

### 1. Rental Operations (`Locations.gs`)
* **Automated Formulas:** Populates empty cells across up to 1,500 rows without overwriting existing data.
  * **Consumed Mileage:** `Return KM - Delivery KM`
  * **Duration (Days):** `Return Date - Delivery Date`
  * **Total HT & TTC:** `Unit Price * Days` and `Total HT + 19% VAT`
  * **Balance & Debt:** Dynamically parses numeric payments or text statuses (e.g., *"paid in full"*, *"did not pay"*) to calculate remaining balances (`RESTE`) and overdue debts (`DETTE`) once a vehicle is returned or the return date passes.
* **Status Color Coding:** Applies conditional formatting to highlight active rentals (red), overdue returns (dark red), completed rentals (green), and upcoming deliveries (yellow).

### 2. Fleet & Maintenance Tracking (`Voitures`)
* **Automatic Fleet Sync:** Scans the rentals sheet for all license plates, adds missing vehicles to the `Voitures` tab, and updates each vehicle's current odometer reading using the latest logged mileage.
* **Oil Change Alerts:** Tracks oil change intervals (5,000 km to 20,000 km dropdowns) and flags vehicles approaching or exceeding their service threshold.
* **Document Expiration:** Monitors insurance and technical inspection expiration dates.

### 3. Reservation Scheduling & Conflict Detection (`Réservations`)
* **Double-Booking Prevention:** Cross-references both the `Réservations` sheet and active contracts in the rentals sheet by license plate (`MAT`) or vehicle model.
* **Live Statuses:** Automatically flags rows as `Conflit` (double-booked), `En cours` (active), `À venir` (upcoming), or `Terminée` (completed).

### 4. Cash Register & Ledger (`Caisse.gs`)
* **Real-Time Balance:** Tracks general and monthly totals for Income (`REVENUS`), Expenses (`DÉPENSES`), and Cash Advances (`ACOMPTES`), updating the net register balance immediately.
* **Contract Auto-Lookup:** Entering a contract number automatically retrieves the client's name and vehicle model from the rentals sheet.
* **Input Normalization:** Automatically converts text-formatted dates and currency strings (e.g., `"12 500,00 DA"`) into valid numbers and dates on edit.

### 5. Automated Daily Digests & Alerts (`Code.gs`)
* **Scheduled Morning Reports:** Sends a responsive HTML email summary each morning detailing today's/tomorrow's deliveries and returns, overdue vehicles (with direct WhatsApp follow-up links), recent returns with mileage driven, outstanding client debts, maintenance alerts, and monthly financial summaries.
* **Hourly & Telegram Notifications:** Optional triggers for upcoming deliveries/returns within a configurable time window.

---

## Project Structure

Ensure your Google Apps Script project contains the following script files alongside each other:

* `Code.gs` — Global configuration (`CONFIG`), custom menu initialization, date/currency parsers, data readers, and email/Telegram notification engine.
* `Locations.gs` — Rental sheet setup, conditional formatting rules, fleet mileage synchronization (`Voitures`), and reservation conflict engine (`Réservations`).
* `Caisse.gs` — Cash register sheet layout, live balance formulas, contract lookup formulas, and `onEdit` input normalization.

---

## Installation & Setup

### Step 1: Open Apps Script
1. Open your Google Sheet.
2. Navigate to **Extensions** > **Apps Script**.

### Step 2: Add the Script Files
1. Paste the contents of `Code.gs` into the default `Code.gs` file.
2. Click the **+** icon next to **Files** > **Script**, name it `Locations`, and paste the locations module code.
3. Create another script file named `Caisse` and paste the cash register module code.
4. Click **Save** (`Ctrl + S`).

### Step 3: Configure Parameters
In `Code.gs`, update the `CONFIG` object at the top of the file to match your environment:
* `DESTINATAIRES`: Comma-separated email addresses to receive daily reports.
* `SOCIETE` & `SIGNATURE`: Company name and email footer signature.
* `FEUILLE_LOCATIONS`: Exact name of your main rentals tab (defaults to the first tab if not found).
* `HEURE_ENVOI` & `FUSEAU`: Delivery hour and timezone (e.g., `'Africa/Algiers'`).
* `COL`: Verify that column letters match your existing rentals sheet layout.

---

## Usage Guide

1. **Reload your Google Sheet.** A custom menu named **Flotte IAM** will appear in the top toolbar.
2. **Initialize Sheets:**
   * Run **Flotte IAM > 1. Créer les onglets (1re fois)** (or individual sheet preparation actions) to generate headers, formulas, data validations, and conditional formatting without erasing existing data.
   * Run **Mettre à jour les voitures** to populate the `Voitures` tab from your rental history, then manually enter the last oil change mileage for each vehicle (highlighted in orange).
   * Run **Préparer la feuille Caisse** to set up the cash register tab.
3. **Configure Notifications:**
   * Select **Flotte IAM > 2. Choisir le contenu de l’e-mail** to toggle which sections appear in your daily digest.
   * Run **Flotte IAM > 3. Envoyer un e-mail de test** to verify formatting and delivery.
   * Run **Flotte IAM > 4. Activer l’envoi automatique** to schedule the daily morning trigger.
4. **Enable Auto-Formatting on Edit (Optional):**
   * In the Apps Script editor, go to **Triggers** (alarm clock icon) > **Add Trigger**.
   * Set function `auCalculCaisse_` to run **From spreadsheet** > **On edit** so text-entered numbers and dates in the `Caisse` tab automatically convert to numeric formats.

---

## License

MIT License
