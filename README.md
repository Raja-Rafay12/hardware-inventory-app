# 🛠️ Hardware Inventory App

A premium, cloud-enabled Desktop inventory management and Point of Sale (POS) system designed specifically for hardware and tool supply stores. Built on top of **Electron**, **React**, and **PostgreSQL (Neon)**, this application allows store owners to track stock levels, calculate profits, manage a massive inventory catalog, issue invoices, and send automated client emails.

---

## 🏗️ Architecture Overview

The application utilizes Electron's secure architecture with isolated context environments, communicating via Inter-Process Communication (IPC) to link the frontend React renderer with backend Node services and cloud infrastructure.

```mermaid
flowchart TD
    subgraph "Desktop App Client (Local)"
        Renderer["React Renderer Process<br/>Vite & UI Components"]
        Preload["Preload Script<br/>Context Bridge & IPC Renderer"]
        Main["Electron Main Process<br/>Node.js Environment"]
        
        Renderer <--> |IPC Calls & Invokes| Preload
        Preload <--> |IPC Main Handlers| Main
    end
    
    subgraph "Cloud Infrastructure"
        DB[("Cloud PostgreSQL<br/>Neon.tech")]
        SMTP["Google SMTP Server<br/>Nodemailer / Gmail API"]
        
        Main --> |Secure Connection Pool| DB
        Main --> |SMTP Welcome Emails| SMTP
    end

    style Renderer fill:#f9f6f0,stroke:#d9720b,stroke-width:2px;
    style Main fill:#f0f5fa,stroke:#0f60b6,stroke-width:2px;
    style DB fill:#ecfdf5,stroke:#10b981,stroke-width:2px;
    style SMTP fill:#fff5f5,stroke:#ef4444,stroke-width:2px;
```

---

## ✨ Core Features

*   📊 **Real-time Dashboard**: Track total sales, profit margins, active expenses, and inventory health metrics at a glance.
*   📦 **Inventory Catalog Management**: Pre-loaded catalog with over **1,800+ hardware items** sorted into categories (Electrical, Plumbing, Paint, Tools, etc.) with custom markup and alert limits.
*   🧾 **Point of Sale (POS) Billing**: Dynamic basket calculator featuring instant profit analysis, client search, print/export receipts, and sales dispatch logs.
*   📧 **Automated Welcome Emails**: Automatically registers new staff accounts and dispatches visually designed, professional welcome emails via Google SMTP.
*   🔒 **Secure Cloud Synchronization**: Offloads accounts, sales, and catalog datasets to a secure PostgreSQL database, ensuring multi-device persistence.

---

## 💻 Tech Stack

*   **Frontend**: React (19.x), Vite, Tailwind-like Vanilla CSS styling, Lucide React (Icons).
*   **Desktop Shell**: Electron (43.x) with secure Context Isolation.
*   **Database**: PostgreSQL client (`pg`), fully optimized with indexes for query execution speeds.
*   **Authentication**: Salted password hashing via `bcryptjs`.
*   **Mailing**: Transporter relays via `nodemailer`.
*   **Data Parsers**: OpenXML sheet processing via `xlsx`.

---

## ⚙️ Project Structure

```text
├── 📂 public/                  # Static assets & SVG icons
├── 📂 src/
│   ├── 📂 components/          # React authentication & UI widgets
│   ├── 📄 HardwareInventoryApp.jsx  # Main application logic & UI layout
│   ├── 📄 supabase.js          # Supabase client integration hooks
│   └── 📄 main.jsx             # React mounting entry point
├── 📄 electron-main.cjs        # Main Electron process, database pools, and SMTP relays
├── 📄 preload.cjs              # Safe context-bridge interfaces exposed to React
├── 📄 postgres-schema.sql      # Clean relational database tables & indexes configuration
├── 📄 extract_products.py      # Python utility to ingest raw Excel listings into the catalog
├── 📄 inspect_xlsx.py          # Python utility to debug Excel spreadsheet dimensions
├── 📄 .env.example             # Template file showing expected environment keys
└── 📄 package.json             # NPM dependencies & Desktop build commands
```

---

## 🚀 Setup & Installation

### Prerequisites
- [Node.js](https://nodejs.org/) (v18 or higher recommended)
- A cloud PostgreSQL instance (such as a free tier database on [Neon.tech](https://neon.tech/))
- A Gmail account with an App Password enabled (if using welcome emails)

### 1. Clone & Install Dependencies
```bash
npm install
```

### 2. Configure Environment Variables
Create a `.env` file in the root folder (this file is ignored by version control to protect your credentials). Copy the structure from `.env.example`:

```ini
# Cloud PostgreSQL Connection URI
PG_CONNECTION_STRING=postgresql://username:password@hostname:5432/databasename?sslmode=require

# Google SMTP (Gmail) Email Server Details
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_SECURE=true
SMTP_USER=your-email@gmail.com
SMTP_PASS=your-16-character-google-app-password
SMTP_FROM="Hardware Inventory <your-email@gmail.com>"
```

### 3. Initialize Database
Execute the SQL statements in [postgres-schema.sql](file:///c:/Users/rafyt/.gemini/antigravity-ide/scratch/hardware-inventory-app/postgres-schema.sql) in your PostgreSQL query editor to set up tables and optimized query indexes.

### 4. Running Locally

To run the application in development mode:
```bash
# Start Vite development server
npm run dev

# Launch Electron shell
npm start
```

To compile and package the app into a portable desktop executable:
```bash
npm run dist
```

---

## 📊 Catalog Ingestion Pipeline

If you want to parse new Excel files into seed inventory, the project includes helper scripts:
1. **[inspect_xlsx.py](file:///c:/Users/rafyt/.gemini/antigravity-ide/scratch/hardware-inventory-app/inspect_xlsx.py)**: Scans dimensions and structure of Excel pricing lists inside the user's `Documents` directory.
2. **[extract_products.py](file:///c:/Users/rafyt/.gemini/antigravity-ide/scratch/hardware-inventory-app/extract_products.py)**: Auto-categorizes, cleans null characters, computes default markups, and builds `products.json` relative to the workspace directory.
