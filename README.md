# Material Inventory Automation System
**Automates: Purchase Orders → Change Orders → Invoices → Excel Tracker**

---

## What This System Does

1. **Upload PO PDF** → Claude reads it, extracts all materials (Type, T, W, Length, Qty, Cost)
2. **Upload Change Order PDFs** → Claude matches each CO item to your PO and updates quantities
3. **Upload Invoice PDFs** → Claude matches delivered items to your PO, logs qty delivered per invoice
4. **Download Excel** → Auto-generated inventory.xlsx with:
   - PO Qty | CO Qty | PO+CO Qty
   - One column per invoice date with delivered quantities
   - Total Delivered | Remaining | LF | BF calculations

---

## Setup (First Time)

### Requirements
- Python 3.9 or higher
- An Anthropic API key (get one at console.anthropic.com)

### Step 1: Set your API key

**Mac/Linux:**
```bash
export ANTHROPIC_API_KEY=sk-ant-your-key-here
```

**Windows:**
```cmd
set ANTHROPIC_API_KEY=sk-ant-your-key-here
```

Or add it permanently to your `.bashrc` / `.zshrc` / Windows Environment Variables.

### Step 2: Start the system

**Mac/Linux:**
```bash
chmod +x start.sh
./start.sh
```

**Windows:**
```
Double-click start.bat
```

### Step 3: Open the Web UI

Open `index.html` in your browser (Chrome recommended).
The UI connects to `localhost:8000` which is running the backend.

---

## How to Use

### Starting a New Project
1. Type the project name (e.g. "Willow Way Apts") in the top bar
2. Click **+ Create**

### Uploading a Purchase Order
1. Select your project from the dropdown
2. Click "Click to select PDF" under **Purchase Order**
3. Pick your PO PDF → Click **Upload PO**
4. Claude reads all material lines — takes ~20-30 seconds

### Uploading a Change Order
1. Select the CO PDF
2. Enter the CO date (optional)
3. Click **Upload Change Order**
4. Claude matches CO items to your PO and updates quantities

### Uploading an Invoice
1. Select the Invoice PDF
2. Enter the **Invoice Number** (e.g. `60126004`)
3. Enter the invoice date
4. Click **Upload Invoice**
5. Claude matches each delivered item to your PO materials

### Downloading the Excel
Click **⬇ Download Excel** at the top right.
The file will match your manual Willow Way format:
- Columns for each invoice date
- Totals and remaining auto-calculated with Excel formulas

---

## Excel Output Format

| Column | Description |
|--------|-------------|
| Type | Lumber / Panels / LVL / Each |
| Description | Full material spec |
| T, W, Length | Dimensions |
| PO Qty | Original PO quantity |
| CO Qty | Net change from all COs |
| PO+CO Qty | Total quantity needed |
| L/F per Pc | Linear feet per piece |
| B/F per Pc | Board feet per piece |
| Cost/Unit | Unit cost (MBF/MSF/Each) |
| Total Cost | Pre-tax cost |
| Total Cost+Tax | With tax |
| INV#XXXXX Date | Qty delivered per invoice |
| Total Delivered | Sum of all invoice columns |
| Total Delivered LF | Delivered × LF/Pc |
| Total Delivered B/F | Delivered × BF/Pc |
| Remaining | PO+CO − Delivered |

---

## File Storage

All project data is stored in the `projects/` folder:
```
projects/
  Willow Way Apts/
    items.json        ← material data
    meta.json         ← invoice list, CO count
    inventory.xlsx    ← auto-rebuilt Excel
uploads/              ← uploaded PDFs (kept for reference)
```

---

## Troubleshooting

**"Claude API error"** → Check your ANTHROPIC_API_KEY is set correctly

**"No match found for invoice items"** → The invoice format may differ significantly from PO. Claude will mark them as UNMATCHED — you can manually adjust in the Excel.

**Server won't start** → Make sure port 8000 is free. Kill any existing process:
```bash
# Mac/Linux
lsof -ti:8000 | xargs kill
# Windows
netstat -ano | findstr :8000
```

---

## Future Improvements (Roadmap)
- [ ] Manual match correction in the web UI
- [ ] Multi-vendor support (not just Mathews Lumber)
- [ ] Dashboard with cost tracking charts
- [ ] PostgreSQL database for scale
- [ ] Cloud deployment (AWS/GCP)
