# 💰 OFX to XLSX/CSV Converter

<div align="center">

**A powerful and simple Node.js tool to convert multiple OFX bank statement files to XLSX and CSV formats**

[Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Output Format](#-output-format) • [Contributing](#-contributing)

</div>

---

## 📋 Overview

This tool automatically processes multiple OFX (Open Financial Exchange) files and converts them into Excel spreadsheets (XLSX) and CSV files. Perfect for financial analysis, accounting, and data integration tasks.

### ✨ Features

- 🚀 **Batch Processing** - Convert multiple OFX files at once
- 📊 **Dual Output** - Generates both XLSX and CSV formats
- 💼 **Professional Format** - Clean, organized spreadsheet structure
- 📈 **Financial Summary** - Shows credits, debits, and balance for each file
- 🎯 **Zero Configuration** - Just drop your files and run
- 🔄 **Preserves Data** - Maintains original transaction details
- ⚡ **Fast & Efficient** - Processes thousands of transactions in seconds

---

## 🛠️ Installation

### Prerequisites

- [Node.js](https://nodejs.org/) (version 14 or higher)
- npm (comes with Node.js)

### Quick Start

1. **Clone the repository**
```bash
git clone https://github.com/ojosealbertoalves/myofxconvertertocsv.git
cd myofxconvertertocsv
```

2. **Install dependencies**
```bash
npm install
```

3. **Create the input folder**
```bash
mkdir arquivos-ofx
```

4. **You're ready to go!** 🎉

---

## 🚀 Usage

### Basic Usage

1. Place your `.ofx` files in the `arquivos-ofx/` folder
2. Run the converter:
```bash
node server.js
```
3. Find your converted files in the `convertidos/` folder

### Example

```bash
# Place files in input folder
arquivos-ofx/
  ├── january-statement.ofx
  ├── february-statement.ofx
  └── march-statement.ofx

# Run converter
node server.js

# Get converted files
convertidos/
  ├── january-statement.xlsx
  ├── january-statement.csv
  ├── february-statement.xlsx
  ├── february-statement.csv
  ├── march-statement.xlsx
  └── march-statement.csv
```

---

## 📊 Output Format

### Spreadsheet Columns

| Column   | Description                          |
|----------|--------------------------------------|
| `tipo`   | Transaction type (CREDIT or DEBIT)   |
| `data`   | Transaction date (DD/MM/YYYY)        |
| `valor`  | Transaction amount                   |
| `descric`| Transaction description              |
| `id`     | Unique transaction identifier        |
| `checks` | Empty column for manual annotations  |

### Sample Output

```
tipo    | data       | valor    | descric
--------|------------|----------|----------------------------------
CREDIT  | 27/10/2025 | 1000.00  | PIX Received from COMPANY XYZ
DEBIT   | 27/10/2025 | -480.00  | PIX Sent to JOHN DOE
DEBIT   | 27/10/2025 | -129.90  | PIX Sent to JANE SMITH
```

---

## 📁 Project Structure

```
ofxconverter/
│
├── arquivos-ofx/           # Input folder (place your .ofx files here)
│   ├── statement1.ofx
│   └── statement2.ofx
│
├── convertidos/            # Output folder (auto-created)
│   ├── statement1.xlsx
│   ├── statement1.csv
│   ├── statement2.xlsx
│   └── statement2.csv
│
├── node_modules/           # Dependencies (auto-created)
├── package.json            # Project configuration
├── server.js               # Main converter script
└── README.md               # Documentation
```

---

## 💻 Technical Details

### Dependencies

- **xlsx** - Excel file generation and manipulation
- **Node.js built-in modules** - File system operations

### How It Works

1. **Parse OFX** - Reads and extracts transaction data from OFX files
2. **Format Data** - Converts dates and organizes information
3. **Generate Files** - Creates XLSX and CSV with proper formatting
4. **Summary Report** - Displays financial totals and statistics

### Supported OFX Formats

- OFX version 1.0.2
- OFX version 2.x
- Bank statements (BANKMSGSRSV1)
- Credit card statements

---

## 📈 Console Output Example

```
============================================================
🚀 CONVERSOR DE ARQUIVOS OFX PARA XLSX/CSV
============================================================

📂 Found 3 OFX file(s) to process:

[1/3] Processing: january-statement.ofx
   ✓ 856 transactions found
   ✓ Credits: R$ 45000.00
   ✓ Debits: R$ 42500.00
   ✓ Balance: R$ 2500.00
   ✓ Saved: january-statement.xlsx
   ✓ Saved: january-statement.csv

[2/3] Processing: february-statement.ofx
   ✓ 923 transactions found
   ✓ Credits: R$ 48000.00
   ✓ Debits: R$ 45000.00
   ✓ Balance: R$ 3000.00
   ✓ Saved: february-statement.xlsx
   ✓ Saved: february-statement.csv

[3/3] Processing: march-statement.ofx
   ✓ 1045 transactions found
   ✓ Credits: R$ 52000.00
   ✓ Debits: R$ 49000.00
   ✓ Balance: R$ 3000.00
   ✓ Saved: march-statement.xlsx
   ✓ Saved: march-statement.csv

============================================================
📊 FINAL SUMMARY
============================================================
✅ Files processed: 3/3
📈 Total transactions converted: 2824
📁 Files saved in: ./convertidos
============================================================
```

---

## 🔧 Configuration

You can customize folder names by editing these lines in `server.js`:

```javascript
const pastaOfx = './arquivos-ofx';      // Input folder
const pastaDestino = './convertidos';   // Output folder
```

---

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Ideas for Contributions

- [ ] Add support for more date formats
- [ ] Implement file drag-and-drop interface
- [ ] Add GUI version
- [ ] Support for other financial file formats (QIF, QFX)
- [ ] Add filtering options (date range, amount range)
- [ ] Generate summary reports with charts

---

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🐛 Troubleshooting

### Common Issues

**Error: "Cannot find module 'xlsx'"**
```bash
Solution: npm install
```

**Error: "Folder 'arquivos-ofx' does not exist"**
```bash
Solution: mkdir arquivos-ofx
```

**Error: "No .ofx files found"**
```bash
Solution: Place .ofx files in the arquivos-ofx/ folder
```

**Error: "node is not recognized"**
```bash
Solution: Install Node.js from https://nodejs.org
```

---

## 📫 Contact

Alberto Alves - [@ojosealbertoalves](https://github.com/ojosealbertoalves)

Project Link: [https://github.com/ojosealbertoalves/myofxconvertertocsv](https://github.com/ojosealbertoalves/myofxconvertertocsv)

---

## ⭐ Show your support

Give a ⭐️ if this project helped you!

---

## 📚 Additional Resources

- [OFX Specification](http://www.ofx.net/)
- [Node.js Documentation](https://nodejs.org/docs/)
- [SheetJS (xlsx) Library](https://sheetjs.com/)

---

<div align="center">

**Made with ❤️ and JavaScript**

</div>
