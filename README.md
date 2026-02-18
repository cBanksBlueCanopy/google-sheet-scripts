# 📊 Google Sheet Scripts

A collection of Google Apps Script utilities designed to streamline the process of formatting and preparing Google Sheets data for seamless importing into WooCommerce via the [WP All Import](https://www.wpallimport.com/) plugin.

---

## 📁 Scripts Overview

| Script | Description |
|---|---|
| `attributeSeparator.js` | Splits combined product attribute strings into separate, properly formatted columns |
| `titleCase.js` | Converts text in a column to Title Case formatting |
| `wooCommerceStandard.js` | Applies WooCommerce-standard formatting and structure to product data |
| `wordpressImageURLReplacer.js` | Finds and replaces image URLs to match your WordPress media library paths |

---

## 🚀 Getting Started

### Step 1 — Open Google Apps Script

1. Open your Google Sheet.
2. In the top menu, click **Extensions → Apps Script**.
3. This will open the Apps Script editor in a new tab.

### Step 2 — Add a Script

1. In the Apps Script editor, delete any placeholder code in the default `Code.gs` file (or create a new script file by clicking the **+** next to "Files").
2. Copy the contents of whichever script you want to use from this repository.
3. Paste the code into the editor.
4. Click the **💾 Save** icon (or press `Ctrl+S` / `Cmd+S`).

### Step 3 — Run a Script

1. In the Apps Script editor, use the **function dropdown** at the top to select the function you want to run.
2. Click the **▶ Run** button.
3. The first time you run a script, Google will ask you to **authorize permissions** — click through the prompts and allow access to your spreadsheet.

> **Tip:** You can also trigger scripts directly from your spreadsheet by setting up a custom menu. See [Google's documentation](https://developers.google.com/apps-script/guides/menus) for details.

---

## 📋 Script Usage

### `attributeSeparator.js`

Splits product attribute data that is combined in a single cell (e.g., `"Color: Red | Size: Large"`) into individual columns, formatted to meet WooCommerce attribute import requirements.

**How to use:**
1. Paste the script into the Apps Script editor.
2. Update any configuration variables at the top of the script to match your sheet's column layout (e.g., source column, target columns).
3. Run the function from the editor.

---

### `titleCase.js`

Converts text strings to Title Case (e.g., `"red running shoes"` → `"Red Running Shoes"`). Useful for cleaning up product names and category labels before import.

**How to use:**
1. Paste the script into the Apps Script editor.
2. Set the target column or range in the script's configuration variables.
3. Run the function — the selected cells will be updated in place.

---

### `wooCommerceStandard.js`

Applies WooCommerce-standard formatting rules across your product data. This may include normalizing fields such as product status, stock values, price formatting, or other fields required by the All Import plugin's column mapping.

**How to use:**
1. Paste the script into the Apps Script editor.
2. Review and update the column references at the top of the script to match your spreadsheet's structure.
3. Run the function to process your data.

---

### `wordpressImageURLReplacer.js`

Finds image URLs in your sheet and replaces them with the correct WordPress media library URLs. This is especially useful when migrating products from one domain or staging environment to another.

**How to use:**
1. Paste the script into the Apps Script editor.
2. Set the `oldBaseURL` and `newBaseURL` variables to match your source and destination domains.
3. Specify the column containing your image URLs.
4. Run the function — all URLs in the target column will be updated.

---

## 🛠 Tips & Best Practices

- **Always work on a copy.** Before running any script, duplicate your sheet (`Right-click sheet tab → Duplicate`) to preserve your original data.
- **Check column references.** Each script may reference specific columns by letter or index. Review the configuration variables at the top of each script and update them to match your spreadsheet before running.
- **Authorization is required once.** The first time you run a script, Google will prompt you to grant permission. This is expected — the scripts only access the spreadsheet they are attached to.
- **Combine scripts as needed.** You can copy multiple scripts into separate `.gs` files within the same Apps Script project and run them in sequence to fully prepare your data for WooCommerce import.

---

## 🔗 Related Resources

- [WP All Import Plugin](https://www.wpallimport.com/)
- [WooCommerce Product CSV Import Schema](https://woocommerce.com/document/product-csv-importer-exporter/)
- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Google Apps Script — Spreadsheet Service](https://developers.google.com/apps-script/reference/spreadsheet)

---

## Contributing

Contributions are welcome! If you have a useful Google Sheets script for WooCommerce data prep, feel free to open a pull request.

1. Fork the repository.
2. Create a new branch: `git checkout -b feature/your-script-name`
3. Add your script file and update this README with usage instructions.
4. Submit a pull request.

---

## 📄 License

This project is open source. Feel free to use and adapt these scripts for your own projects.