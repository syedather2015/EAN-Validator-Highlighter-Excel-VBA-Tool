# 🧾 EAN Validator & Highlighter – Excel VBA Tool

This Excel macro streamlines the process of validating **EAN (European Article Numbers)** by checking formatting rules, ensuring correct length, and automatically applying **13-digit padding**. Invalid entries are visually highlighted, making it easier to spot and correct them during product data onboarding or audit processes.

> Developed by **Syed Ather Rizvi** for product data integrity and quality assurance.

---

## ⚙️ Features

- ✅ Prompts user to select the column with EANs
- 🔢 Automatically pads EANs with leading zeros to make them 13 digits
- 🧠 Applies rule-based checks (no leading "2", no "00000" patterns, etc.)
- 🎨 Highlights invalid EANs with **blue background** and **white text**
- 📊 Works on the **active worksheet** – no setup required

---

## 🧪 Validation Rules Applied

A valid EAN must:
- Be numeric
- Be 13 digits or less (padded if needed)
- **Not start with `2`**
- **Not contain `00000`** in positions:
  - 1–3
  - 3–7
  - 8–13

---

## 🔁 How to Use

1. Open your Excel file with EAN data
2. Press `Alt + F11` to open the **VBA Editor**
3. Insert a new Module and paste the macro code
4. Run `ValidateAndHighlight` from the Macro window
5. When prompted, enter the **column letter** where EANs exist (e.g., `B`, `D`, etc.)

---

## 💡 Use Cases

- Product onboarding & GTIN validation  
- Retail or eCommerce data audits  
- Supplier catalog quality checks  
- Marketing intelligence & pricing validation  

---

## 📌 Visual Output

- Valid EANs are auto-corrected to 13 digits
- Invalid EANs are marked **blue** with **white text** for easy identification

---

## 📄 Code Highlights

```vba
If IsValidEAN(ean) Then
    newEAN = WorksheetFunction.Rept("0", 13 - Len(ean)) & ean
    If IsValidEAN(newEAN) Then
        cell.Value = newEAN
    Else
        HighlightCell cell
    End If
Else
    HighlightCell cell
End If
