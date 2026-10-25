# Company X — Slow Stock Action Tool

A Python + Streamlit application for the Company X Management Trainee case study.

## What it does

- Accepts either one Excel workbook containing four data sheets or four separate Excel files.
- Detects tables from column headings (the README workbook is optional and ignored).
- Standardizes SKU and month formats.
- Reports duplicate sales rows, negative stock, inactive items with stock, missing industries, unmatched SKUs, and other data issues.
- Combines stock and sales for products marked `SWITCH`, reporting under the replacement SKU.
- Calculates aged stock quantity/value, six-month sales, average monthly sales, months to clear, and item class.
- Shows summary counts, top 10 Push items, filters, customer targets, and a mobile-friendly call-list view.
- Exports the results to Excel.

## Run locally

1. Install Python 3.10 or newer if needed.
2. Extract this ZIP into a folder.
3. Open Command Prompt/Terminal in that folder.
4. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

5. Start the app:

   ```bash
   streamlit run app.py
   ```

6. Streamlit will display a local browser address. Open it and upload the Excel file(s).

## Using Jupyter Notebook

If you already have Jupyter open, you can install the dependencies in a notebook cell:

```python
%pip install -r requirements.txt
```

Save `app.py` in the same project folder as `requirements.txt`, then run this in a notebook cell:

```python
!streamlit run app.py
```

If that command does not launch correctly from your Jupyter environment, use Command Prompt in the project folder and run `streamlit run app.py`.

## Upload options

- Option A: one workbook with sheets for Item Master, Stock Aging, Sales Lines and Customers.
- Option B: the four separate Excel files.
- `0_README.xlsx` is optional and will be ignored because it is not one of the four data tables.

## Important calculation choices

- `SWITCH` stock and sales are combined under the replacement SKU. The replacement SKU's pack quantity and unit cost are used.
- Inactive products are excluded from active item classification; inactive items with stock are flagged.
- Negative stock bucket values are set to zero only in the calculation copy, and are reported for source correction.
- Exact duplicate sales rows are excluded from calculations and reported for validation with the data owner.
- Aged quantity = `Qty 91-180 days + Qty 180+ days`.
- Aged value = `Aged quantity × unit cost`.
- Average monthly sales = `six-month sales quantity ÷ 6`.
- Months to clear = `total stock ÷ average monthly sales`. If there are no sales, it displays `No sales`.
- Classification order: Red List if six-month sales are below one master pack; otherwise No Action if aged stock is zero or months to clear is below 3; otherwise Push.
- Customer targets are limited to customers with total six-month purchase value strictly greater than ₹100,000. A customer is assigned to only the first applicable bucket: 1) bought the item including a switch partner; 2) same industry as a buyer; 3) bought another item in the same category. Targets sort by bucket first and purchase value descending within each bucket.

## Before submitting

- Test the app with the supplied data.
- Review every data-quality warning and make sure the handling choices match your written case-study answers.
- Do not include passwords, API keys or private credentials in the repository.
- Deploy using a hosting account you control and confirm the public link works in a private/incognito browser.
