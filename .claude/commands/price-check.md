Check current prices for all tracked products and update the KIABI Price Tracker sheet.

## Steps

### 1. Load tracked products
```bash
cd /Users/mm/Documents/Projects/Kiabi && python3 agents/sheets.py list
```
This returns a JSON array. Each item has: row, name, competitor, url, current_price (the stored price).

If the list is empty, tell the user to add products first with /price-add.

### 2. Fetch and extract prices
For each product in the list, use WebFetch on the product URL. From the fetched content:
- Look for price in structured data (JSON-LD, meta tags), visible price text, or ARS/$ amounts near the product name
- Extract the price in Argentine Pesos. Always prefix with "ARS" (e.g. "ARS 13.425", "ARS 37.425")
- If the page shows a discounted price, use the discounted (final) price
- If the page cannot be fetched or price is not found, use "N/A"

### 3. Calculate deltas
For each product:
- If current_price (stored) is empty or "N/A" → delta_pct = "new"
- Otherwise, parse both prices numerically (strip ARS, $, dots, commas) and calculate:
  delta_pct = ((new - old) / old * 100) formatted as "+5.2%" or "-3.1%"
- If either price is N/A → delta_pct = "N/A"

### 4. Write updates to sheet
Build the update payload as JSON — one object per product regardless of whether price changed:
```json
[
  {"row": 2, "current_price": "ARS 15.999", "delta_pct": "+5.2%"},
  {"row": 3, "current_price": "ARS 37.425", "delta_pct": "new"}
]
```

Write to a temp file and pipe to the update command:
```bash
cat > /tmp/kiabi_price_updates.json << 'ENDJSON'
[PASTE JSON PAYLOAD HERE]
ENDJSON
cd /Users/mm/Documents/Projects/Kiabi && python3 agents/sheets.py update < /tmp/kiabi_price_updates.json
```

### 5. Print summary
After updating, print a markdown table showing:
| Product | Competitor | Previous | Current | Delta |
For any product with a delta (not "new" or "N/A"), bold that row and note it as a price change.

Note: Some pages may use JavaScript rendering. If a price cannot be extracted, log "N/A" and continue — do not retry.
