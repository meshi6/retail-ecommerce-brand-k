Add a product to the KIABI price tracker sheet.

Ask the user for the following if not already provided in the command arguments:
1. **Product name** — descriptive label (e.g. "Jean básico mujer azul")
2. **Competitor** — one of: Zara, Indian, Renner, Le Utthe, Cuesta Blanca, Cheeky
3. **Product URL** — the exact product page URL to track

Once you have all three, run:

```bash
cd /Users/mm/Documents/Projects/Kiabi && python3 agents/sheets.py add "[name]" "[competitor]" "[url]"
```

Replace the placeholders with the actual values the user provided. Quote each argument.

After running, confirm success and show the total count of tracked products from the command output.
