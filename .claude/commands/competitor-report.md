Generate a competitive intelligence brief for KIABI Argentina and save it to Google Docs.

KIABI positioning: affordable fashion for the whole family, European style and reliable quality.
Primary categories: family baskets, kids, basics, denim, everyday fashion.
Key differentiator vs competitors: wider family proposition (vs Cheeky kids-only, vs Cuesta Blanca young-women-only).

## Competitors to research
- Zara (Argentina) — premium fast fashion benchmark
- Indian (indian.ar) — main price competitor
- Renner — family fashion, price positioning
- Le Utthe — women's fashion
- Cuesta Blanca — young women's fashion
- Cheeky — kids specialization

## Steps

### 1. Research each competitor
For each competitor, run a web search covering:
- Current promotions or sales
- New collections or launches in the past 2 weeks
- Pricing signals (are they running discounts? what %?)
- Any messaging or positioning shifts
- Social media campaigns or notable activity

Use search queries like:
- "[Competitor] Argentina promociones [current month year]"
- "[Competitor] Argentina nueva colección [current month year]"
- "[Competitor] Argentina precios rebajas"

### 2. Write the intelligence brief
Structure the report as follows:

```
KIABI ARGENTINA — COMPETITOR INTELLIGENCE BRIEF
Period: [date range]
Generated: [today's date]

EXECUTIVE SUMMARY
[2-3 sentences: biggest moves this period and top implication for KIABI]

---

COMPETITOR SNAPSHOTS

[Competitor Name]
• Price posture: [premium / mid / value / discounting]
• Active promotions: [description or "none detected"]
• New launches: [description or "none detected"]
• Messaging angle: [what they're emphasizing]
• KIABI implication: [one sentence on what this means for KIABI]

[Repeat for each competitor]

---

STRATEGIC NOTES FOR KIABI
• [2-4 actionable observations for the Argentina launch]
```

### 3. Save to Google Docs
Write the full report text to stdin of the docs helper:
```bash
cat > /tmp/kiabi_competitor_report.txt << 'ENDREPORT'
[PASTE FULL REPORT TEXT HERE]
ENDREPORT
cd /Users/mm/Documents/Projects/Kiabi && python3 agents/docs.py append < /tmp/kiabi_competitor_report.txt
```

### 4. Confirm
Print the Google Doc URL returned by the command.
Also print the Executive Summary to the chat so the user sees the key findings immediately.
