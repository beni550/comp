# CLAUDE.md — Product Catalog Repository

## Overview

This repository manages the **product category taxonomy** for a retail store. The goal is to ensure all products in the store system (managed via "קומקס") are properly classified with their full 4-level category hierarchy:

```
תחום (Domain) → מחלקה (Department) → קבוצה (Group) → קבוצת משנה (Subgroup)
```

---

## Files in This Repository

### `קבוצות קומקס.xlsx` — Taxonomy / Category Master Table
The authoritative category structure. Contains ~967 rows representing all valid category paths.

**Columns:**
| Column | Name | Description |
|--------|------|-------------|
| A | תחום (ID) | Domain numeric code |
| B | שם תחום | Domain name |
| C | מחלקה (ID) | Department numeric code |
| D | שם מחלקה | Department name |
| E | קבוצה (ID) | Group numeric code |
| F | שם קבוצה | Group name |
| G | ק.משנה (ID) | Subgroup numeric code |
| H | שם ק.משנה | Subgroup name |
| I | פריטים | Number of items currently in subgroup |

**Domains (תחום) in the taxonomy:**
- `מזון יבש` — Dry food (canned goods, pasta, spices, sweets, etc.)
- `מצוננים` — Refrigerated (meat, fish, dairy, deli, prepared food)
- `משקאות` — Beverages
- `ניקיון` — Cleaning & hygiene products
- `פארם` — Pharmacy / personal care
- `פירות וירקות` — Fruit & vegetables (incl. nuts/seeds, fresh produce)
- `קפואים` — Frozen foods
- `NF` — Non-Food (household items, textiles, electronics, holidays, leisure)

---

### `פריטים ללא קבוצת משנה.xlsx` — Products Needing Categorization
Contains **2,199 products** that need their category hierarchy assigned or completed.

**Columns:**
| Column | Name | Description |
|--------|------|-------------|
| A | פריט | Item barcode / ID |
| B | שם פריט | Product name (Hebrew) |
| C | תאור מחלקת על | Domain name (may be empty) |
| D | שם מחלקה | Department name (may be empty) |
| E | שם קבוצה | Group name (may be empty) |
| F | שם קבוצת משנה | Subgroup name (TARGET — always empty, must be filled) |
| G | ספק | Supplier ID |
| H | שם ספק | Supplier name |

**Current state of products:**
- **0 products** have all 4 levels filled
- **555 products** have domain+dept (and sometimes group) but missing subgroup
- **1,644 products** have no category data at all

---

## Classification Task

The primary task is to update all 2,199 products with the correct category hierarchy from the taxonomy file.

### Priority Groups (products with partial data, 555 items):

| Count | Domain | Dept | Group | Action needed |
|-------|--------|------|-------|---------------|
| 144 | פירות וירקות | פיצוחים | — | Assign group + subgroup (needs new subgroups) |
| 79 | מצוננים | דגים טריים | דגים טריים | Assign subgroup |
| 69 | מצוננים | קצביה בקר טרי | חלקי בשר טרי | Assign subgroup |
| 53 | NF | טקסטיל | — | Assign group + subgroup |
| 33 | מצוננים | קצביה עופות טריים | הודו טרי | Assign subgroup |
| 27 | מצוננים | קצביה עופות טריים | חלקי עוף טרי | Assign subgroup |
| 21 | פירות וירקות | ירקות | פטריות ונבטים | Assign subgroup |

### Products needing NEW categories to be created:

1. **Nuts/Seeds (פיצוחים)** — 144 products under "פיצוחים" dept need subgroups created:
   - `אגוזים` (walnuts, Brazil nuts, hazelnuts)
   - `בוטנים` (peanuts)
   - `גרעינים` (sunflower seeds, watermelon seeds, pumpkin seeds)
   - `שקדים` (almonds)
   - `קשיו` (cashews)
   - `פיסטוק` (pistachios)
   - `פקאן ומקדמיה` (pecans and macadamia)
   - `חומוס ודגנים קלויים` (roasted chickpeas and grains)
   - `תערובות פיצוחים` (mixed nuts)
   - `חטיפי פיצוחים` (nut snacks — קבוקים, קרנצוס, etc.)
   - `ממתקי פיצוחים` (nut-based candies)

2. **Fresh spices/herbs (תבלינים טריים)** — physically located in the nuts area but botanically belong under:
   - `פירות וירקות → ירקות → עשבי תיבול טריים` (new subgroup to create)
   - Examples: חזרת (horseradish), לוף (arum)

3. **Purim/Holiday accessories** — already in taxonomy under:
   - `NF → חגים → פורים → אביזרים למשלוחי מנות`
   - `NF → חגים → פורים → תחפושות ואביזרים`
   - `NF → חגים → פורים → מגילות אסתר`

4. **Religious items** (ציצית, טלית, פתילים) — need a new category, likely:
   - `NF → יודאיקה` (new department to create) or `NF → פנאי → ספרי קודש`

---

## Scripts

### `categorize_products.py`
Auto-classifies products using rule-based Hebrew keyword matching against the taxonomy.

**Usage:**
```bash
python3 categorize_products.py
```

**Output:**
- `פריטים_מעודכנים.xlsx` — Updated products file with all category fields filled
- Console summary of: auto-classified, flagged for review, suggested new categories

**Classification logic:**
1. Exact match on existing dept/group → find best subgroup from taxonomy
2. Keyword match on product name → assign domain/dept/group/subgroup
3. Unknown products → flagged as `לבדיקה` (needs review) with best-guess suggestion

---

## Development Notes

- **Language**: All product/category data is in Hebrew (RTL)
- **Encoding**: Excel files use UTF-8 compatible encoding via openpyxl
- **VAT note**: Some categories have "ללא מע"מ" (VAT-exempt) variants for fresh produce
- **Supplier mapping**: The `ספק` column can provide hints about product type (e.g., supplier 550013 = nuts supplier גרעיני הבית)

---

## Key Rules & Conventions

1. **Never invent category IDs** — only use IDs from the taxonomy file
2. **When creating new subgroups**, use the same numeric pattern (e.g., if group is `150`, subgroups are `1501`, `1502`, etc.)
3. **Fresh items** sold alongside nuts (herbs, spices) belong under `פירות וירקות`, not `פיצוחים`
4. **Holiday items** (פורים, חנוכה, etc.) → always `NF → חגים`
5. **Religious/Judaica items** → needs new `NF → יודאיקה` department
6. **"פריט חדש/ לא קיים במגוון"** items → skip or mark as discontinued
