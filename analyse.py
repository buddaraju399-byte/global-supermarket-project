import sys, statistics

data_file = r".bob/tmp/office-dumps/global_superstore_2016-01ade3ef/data.txt"

col_map = {}
sales_by_region = {}
sales_by_category = {}
all_sales = []
seen_order_ids = set()
duplicate_order_ids = set()
missing_postal = 0
total_rows = 0

with open(data_file, encoding="utf-8") as f:
    for line in f:
        line = line.rstrip("\n")
        if not line.strip():
            continue
        tab_idx = line.find("\t")
        if tab_idx == -1:
            continue
        header = line[:tab_idx]
        cell_part = line[tab_idx+1:]

        import re
        row_match = re.search(r"Orders/row\[(\d+)\]", header)
        if not row_match:
            continue
        row_num = int(row_match.group(1))

        record = {}
        for cell in cell_part.split("\t"):
            eq_idx = cell.find("=")
            if eq_idx == -1:
                continue
            key = cell[:eq_idx]
            val = cell[eq_idx+1:]
            col = key.rstrip("0123456789")
            record[col] = val

        if row_num == 1:
            col_map = dict(record)
            continue

        total_rows += 1

        order_id = record.get("B", "")
        if order_id in seen_order_ids:
            duplicate_order_ids.add(order_id)
        else:
            seen_order_ids.add(order_id)

        postal = record.get("I", "")
        if not postal or postal.strip() == "":
            missing_postal += 1

        region = record.get("M", "Unknown")
        category = record.get("P", "Unknown")
        try:
            sales = float(record.get("S", 0) or 0)
        except ValueError:
            sales = 0

        sales_by_region[region] = sales_by_region.get(region, 0) + sales
        sales_by_category[category] = sales_by_category.get(category, 0) + sales
        all_sales.append(sales)

all_sales_sorted = sorted(all_sales)
mean_val = statistics.mean(all_sales)
median_val = statistics.median(all_sales)
mode_val = statistics.mode(all_sales)

print("=== COLUMN MAP ===")
for k, v in col_map.items():
    print(f"  {k}: {v}")

print("\n=== SUMMARY ===")
print(f"Total data rows: {total_rows}")
print(f"Duplicate Order IDs (count): {len(duplicate_order_ids)}")
print(f"Rows with missing Postal Code: {missing_postal}")
print(f"\nMean Sales: {mean_val:.4f}")
print(f"Median Sales: {median_val:.4f}")
print(f"Mode Sales: {mode_val:.4f}")

n = len(all_sales_sorted)
top10_threshold = all_sales_sorted[n - 10] if n >= 10 else 0
print(f"Top-10 threshold (min of top 10): {top10_threshold:.4f}")

print("\n=== SALES BY REGION ===")
for r, s in sorted(sales_by_region.items(), key=lambda x: -x[1]):
    print(f"  {r}: {s:.2f}")

print("\n=== SALES BY CATEGORY ===")
for c, s in sorted(sales_by_category.items(), key=lambda x: -x[1]):
    print(f"  {c}: {s:.2f}")
