import duckdb
import pandas as pd

out = "output/import_ready.csv"
types = "data/Asset_Types.xlsx"

def check(name, sql):
    result = duckdb.sql(sql).fetchall()
    print(("PASS  " if not result else "FAIL  ") + name)
    for row in result[:10]:
        print("      ", row)

# 1. No serial number appears twice
check("No duplicate serial numbers", f"""
    SELECT serial_number, COUNT(*) FROM read_csv('{out}', all_varchar=true)
    WHERE serial_number IS NOT NULL
    GROUP BY serial_number HAVING COUNT(*) > 1
""")

# 2. Condition is 1 to 5 or blank
check("Condition is 1 to 5 or blank", f"""
    SELECT serial_number, condition FROM read_csv('{out}', all_varchar=true)
    WHERE condition IS NOT NULL AND condition NOT IN ('1','2','3','4','5')
""")

# 3. Install year is 4 digits or blank
check("Install year is 4 digits or blank", f"""
    SELECT serial_number, install_year FROM read_csv('{out}', all_varchar=true)
    WHERE install_year IS NOT NULL AND NOT regexp_full_match(install_year, '\\d{{4}}')
""")

# 4. Every ready row has no reason, every needs_review row has one
check("Status and reason agree", f"""
    SELECT serial_number, status, review_reason FROM read_csv('{out}', all_varchar=true)
    WHERE (status = 'ready' AND review_reason IS NOT NULL)
       OR (status = 'needs_review' AND review_reason IS NULL)
""")

# 5. Every ready row has an asset type on FMX's list
fmx_types = pd.read_excel(types, dtype=str)   # DuckDB can query this pandas table by its name
check("Ready rows use FMX asset types", f"""
    SELECT o.serial_number, o.asset_type FROM read_csv('{out}', all_varchar=true) o
    WHERE o.status = 'ready'
      AND o.asset_type NOT IN (SELECT asset_type FROM fmx_types)
""")
