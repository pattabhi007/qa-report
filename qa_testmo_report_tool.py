
import pandas as pd
import matplotlib.pyplot as plt
import glob
import os
from io import StringIO

INPUT_FOLDER = "test_runs"
OUTPUT_EXCEL = "qa_automation_report.xlsx"
OUTPUT_CHART = "qa_pass_fail_trend.png"

def extract_results(file_path):
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    marker = '"Test key","Test ID","Test"'
    idx = content.find(marker)
    if idx == -1:
        return None

    df = pd.read_csv(StringIO(content[idx:]))

    if "Status" not in df.columns:
        return None

    total = len(df)
    passed = (df["Status"].astype(str).str.lower() == "passed").sum()
    failed = (df["Status"].astype(str).str.lower() == "failed").sum()
    skipped = (df["Status"].astype(str).str.lower().isin(["skipped","blocked"])).sum()

    run_id = os.path.basename(file_path).replace(".csv","")

    return {
        "Run": run_id,
        "Total": total,
        "Passed": passed,
        "Failed": failed,
        "Skipped": skipped,
        "PassRate": round((passed/total)*100,2) if total else 0
    }

def main():
    files = glob.glob(os.path.join(INPUT_FOLDER,"*.csv"))
    summary = []

    for f in files:
        res = extract_results(f)
        if res:
            summary.append(res)

    if not summary:
        print("No valid results found.")
        return

    df = pd.DataFrame(summary).sort_values("Run")

    print(df)

    df.to_excel(OUTPUT_EXCEL, index=False)

    plt.figure()
    plt.plot(df["Run"], df["Passed"], marker="o", label="Passed")
    plt.plot(df["Run"], df["Failed"], marker="o", label="Failed")
    plt.xticks(rotation=45)
    plt.ylabel("Test Count")
    plt.title("QA Automation Trend")
    plt.legend()
    plt.tight_layout()
    plt.savefig(OUTPUT_CHART)

    print("Report generated:")
    print(OUTPUT_EXCEL)
    print(OUTPUT_CHART)

if __name__ == "__main__":
    main()
