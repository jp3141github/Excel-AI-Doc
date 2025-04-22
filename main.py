import os
from insight_engine import (
    load_excel_sheet, generate_summary_stats,
    detect_quality_issues, build_gpt_prompt,
    query_openai_insights, export_report_to_word
)

# === CONFIGURATION ===
# Set working directory
os.chdir(r"C:\Users\QLCY\OneDrive - Direct Line Group\Documents\Excel Spreadsheet AI Documentation Tool\gt ifoa v0.1 alt tool python - FDP")

# Set your Excel file path and sheet
file_path = "Q1 2025 WD3_FinalforFDPOutputIncludesSection-2025.04.04-17.27-06732.xlsx"
sheet_name = "Sheet1"

# === EXECUTION ===
# Load Excel content
with open(file_path, "rb") as f:
    content = f.read()

# Process dataset
df = load_excel_sheet(content, sheet_name=sheet_name)
stats_info = generate_summary_stats(df)
issues = detect_quality_issues(df)
prompt = build_gpt_prompt(stats_info["stats"], stats_info["structure"], issues)
insights = query_openai_insights(prompt)

# ✅ Show GPT-4 output in console
print("\n🧠 GPT-4 Insight Output Preview:\n")
print(insights)
print("\n📦 Saving to Word document...")

# Export report
doc_path = export_report_to_word(prompt, stats_info["stats"], stats_info["structure"], issues, insights)

# Output result
print("\n✅ GPT-4 Insights Generated and Report Saved!")
print(f"📄 File saved at: {doc_path}")
