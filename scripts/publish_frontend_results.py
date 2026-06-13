import os
import openpyxl

def parse_report(filepath):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws_summary = wb['Summary']
        rows = list(ws_summary.values)
        headers = [str(h) for h in rows[0]]
        data = rows[1]
        summary_dict = dict(zip(headers, data))
        
        ws_details = wb['Test Details']
        detail_rows = list(ws_details.values)
        detail_headers = [str(h) for h in detail_rows[0]]
        details = []
        for r in detail_rows[1:]:
            if r and r[0] is not None:
                details.append(dict(zip(detail_headers, r)))
        return summary_dict, details
    except Exception as e:
        print(f"Error parsing {filepath}: {e}")
        return {}, []

def main():
    import sys
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    frontend_path = os.path.join(project_root, "REPORTS", "frontend", "frontend saathi care result.xlsx")
    frontend_summary, frontend_details = parse_report(frontend_path)
    
    markdown_output = []
    markdown_output.append("# 💻 Frontend Test Verification Dashboard\n")
    if frontend_summary:
        markdown_output.append("## 💻 Frontend Test Suite Summary")
        markdown_output.append("| Metric | Value |")
        markdown_output.append("|---|---|")
        markdown_output.append(f"| **Test Suite** | {frontend_summary.get('Test Suite', 'Frontend Tests')} |")
        markdown_output.append(f"| **Total Test Cases** | {frontend_summary.get('Total Tests', len(frontend_details))} |")
        markdown_output.append(f"| **Passed** | ✅ {frontend_summary.get('Passed', 'N/A')} |")
        markdown_output.append(f"| **Failed** | ❌ {frontend_summary.get('Failed', 'N/A')} |")
        markdown_output.append(f"| **Pass Rate** | **{frontend_summary.get('Pass Rate %', 'N/A')}%** |")
        markdown_output.append(f"| **Duration** | {frontend_summary.get('Duration (sec)', 'N/A')} sec |")
        markdown_output.append(f"| **Timestamp** | {frontend_summary.get('End Time', 'N/A')} |")
        markdown_output.append("\n")
    
    if frontend_details:
        markdown_output.append(f"### 📋 Frontend Test Cases Detail Breakdowns")
        markdown_output.append(f"<details><summary>Click to view all Frontend Test Cases ({len(frontend_details)} tests)</summary>\n")
        markdown_output.append("| No. | Category | Test Name | Status |")
        markdown_output.append("|---|---|---|---|")
        for r in frontend_details:
            status_emoji = "✅ PASSED" if str(r.get("Status")).upper() == "PASSED" else "❌ FAILED"
            markdown_output.append(f"| {r.get('No.', '-')} | {r.get('Category', '-')} | `{r.get('Test Name', '-')}` | {status_emoji} |")
        markdown_output.append("\n</details>\n")
    
    full_markdown = "\n".join(markdown_output)
    summary_file = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary_file:
        with open(summary_file, "a", encoding="utf-8") as f:
            f.write(full_markdown)
    else:
        print(full_markdown)

if __name__ == "__main__":
    main()
