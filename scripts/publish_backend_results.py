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
    backend_path = os.path.join(project_root, "REPORTS", "backend", "backend saati care result.xlsx")
    backend_summary, backend_details = parse_report(backend_path)
    
    markdown_output = []
    markdown_output.append("# ⚙️ Backend Test Verification Dashboard\n")
    if backend_summary:
        markdown_output.append("## ⚙️ Backend Test Verification Summary")
        markdown_output.append("| Metric | Value |")
        markdown_output.append("|---|---|")
        markdown_output.append(f"| **Test Suite** | {backend_summary.get('Test Suite', 'Backend Tests')} |")
        markdown_output.append(f"| **Total Test Cases** | {backend_summary.get('Total Tests', len(backend_details))} |")
        markdown_output.append(f"| **Passed** | ✅ {backend_summary.get('Passed', 'N/A')} |")
        markdown_output.append(f"| **Failed** | ❌ {backend_summary.get('Failed', 'N/A')} |")
        markdown_output.append(f"| **Pass Rate** | **{backend_summary.get('Pass Rate %', 'N/A')}%** |")
        markdown_output.append(f"| **Duration** | {backend_summary.get('Duration (sec)', 'N/A')} sec |")
        markdown_output.append(f"| **Timestamp** | {backend_summary.get('End Time', 'N/A')} |")
        markdown_output.append("\n")
    
    if backend_details:
        markdown_output.append(f"### 🔐 Backend Test Cases Detail Breakdowns")
        markdown_output.append(f"<details><summary>Click to view all Backend Test Cases ({len(backend_details)} tests)</summary>\n")
        markdown_output.append("| No. | Category | Test Name | Status |")
        markdown_output.append("|---|---|---|---|")
        for r in backend_details:
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
