import pandas as pd


def generate_excel_report(records: list, output_path: str) -> None:
    print(f"\n[REPORTER] generating report with {len(records)} records...")

    successful = [r for r in records if not r.get("_error")]
    failed = [r for r in records if r.get("_error")]

    for r in successful:
        r.pop("_error", None)

    df_success = pd.DataFrame(successful)
    df_failed = pd.DataFrame(failed)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_success.to_excel(writer, sheet_name="Results", index=False)
        df_failed.to_excel(writer, sheet_name="Failures", index=False)

    print(f"[REPORTER] report saved to {output_path}")
    print(f"[REPORTER] {len(successful)} successful, {len(failed)} failed")
