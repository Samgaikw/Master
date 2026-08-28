"""
SQL Action Generator - Version 1

Purpose:
    Read SQL actions from the Excel "Extracted_queries" sheet and update
    PySpark job files.

Fixed business rules:
    - project_name must be EDW_PRD_BID
    - INSERT is the only action that remains executable
    - DELETE/BEFORE SQL is added but commented
    - EXECUTE/AFTER SQL is added but commented
    - Existing write.parquet() lines are commented
    - INSERT is placed after createOrReplaceTempView(...)
    - DELETE/BEFORE is placed before createOrReplaceTempView(...)
    - EXECUTE/AFTER is placed after the INSERT block
    - Original .py files are NOT overwritten; *_generated.py is created

Excel expected columns:
    project_name
    JobName
    Connector
    SQL Type
    SQL Statement

Usage:
    1. Put this file in the folder containing your .py jobs.
    2. Put the Excel file in the same folder, or change EXCEL_FILE below.
    3. Run:
           python sql_action_generator.py
    4. Generated files will be placed in ./generated/
"""

from pathlib import Path
import re
import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

PROJECT = "EDW_PRD_BID"

EXCEL_FILE = Path("Master_sheet.xlsx")
SHEET_NAME = "Extracted_queries"

OUTPUT_DIR = Path("generated")


# ============================================================
# HELPERS
# ============================================================

def clean(value):
    """Return a normalized string."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_sql(sql):
    """Normalize SQL for comparison/search."""
    sql = clean(sql)
    sql = sql.replace('"""', '"')
    sql = re.sub(r"\s+", " ", sql)
    return sql.strip().lower()


def sql_words(sql):
    """Extract useful SQL tokens such as DSLink/view/table names."""
    return set(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", clean(sql)))


def comment_line(line):
    """Comment a Python line without disturbing indentation."""
    if line.lstrip().startswith("#"):
        return line

    indent = line[:len(line) - len(line.lstrip())]
    return indent + "# " + line[len(indent):]


def comment_block(lines):
    """Comment every non-empty line in a block."""
    return [comment_line(line) if line.strip() else line for line in lines]


def find_job_name(code):
    """Read job_name = '...' from the Python file."""
    patterns = [
        r'^\s*job_name\s*=\s*["\']([^"\']+)["\']',
        r'^\s*job_name\s*=\s*f["\']([^"\']+)["\']',
    ]

    for pattern in patterns:
        match = re.search(pattern, code, flags=re.MULTILINE)
        if match:
            return match.group(1).strip()

    return None


def find_view_names(code):
    """
    Return all createOrReplaceTempView names and their line numbers.
    """
    matches = []
    pattern = r'createOrReplaceTempView\s*\(\s*["\']([^"\']+)["\']\s*\)'

    for match in re.finditer(pattern, code):
        line_no = code[:match.start()].count("\n")
        matches.append((line_no, match.group(1)))

    return matches


def find_best_view(code, sql_statement):
    """
    Identify the most likely DSLink temp view referenced by the SQL.
    Example:
        SELECT * FROM DSLink6_validated_df_vw
    """
    views = find_view_names(code)
    if not views:
        return None

    tokens = sql_words(sql_statement)

    candidates = []
    for line_no, view_name in views:
        view_tokens = sql_words(view_name)
        score = len(tokens.intersection(view_tokens))

        # Strong match if exact view name occurs in SQL.
        if view_name.lower() in normalize_sql(sql_statement):
            score += 100

        candidates.append((score, line_no, view_name))

    candidates.sort(reverse=True)

    if candidates and candidates[0][0] > 0:
        return candidates[0]

    return None


def find_connector_blocks(lines):
    """
    Find connector activity blocks.

    A block starts at:
        # Oracle Connector Activity: ...
        # ODBC Connector Activity: ...
        # DB2 Connector Activity: ...

    and ends immediately before the next connector activity or EOF.
    """
    starts = []

    pattern = re.compile(
        r"^\s*#\s*(?:Oracle|ODBC|DB2|SQL Server|MSSQL).*?Connector Activity:",
        re.IGNORECASE
    )

    for i, line in enumerate(lines):
        if pattern.search(line):
            starts.append(i)

    blocks = []

    for idx, start in enumerate(starts):
        end = starts[idx + 1] if idx + 1 < len(starts) else len(lines)
        blocks.append((start, end))

    return blocks


def find_matching_connector_block(lines, sql_statement):
    """
    Try to identify the connector block using SQL/table/view tokens.

    Preference:
      1. Exact DSLink validated view in SQL
      2. Target table in SQL appearing in the block
      3. Unique meaningful token overlap
    """
    sql_norm = normalize_sql(sql_statement)
    tokens = sql_words(sql_statement)

    blocks = find_connector_blocks(lines)

    if not blocks:
        return None

    # First: exact DSLink/view match.
    for start, end in blocks:
        block_text = "\n".join(lines[start:end])
        view_match = re.search(
            r'DSLink[A-Za-z0-9_]*_validated_df_vw',
            sql_statement,
            re.IGNORECASE
        )
        if view_match and view_match.group(0).lower() in block_text.lower():
            return start, end

    # Second: score token overlap.
    scored = []

    for start, end in blocks:
        block_text = "\n".join(lines[start:end])
        block_tokens = sql_words(block_text)

        score = len(tokens.intersection(block_tokens))

        # Table name immediately after INSERT INTO / DELETE FROM is useful.
        table_match = re.search(
            r'(?:insert\s+into|delete\s+from)\s+([A-Za-z0-9_.]+)',
            sql_norm,
            re.IGNORECASE
        )
        if table_match:
            table = table_match.group(1).lower()
            if table in block_text.lower():
                score += 50

        scored.append((score, start, end))

    scored.sort(reverse=True)

    if scored and scored[0][0] >= 2:
        # Reject ambiguous weak matches.
        if len(scored) == 1 or scored[0][0] > scored[1][0]:
            return scored[0][1], scored[0][2]

    return None


def format_sql_block(sql_statement, executable):
    """
    Create the standard db2_sql + spark.sql block.
    """
    sql_statement = clean(sql_statement)

    lines = [
        'db2_sql = f"""' + sql_statement + '"""',
        'spark.sql(db2_sql)',
    ]

    if not executable:
        lines = comment_block(lines)

    return lines


def find_insert_anchor(block_lines):
    """
    Find createOrReplaceTempView(...) line.
    INSERT goes immediately after this line.
    """
    for i, line in enumerate(block_lines):
        if "createOrReplaceTempView" in line:
            return i

    return None


def find_existing_insert_end(block_lines, start_index):
    """
    If an INSERT block already exists immediately after the temp view,
    return its end index so we can avoid duplicate generation.
    """
    i = start_index + 1

    while i < len(block_lines) and not block_lines[i].strip():
        i += 1

    if i >= len(block_lines):
        return None

    if "db2_sql" in block_lines[i]:
        end = i + 1

        while end < len(block_lines) and (
            "spark.sql(db2_sql)" in block_lines[end]
            or not block_lines[end].strip()
        ):
            end += 1

        return end

    return None


def generate_for_job(py_path, excel_rows):
    """
    Update one Python job using the matching Excel rows.
    """
    original = py_path.read_text(encoding="utf-8")
    job_name = find_job_name(original)

    if not job_name:
        return False, ["Could not find job_name = ..."], None

    rows = excel_rows[
        excel_rows["JobName"].str.strip().str.upper() == job_name.upper()
    ].copy()

    if rows.empty:
        return False, [f"No Excel rows found for {job_name}"], None

    lines = original.splitlines(keepends=True)
    report = []

    # --------------------------------------------------------
    # Comment existing write.parquet() lines.
    # --------------------------------------------------------
    for i, line in enumerate(lines):
        if ".write.parquet(" in line and not line.lstrip().startswith("#"):
            lines[i] = comment_line(line)
            report.append(f"Line {i+1}: commented write.parquet()")

    # --------------------------------------------------------
    # Process actions.
    # --------------------------------------------------------
    # Work from bottom to top so inserted lines don't invalidate indexes.
    actions = []

    for _, row in rows.iterrows():
        sql_type = clean(row["SQL Type"]).upper()
        sql_statement = clean(row["SQL Statement"])

        if not sql_statement:
            continue

        actions.append((sql_type, sql_statement))

    # Process each connector block independently.
    for sql_type, sql_statement in actions:
        match = find_matching_connector_block(lines, sql_statement)

        if not match:
            report.append(
                f"WARNING: could not identify connector block for {sql_type}: "
                f"{sql_statement[:100]}"
            )
            continue

        block_start, block_end = match
        block = lines[block_start:block_end]

        anchor = find_insert_anchor(block)

        if anchor is None:
            report.append(
                f"WARNING: no createOrReplaceTempView() found for {sql_type}: "
                f"{sql_statement[:100]}"
            )
            continue

        # Convert block-relative index to file index.
        anchor_file_index = block_start + anchor

        if sql_type in {"INSERT"}:
            # INSERT is the only executable action.
            existing_end = find_existing_insert_end(
                lines[block_start:block_end],
                anchor
            )

            insert_lines = format_sql_block(
                sql_statement,
                executable=True
            )

            # Add a clear generator marker.
            insert_lines = [
                "    # Generated SQL action: INSERT\n"
            ] + [
                ("    " + x.lstrip()) if x.strip() else x
                for x in insert_lines
            ]

            # Put INSERT directly after temp view.
            insert_at = anchor_file_index + 1

            # Avoid exact duplicate SQL.
            current_text = "\n".join(lines)
            if sql_statement.lower() in current_text.lower():
                report.append(
                    f"SKIPPED duplicate INSERT: {sql_statement[:100]}"
                )
            else:
                lines[insert_at:insert_at] = insert_lines
                report.append(
                    f"ADDED active INSERT after createOrReplaceTempView: "
                    f"{sql_statement[:100]}"
                )

        elif sql_type in {"BEFORE", "DELETE"}:
            # DELETE must remain commented.
            delete_lines = format_sql_block(
                sql_statement,
                executable=False
            )

            delete_lines = [
                "    # Generated SQL action: DELETE (intentionally commented)\n"
            ] + [
                ("    " + x.lstrip()) if x.strip() else x
                for x in delete_lines
            ]

            # BEFORE = immediately before createOrReplaceTempView.
            insert_at = anchor_file_index

            current_text = "\n".join(lines)
            if sql_statement.lower() in current_text.lower():
                report.append(
                    f"SKIPPED duplicate DELETE: {sql_statement[:100]}"
                )
            else:
                lines[insert_at:insert_at] = delete_lines
                report.append(
                    f"ADDED commented DELETE before createOrReplaceTempView: "
                    f"{sql_statement[:100]}"
                )

        elif sql_type in {"AFTER", "EXECUTE"}:
            # EXECUTE must remain commented.
            execute_lines = format_sql_block(
                sql_statement,
                executable=False
            )

            execute_lines = [
                "    # Generated SQL action: EXECUTE (intentionally commented)\n"
            ] + [
                ("    " + x.lstrip()) if x.strip() else x
                for x in execute_lines
            ]

            # Find the INSERT block after the temp view.
            # Re-scan because previous insertions may have shifted indexes.
            block_end = min(block_end + 20, len(lines))
            insert_anchor = None

            for i in range(anchor_file_index + 1, block_end):
                if "spark.sql(db2_sql)" in lines[i]:
                    insert_anchor = i
                    break

            if insert_anchor is None:
                report.append(
                    f"WARNING: could not find INSERT execution line for AFTER: "
                    f"{sql_statement[:100]}"
                )
                continue

            insert_at = insert_anchor + 1

            current_text = "\n".join(lines)
            if sql_statement.lower() in current_text.lower():
                report.append(
                    f"SKIPPED duplicate EXECUTE: {sql_statement[:100]}"
                )
            else:
                lines[insert_at:insert_at] = execute_lines
                report.append(
                    f"ADDED commented EXECUTE after INSERT: "
                    f"{sql_statement[:100]}"
                )

        else:
            report.append(
                f"WARNING: unsupported SQL Type '{sql_type}' - skipped"
            )

    new_code = "".join(lines)

    output_path = OUTPUT_DIR / f"{py_path.stem}_generated.py"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(new_code, encoding="utf-8")

    return True, report, output_path


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 70)
    print("SQL ACTION GENERATOR")
    print("=" * 70)
    print(f"Project filter : {PROJECT}")
    print(f"Excel          : {EXCEL_FILE}")
    print(f"Sheet          : {SHEET_NAME}")
    print()

    if not EXCEL_FILE.exists():
        print(f"ERROR: Excel file not found: {EXCEL_FILE}")
        return

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    required = {
        "project_name",
        "JobName",
        "SQL Type",
        "SQL Statement",
    }

    missing = required - set(df.columns)
    if missing:
        print(f"ERROR: Missing Excel columns: {sorted(missing)}")
        return

    # Always use EDW_PRD_BID.
    df = df[
        df["project_name"].astype(str).str.strip().str.upper()
        == PROJECT.upper()
    ].copy()

    print(f"Excel rows for {PROJECT}: {len(df)}")
    print()

    py_files = [
        p for p in Path(".").glob("*.py")
        if p.name != Path(__file__).name
        and not p.name.endswith("_generated.py")
    ]

    if not py_files:
        print("No Python job files found in current folder.")
        return

    for py_path in py_files:
        print("-" * 70)
        print(f"JOB FILE: {py_path.name}")

        ok, report, output_path = generate_for_job(py_path, df)

        for message in report:
            print("  " + message)

        if ok and output_path:
            print(f"  OUTPUT: {output_path}")

    print()
    print("=" * 70)
    print("Completed. Original .py files were not overwritten.")
    print("=" * 70)


if __name__ == "__main__":
    main()
