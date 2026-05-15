# Money Helper

`money_helper.py` summarizes bank, credit-card, and other transaction files by category.

It reads one or more CSV or Excel statement files, keeps only transactions inside a required date range, uses the sign of each amount to decide whether the row is income or spending, assigns each row to a category from `category.md`, and prints a final summary.

The script does not write a new output file. It prints the report in the terminal.

## How To Use This Project

First, download this project to your local computer.

Open your terminal window, find or create an empty folder where you want to keep this project, and type:

```bash
git clone https://github.com/XipingGong/money_helper.git
cd money_helper
```

Next, download your own statement files from your bank, credit-card, brokerage, or payroll websites. Save those files on your local computer, preferably in the same folder as `money_helper.py`.

Then, prepare your own `category.md`. The provided `category.md` is only an example for one family. You should revise it so the income names, bills, stores, schools, subscriptions, transfers, and other categories match your own family. When you run `money_helper.py`, the script reads each transaction row and compares the row text with the rules in `category.md`. If a transaction matches one category rule, that transaction is counted in that category. If it does not match any named category, it is counted under `Others`.

After the project, statement files, and `category.md` are ready, you have two ways to use this tool.

### Option 1: Use Python Yourself

If you want to run the script yourself, read the sections below and use a command like this:

```bash
# You need to determine "your_input_statements.csv", dates, etc.
python money_helper.py -i "your_input_statements.csv" --start_date "2026-03-01" --end_date "2026-03-31" --category category.md --print_mode 0
```

This is the direct Python workflow. You control the command, date range, input files, category file, and print mode.

### Option 2: Use A Coding AI Agent

You can also ask a coding AI agent to help run the script, explain the output, or improve `category.md`.

To do this:

1. Install or open the coding AI agent tool you want to use.
2. Open your terminal.
3. Go to this project directory with `cd money_helper`.
4. Start or open the coding AI agent window in this directory.
5. Talk to the coding AI agent in normal language.

Example request:

```text
Please run money_helper.py for March 2026 using all CSV files in this folder, my category.md file, and print_mode 0. Then explain the final summary.
```

Another useful request:

```text
Please run money_helper.py with --print Others and suggest how I should update category.md. Do not change my statement files.
```

## What This Script Is Used For

Use this script when you want to answer questions such as:

- How much money came in during a month?
- How much money went out during a month?
- Which spending categories used the most money?
- Which transactions were not matched by my category rules?
- Do the category totals add up to the final balance?

The final summary includes:

- processed input files
- matched transaction row count
- total input
- total output
- net balance
- category totals
- `Others` totals
- row-count and balance checks

## Input File Requirements

Each input file must be a CSV or Excel statement file.

Supported file types:

- `.csv`
- `.xlsx`
- `.xlsm`

Legacy `.xls` files are not supported. Convert them to `.xlsx` or `.csv` first.

Each file must contain a header row with:

- a date column
- an amount column

The script recognizes these date column names:

- `Date`
- `Transaction Date`
- `Trans. Date`

The script recognizes the amount column name:

- `Amount`

Capitalization does not matter. For example, `amount`, `Amount`, and `AMOUNT` are treated the same.

The sign of the amount matters:

- positive amount: treated as `Input`
- negative amount: treated as `Output`
- zero amount: ignored

Examples:

```text
Amount = 2500.00   -> Input
Amount = -45.18    -> Output
Amount = 0.00      -> ignored
```

Rows are skipped when the date or amount cannot be parsed.

## How The Script Works

The script follows these steps:

1. Read the input file paths, start date, end date, category file, row-print option, and print mode.
2. Expand input file patterns such as `*.csv` into real file names.
3. Read each CSV or Excel file.
4. Find the header row by looking for both a date column and an amount column.
5. Parse each transaction date and amount.
6. Keep only rows whose date is between `--start_date` and `--end_date`, including both dates.
7. Use the amount sign to choose either the `# Input` rules or the `# Output` rules from `category.md`.
8. Match the transaction row against the rules in order.
9. Put unmatched rows into `Others`.
10. Print per-file summaries and a final cross-file summary.

## The Role Of `category.md`

`category.md` tells the script how to group transactions.

It must contain two sections:

```markdown
# Input
- Others

# Output
- Others
```

`# Input` is used for positive amounts.

`# Output` is used for negative amounts.

`Others` is the fallback category. Keep it in both sections so unmatched rows are still counted.

Example:

```markdown
# Input
- 'University of Ge' & XXX,
- 'University of Ge' & XXX,
- 'IRS TREAS',
- Webull
- Others

# Output
- HOME, NSM, MORTGAGE,
- Groceries, Food & Drink, 'COSTCO WHSE', Shopping,
- Bills & Utilities, COMCAST, 'MINT MOBILE', EMC,
- Gas, Automotive, 'COSTCO GAS',
- Travel
- Webull
- Others
```

## How Category Rules Match Rows

For each transaction, the script combines the row values into normalized text and searches that text.

Rules are checked from top to bottom. The first matching named rule wins. If no named rule matches, the row goes to `Others`.

Comma-separated terms mean "match any of these terms".

Example:

```markdown
- HOME, NSM, MORTGAGE,
```

This means a negative transaction row goes into this output category if the row text contains `HOME`, `NSM`, or `MORTGAGE`.

Use `&` when multiple terms must appear in the same row.

Example:

```markdown
- 'University of Ge' & XXX,
```

This means a positive transaction row must contain both `University of Ge` and `XXX`.

Use quotes around phrases that contain spaces.

Example:

```markdown
- 'COSTCO WHSE'
```

Without quotes, a term with spaces is invalid.

Use `OR` when you want to make the choice explicit.

Example:

```markdown
- COMCAST OR EMC,
```

Do not mix `&` and `OR` in the same rule.

## Basic Command

Run the script like this:

```bash
python money_helper.py -i "*.csv" --start_date "2026-03-01" --end_date "2026-03-31" --category category.md
```

Meaning:

- `python money_helper.py`: run the script.
- `-i "*.csv"`: use every CSV file in the current folder.
- `--start_date "2026-03-01"`: include transactions on or after March 1, 2026.
- `--end_date "2026-03-31"`: include transactions on or before March 31, 2026.
- `--category category.md`: use `category.md` as the category rule file.

You can also write `--input` instead of `-i`:

```bash
python money_helper.py --input "*.csv" --start_date "2026-03-01" --end_date "2026-03-31" --category category.md
```

## Example With Final Summary Only

```bash
python money_helper.py -i "*[vV]" --start_date "2026-03-01" --end_date "2026-03-31" --category category.md --print_mode 0
```

Meaning:

- `-i "*[vV]"`: choose files whose names end with lowercase `v` or uppercase `V`.
- `*`: means "any characters before this point".
- `[vV]`: means "either `v` or `V`".
- Together, `*[vV]` commonly matches files ending in `.csv` or `.CSV`.
- The quotes keep the pattern together as one input value.
- `--start_date "2026-03-01"`: start the report at March 1, 2026.
- `--end_date "2026-03-31"`: end the report at March 31, 2026.
- `--category category.md`: use the category rules from `category.md`.
- `--print_mode 0`: print only the final summary, not the detailed per-file sections.

`--print_mode 1` is the default. It prints the detailed workflow, per-file sections, selected rows, and final summary.

## Specific Example Output

This example summarizes February 2026 transactions from all matching CSV files and prints only the final summary.

Command:

```bash
python money_helper.py -i "*[vV]" --start_date "2026-02-01" --end_date "2026-02-28" --category category.md --print_mode 0
```

Output:

````markdown
# Final Summary
```text
Description: This final summary is the sum of all per-file category totals after the same date, amount, and category-matching rules were applied to each file.

Processed data files:
1. /home/XXX/work/ai_agents/money_helper/Chase0825_Activity20240430_20260430_20260430_XXX.CSV
2. /home/XXX/work/ai_agents/money_helper/Chase2190_Activity20240429_20260429_XXX.csv
3. /home/XXX/work/ai_agents/money_helper/Discover-AllAvailable-20260502_XXX.csv
4. /home/XXX/work/ai_agents/money_helper/citi_XXX.CSV
5. /home/XXX/work/ai_agents/money_helper/stmt_boa_XXX.csv
6. /home/XXX/work/ai_agents/money_helper/stmt_boa_XXX.csv

Matched Rows: 86
Total Balance ($) = Input - Output + Others = 13,248.14 - 12,952.60 + 702.96 = 998.50

Input by Category ($): 13,248.14
  University of Ge & XXX (1): 3,450.68 (26.05%)
  University of Ge & XXX (2): 6,797.46 (51.31%)
  IRS TREAS (0): 0.00 (0.00%)
  Webull (1): 3,000.00 (22.64%)

Output by Category ($): 12,952.60
  HOME, NSM, MORTGAGE (1): 3,098.94 (23.93%)
  Groceries, Food & Drink, COSTCO WHSE, Shopping (31): 2,901.20 (22.40%)
  Education, FAYETTE COUNTY, Music Garden, KEDRON, AFTERCARE (7): 645.72 (4.99%)
  Bills & Utilities, COMCAST, MINT MOBILE, EMC (4): 1,802.48 (13.92%)
  Gas, Automotive, COSTCO GAS (6): 196.21 (1.51%)
  Health & Wellness (0): 0.00 (0.00%)
  Travel (5): 308.05 (2.38%)
  Webull (1): 4,000.00 (30.88%)

Others ($): 702.96
  Input Others (11): 13,036.75
  Output Others (16): 12,333.79
  Input - Output: 13,036.75 - 12,333.79 = 702.96

Checks:
  Row count check: 1 + 2 + 0 + 1 + 1 + 31 + 7 + 4 + 6 + 0 + 5 + 1 + 11 + 16 = 86 --> PASS
  Balance check: 3,450.68 + 6,797.46 + 0.00 + 3,000.00 + -3,098.94 + -2,901.20 + -645.72 + -1,802.48 + -196.21 + 0.00 + -308.05 + -4,000.00 + 13,036.75 + -12,333.79 = 998.50 --> PASS
```
````

How to read this output:

- `Processed data files` shows the exact files included in the report.
- `Matched Rows: 86` means 86 transaction rows passed the date, amount, and zero-amount filters.
- `Input by Category` summarizes positive transactions matched by the `# Input` rules in `category.md`.
- `Output by Category` summarizes negative transactions matched by the `# Output` rules in `category.md`.
- `Others` shows transactions that did not match a named category rule.
- `Total Balance ($)` is calculated as named input minus named output plus net Others.
- `PASS` in the checks means the category row counts and dollar totals add back to the final summary.

## Printing Selected Rows

The `--print` option controls which transaction rows are printed in the detailed report.

Default:

```bash
--print All
```

Useful values:

- `--print All`: print all matched rows.
- `--print Others`: print only rows assigned to `Others`.
- `--print COSTCO`: print rows whose text contains `COSTCO`.
- `--print "COMCAST, EMC"`: print rows whose text contains `COMCAST` or `EMC`.

This option affects row printing in detailed mode. It does not change the totals.

## Recommended Workflow

1. Put all statement files in the same folder as `money_helper.py`.
2. Open each file and confirm it has a date column and an amount column.
3. Edit `category.md` with your own income and spending categories.
4. Run the script with `--print Others` first.
5. Look at the `Others` rows and add new category rules when needed.
6. Run again until `Others` contains only transactions you do not want to categorize.
7. Use `--print_mode 0` when you only need the final summary.

## Using Coding AI Agents

You can ask a coding AI agent to help run this script, improve `category.md`, explain the output, or debug errors.

Do not ask an AI agent to log in to your bank, credit-card account, brokerage account, payroll site, or other financial website. Do not ask it to download statement files for you. These files may contain personal information, account details, transaction history, addresses, or other sensitive data.

Download the statement files yourself from the official bank or credit-card website. After the files are already saved on your own computer, you can ask a coding AI agent to work with the local files.

Example requests you can give to a coding AI agent:

```text
Please run money_helper.py for March 2026 using all CSV files in this folder, category.md, and print_mode 0.
```

```text
Please run this command and explain the final summary:
python money_helper.py -i "*.csv" --start_date "2026-03-01" --end_date "2026-03-31" --category category.md --print_mode 0
```

```text
Please look at the rows printed under Others and suggest updates to category.md.
Do not change the statement files.
```

```text
Please help me revise category.md so grocery, mortgage, utilities, gas, travel, education, and health expenses are grouped clearly.
```

```text
The script says a row has an invalid date. Please inspect the file format and tell me what column or value caused the issue.
```

Good boundaries when using an AI agent:

- You download financial statement files yourself.
- You decide which local files the agent may inspect.
- You ask the agent to run the script, explain the report, or edit `category.md`.
- You avoid sharing passwords, bank login sessions, account recovery codes, or full account numbers.
- You review suggested category changes before relying on the final totals.

## Help Command

To see the script's built-in help:

```bash
python money_helper.py --help
```
