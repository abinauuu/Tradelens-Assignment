# TradeLens

TradeLens is a Python CLI tool that combines trading data from MT5 CSV files and Indian broker Excel files into a single report. It standardizes trade records, converts MT5 timestamps to IST, calculates P&L in INR, and checks trades against configured rules.

## Usage

Install the dependencies:

```bash
pip install -r requirements.txt
```

Run TradeLens:

```bash
python run.py --data sample/data --rules sample/rules.json --out output
```

Run the tests:

```bash
pip install pytest
pytest
```

## Output

The output folder contains:

- `trades.csv` - Normalized trade records.
- `summary.json` - Trading summary and performance statistics.
- `violations.csv` - Detected rule violations.
- `report.md` - A readable report of the results.

The USD-to-INR exchange rate and trading rules are configured in `sample/rules.json`.

Added GROQ api integration for reports but if the key is unavailable or the API call fails, TradeLens generates the report using a built-in template.

## With more time.

- Add more tests for unusual Excel layouts and additional rule types.
- Add extra functionality such as logging etc for better Performance analysis ,Error detection and identification.
