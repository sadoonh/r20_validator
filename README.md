# r20_validator

Requires Python 3.10+, [uv](https://docs.astral.sh/uv/getting-started/installation/), and Microsoft Excel for Windows. In Excel Trust Center, enable **Trust access to the VBA project object model**.

```bash
git clone https://github.com/sadoonh/r20_validator.git
cd r20_validator
uv sync
uv run python duplicate_excel_columns.py
```
