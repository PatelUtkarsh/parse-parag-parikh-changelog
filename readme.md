# PPFAS Mutual Fund Portfolio Parser

A tool to compare PPFAS mutual fund holdings between different months and see which stocks increased or decreased.

## Installation

```bash
composer install
```

## Usage

### Basic Command
```bash
php index.php parse [months-back] [options]
```

### Options
- `--fund-name` or `-f`: Choose fund type
  - `flexi` - Flexi Cap Fund
  - `liquid` - Liquid Fund
  - `hybrid` - Hybrid Fund
  - `tax` - Tax Saver Fund (default)
- `--month-diff-y` or `-y`: Second month to compare (default: 1)

### Examples

**Compare last 2 months for Tax Saver fund:**
```bash
php ./index.php parse 2 --fund-name=tax

**Compare current vs last month for Flexi Cap:**
```bash
php index.php parse 2 --fund-name=flexi
```

**Compare 3 months ago vs last month:**
```bash
php index.php parse 3 --month-diff-y=1
```

## Output

Shows multiple tables:

1. **Overall Changes** - All stocks with increases/decreases
2. **By Category** - Changes grouped by investment type (Equity, Foreign holding, Diposites etc.)

Green numbers = increases, Red numbers = decreases

The tool downloads files automatically and shows percentage changes between the selected months.