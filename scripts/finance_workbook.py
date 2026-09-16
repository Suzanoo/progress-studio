"""Create/refresh FF-2 Finance Input + Cash Flow, always to a new output path."""
import argparse
from datetime import date
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from progress_studio.services.financial_forecast_workbook_service import FinancialForecastWorkbookService


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('source',type=Path);p.add_argument('output',type=Path)
    p.add_argument('--opening-date',type=date.fromisoformat)
    p.add_argument('--actuals-through',type=date.fromisoformat)
    p.add_argument('--currency');p.add_argument('--capacity',type=int)
    a=p.parse_args()
    output=FinancialForecastWorkbookService().generate(a.source,a.output,opening_date=a.opening_date,
        actuals_through=a.actuals_through,currency=a.currency,capacity_rows=a.capacity)
    print(output)

if __name__=='__main__':main()
