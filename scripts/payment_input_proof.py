"""Build the reproducible PI-1 Excel acceptance case from repository fixtures.

Run from the repository root: python scripts/payment_input_proof.py --output tmp/pi1-proof
This is a developer proof, not a new product export path.
"""
from pathlib import Path
import argparse
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from tests.fixtures.payment_requirements import requirement_source, edit_requirements
from progress_studio.services.payment_service import PaymentService
from progress_studio.infrastructure.excel.payment_input_workbook import PaymentInputWorkbook


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',type=Path,default=Path('tmp/pi1-proof'))
    args=parser.parse_args();args.output.mkdir(parents=True,exist_ok=True)
    source=requirement_source(args.output/'PI1_SOURCE.xlsx')
    embedded=args.output/'PI1_EMBEDDED.xlsx'
    standalone=args.output/'PI1_STANDALONE.xlsx'
    PaymentService().prepare_embedded_payment_input(source,embedded,3)
    PaymentInputWorkbook().create(source,standalone,3)
    for path in (embedded,standalone):
        edit_requirements(path)
        wb=load_workbook(path);ws=wb['Payment Input']
        ws.print_area=f'A1:H{ws.max_row}'
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_setup.orientation='landscape'
        ws.page_setup.paperSize=ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=1
        wb.active=wb.sheetnames.index('Payment Input')
        wb.save(path);wb.close()
    (args.output/'EXPECTED.txt').write_text(
        'PI-1 generated proof. Use F9 in Excel before reading formula results.\n'
        'Source Amounts: A100=100, A200=300, M1=0, A300=200, A400=missing, A500=0.\n'
        'WBS 1: P01=40%, P02=15%, P03=15%; Amount=400. A100 explicit zero lowers the state honestly.\n'
        'Nested WBS 1.2: 20% throughout; repeated branch WBS 2: 50% throughout.\n'
        'WBS 3: Check Amount; WBS 4: No weight; empty WBS 5: blank. No COV rows.\n'
        'Edit A200 P02 to 50%: WBS 1 P02/P03 become 37.5%.\n'
        'Clear A100 P01: WBS 1 P01 becomes 15% (before its explicit zero = derived zero).\n'
        'Clear every A200 input: WBS 1 becomes Incomplete. Restore at least one entry to resolve it.\n'
        'On WBS 2 enter A300 P01=0.0000001%, P02=0%, leave P03 blank: P01 displays rounded 0.0%; '
        'P02/P03 look blank but are numeric zero. Activity P02 explicitly shows 0%.\n'
        'Amount/WBS/helpers locked; only ACT requirements editable (and standalone Payment Dates). '
        'Collapse/expand, freeze F8, F9, save/close/reopen. Amber highlights actual decreasing inputs.\n'
        'Embedded: change main A100 Amount to 125.25, then Prepare again. '
        'Amount refreshes; entered requirements remain. F9 alone does not refresh Amount.\n'
        'These are synthetic fixture data, not real project monetary evidence.\n',encoding='utf-8')
    print(args.output.resolve())


if __name__=='__main__':main()
