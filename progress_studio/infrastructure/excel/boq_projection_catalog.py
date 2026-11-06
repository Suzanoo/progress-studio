"""Lossless Project catalog for projection; existing Mapping filtering is unchanged."""
from pathlib import Path
import re
from openpyxl import load_workbook
from progress_studio.domain.boq_payment_projection import CatalogRow, number
from progress_studio.infrastructure.excel.mapping_reader import _stable_boq_id, _formula_fallback_value


def _supported_link(formula):
    ref = r"(?:'(?:[^']|'')+'|[A-Za-z0-9 _.-]+)!\$?[A-Z]+\$?\d+"
    return bool(isinstance(formula,str) and re.fullmatch(
        r'=IF\((?P<ref>'+ref+r')="","",(?P=ref)\)', formula.strip(), re.IGNORECASE))


def read_catalog(path: Path, sheet_name: str):
    formulas = load_workbook(path,read_only=True,data_only=False)
    values = load_workbook(path,read_only=True,data_only=True)
    try:
        if sheet_name not in values.sheetnames: raise ValueError('BOQ sheet missing')
        raw = list(formulas[sheet_name].values); cached = list(values[sheet_name].values)
        if not raw: raise ValueError('Empty BOQ catalog')
        headers = {str(x or '').strip().lower():i for i,x in enumerate(raw[0])}
        if not {'wbs-2','wbs-3','wbs-4','description','amount'} <= headers.keys():
            raise ValueError('BOQ catalog headers missing')
        def get(row,key):
            i=headers.get(key); return row[i] if i is not None and i<len(row) else None
        paths=[str(get(r,'wbs no.') or '').strip() for r in cached[1:]]
        result=[]
        for index,(f,v) in enumerate(zip(raw[1:],cached[1:]),2):
            if all(x is None for x in f):continue
            def resolved(key):
                value=get(v,key);formula=get(f,key)
                return _formula_fallback_value(value,formula,values) if value is not None or _supported_link(formula) else formula
            text=lambda key:str(resolved(key) or '').strip()
            sheet=text('source sheet') or sheet_name
            row=get(v,'source row')
            if row is None: row=index
            if isinstance(row,bool): raise ValueError('Invalid source row')
            try:
                row_num=int(row)
                if row_num<1 or str(row_num)!=str(row).removesuffix('.0'):raise ValueError()
            except (TypeError,ValueError):raise ValueError('Invalid source row')
            raw_amount=get(f,'amount');amount_value=get(v,'amount')
            status='Valid'
            if isinstance(raw_amount,str) and raw_amount.startswith('='):
                if amount_value is None:
                    # Existing BOQ Studio link format is supported; arbitrary
                    # uncalculated business formulas are never evaluated here.
                    if _supported_link(raw_amount):
                        amount_value=_formula_fallback_value(None,raw_amount,values)
                        status='ResolvedSourceLink'
                    else:status='UnresolvedFormula'
                else:status='CachedFormula'
            else:amount_value=raw_amount
            try:
                amount=number(amount_value)
                if amount<0:status='NegativeAmount'
                elif amount==0:status='ZeroAmount'
            except ValueError:
                amount=None
                if status!='UnresolvedFormula':status='MissingAmount' if amount_value in (None,'') else 'InvalidAmount'
            wbs=text('wbs no.');hierarchy=tuple(text(k) for k in ('wbs-2','wbs-3','wbs-4'))
            unit=resolved('unit');qty=resolved('qty')
            has_children=bool(wbs and any(x.startswith(wbs+'.') for x in paths))
            role='group' if has_children or (amount_value in (None,'') and (raw_amount in (None,'') or _supported_link(raw_amount)) and unit in (None,'') and qty in (None,'') and bool(wbs)) else 'item'
            result.append(CatalogRow(f'{sheet}|{row_num}|{index}',_stable_boq_id(sheet,row_num,*hierarchy,text('description')),sheet,row_num,index,wbs,hierarchy,text('description'),unit,qty,raw_amount,amount,status,role))
        return tuple(result)
    finally:
        formulas.close();values.close()
