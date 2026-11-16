"""Report-level overrides. A selected ancestor owns its descendants' money."""
from dataclasses import replace
from decimal import Decimal, localcontext
from types import MappingProxyType
from progress_studio.domain.payment_plan_report import series, ZERO


def adjusted_selection(original, edits):
    if not edits:
        return original
    with localcontext() as ctx:
        ctx.prec = 80
        scopes = dict(original.row_values)
        for key, changes in edits.items():
            scope = scopes[key]
            points = []
            for j,v in enumerate(scope.values):
                if j not in changes:
                    points.append((v.fraction,v.money,v.known_money,v.known_basis,v.known_leaves))
                    continue
                f = changes[j]
                m = None if f is None or scope.basis is None or v.money is None else f*scope.basis
                points.append((f,m,m,scope.basis if m is not None else ZERO,
                               len(scope.activities) if m is not None else 0))
            scopes[key] = replace(scope,values=tuple(series(points)))
        if all(key in original.included_in for key in edits):
            return replace(original,row_values=MappingProxyType(scopes))
        owners = [s for k,s in scopes.items() if k not in original.included_in]
        points, retained, visible = [], [], []
        for j in range(len(original.selected.values)):
            vals = [s.values[j] for s in owners]
            known = [v.known_money for v in vals if v.known_money is not None]
            subtotal = sum(known,ZERO) if known else None
            money = subtotal if all(v.money is not None for v in vals) else None
            basis = original.selected.basis
            fraction = money/basis if basis and money is not None else None
            points.append((fraction,money,subtotal,sum(v.known_basis for v in vals),sum(v.known_leaves for v in vals)))
            kept = sum((v.money for v in vals if v.transition=='suppressed'),ZERO)
            retained.append(kept)
            visible.append(None if subtotal is None else subtotal-kept)
        return replace(original,row_values=MappingProxyType(scopes),
                       selected=replace(original.selected,values=tuple(series(points))),
                       retained_money=tuple(retained),visible_known_money=tuple(visible))
