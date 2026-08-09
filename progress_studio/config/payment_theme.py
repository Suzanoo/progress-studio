from __future__ import annotations

# MS-PAY6 renders P01 only. Keep the palette centralized so MS-PAY7 can
# add P02..Pn without scattering colors through renderer code.
PAYMENT_LINE_COLORS = {
    "P01": "C00000",  # strong red; high contrast against the blue Plan bar
}

PAYMENT_LINE_STYLE = "medium"
