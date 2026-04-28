from __future__ import annotations

from models import InvoiceSummary


def evaluate_margin_edge_sum_p(summary: InvoiceSummary) -> float:
    """
    Mirror MarginEdge formulas for Special Order rows to approximate SUM(P).
    Uses the same structure as Excel: L=J6/Ktot, M=(J9+J10)*(F*J)/J5, N=M/J, O=I+L+N, P=O*J.
    """
    if not summary.items:
        return 0.0
    total_units = sum(i.unit_qty for i in summary.items)
    if total_units <= 0:
        return 0.0
    j5 = summary.total_base_price
    j6 = summary.delivery_fee
    j9 = summary.sales_tax_6
    j10 = summary.local_tax_2
    total_p = 0.0
    for it in summary.items:
        j = float(it.unit_qty)
        f = it.base_price
        iu = it.total_unit_price
        l_ = j6 / total_units
        m = (j9 + j10) * (f * j) / j5 if j5 else 0.0
        n = m / j if j else 0.0
        o = iu + l_ + n
        total_p += o * j
    return total_p
