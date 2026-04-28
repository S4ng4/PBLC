from __future__ import annotations

from models import InvoiceSummary


def evaluate_margin_edge_sum_p(summary: InvoiceSummary) -> float:
    """
    Mirror MarginEdge for LSC / Wholesale:
    L=J6/Ktot, M=(J8+J9)*(F*G)/J5, N=M/G, O=F+L+N, P=O*G
    """
    if not summary.items:
        return 0.0
    total_units = sum(i.unit_qty for i in summary.items)
    if total_units <= 0:
        return 0.0
    j5 = summary.total_gross
    j6 = summary.licensee_discount
    j8 = summary.sales_tax_6
    j9 = summary.local_tax_2
    total_p = 0.0
    for it in summary.items:
        g = float(it.unit_qty)
        f = it.base_price
        l_ = j6 / total_units
        m = (j8 + j9) * (f * g) / j5 if j5 else 0.0
        n = m / g if g else 0.0
        o = f + l_ + n
        total_p += o * g
    return total_p
