from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Tuple


@dataclass
class InvoiceItem:
    line: int
    plcb_code: str
    description: str
    unit_volume: str
    uom: str
    base_price: float
    base_freight: float
    ltmf: float
    total_unit_price: float
    unit_qty: int
    line_amount: float
    case_pack: Optional[int] = None


@dataclass
class InvoiceSummary:
    invoice_type: str
    invoice_number: str
    invoice_date: str
    order_number: str
    customer_name: str
    customer_number: str
    total_base_price: float = 0.0
    delivery_fee: float = 0.0
    ltmf_fee: float = 0.0
    total_taxable: float = 0.0
    sales_tax_6: float = 0.0
    local_tax_2: float = 0.0
    base_freight: float = 0.0
    total_gross: float = 0.0
    licensee_discount: float = 0.0
    taxable_amount: float = 0.0
    total_net_price: float = 0.0
    total_units: int = 0
    items: List[InvoiceItem] = field(default_factory=list)
    pickup_store: str = ""
    order_date: str = ""
    vendor_number: str = ""
    ship_date: str = ""
    payment_method: str = ""
    account_last_four: str = ""
    delivery_method: str = ""
    delivery_date: str = ""


def verify_totals(summary: InvoiceSummary, calculated_total: float) -> Tuple[bool, float]:
    target = summary.total_net_price
    diff = abs(target - calculated_total)
    tolerance = 0.01
    return diff <= tolerance, diff
