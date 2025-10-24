from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Literal
################################################################################

# Neither of these is vital to the current functionality, but including them
# now may help with future-proofing.
SplitType = Literal["100.00%", "0.00%", "SB"]
BusinessUnit = Literal[
    "CMHEATING ELEC RESI INST",
    "CMHEATING ELEC RESI SALE",
    "CMHEATING ELEC RESI SERV",
    "CMHEATING FPLC RESI INST",
    "CMHEATING FPLC RESI MAIN",
    "CMHEATING FPLC RESI SALE",
    "CMHEATING FPLC RESI SERV",
    "CMHEATING HVAC RESI INST",
    "CMHEATING HVAC RESI MAIN",
    "CMHEATING HVAC RESI SALE",
    "CMHEATING HVAC RESI SERV",
    "CMHEATING PLUM RESI INST",
    "CMHEATING PLUM RESI SALE",
    "CMHEATING PLUM RESI SERV"
]

################################################################################
@dataclass(frozen=True)
class Invoice:
    """Dataclass representing a payroll invoice record.

    Most of the data here isn't important, but in the interest of future-proofing
    the application, I've included everything.
    """

    technician: str
    invoice_id: int
    invoice: int
    invoiced_on: datetime
    customer: str
    total: float
    split: SplitType
    subtotal: float
    cost: float
    bonus: float
    pay_adj: float
    nc_total: float
    net_serv_vol: str
    gp: float
    business_unit: BusinessUnit

    @property
    def net_service_volume_flag(self) -> bool:
        return self.net_serv_vol.endswith("*")

################################################################################
@dataclass(frozen=True)
class DirectPayrollAdjustment:
    """Dataclass representing a direct payroll adjustment record."""

    technician: str
    invoice_id: int
    invoice: int
    posted_on: datetime
    memo: str
    amount: float

################################################################################
