"""Source-neutral XML field identity; raw values are validated only if selected."""
from dataclasses import dataclass


@dataclass(frozen=True)
class AmountField:
    identity: str
    name: str
    source: str
    data_type: str
    native_name: str = ""

    @property
    def label(self) -> str:
        detail = f" / {self.native_name}" if self.native_name and self.native_name != self.name else ""
        return f"{self.name}{detail} [{self.identity}; {self.data_type}]"
