from pathlib import Path

from progress_studio.services.distribution import get_distribution, list_distributions


class DistributionPrompt:
    def choose_method(self) -> str | None:
        options = list_distributions()
        print("\n" + "=" * 72)
        print("Select Plan Distribution")
        print("=" * 72)
        print("1. Auto Distribution  [recommended]")
        for index, spec in enumerate(options, start=2):
            print(f"{index}. {spec.name}")
        print("6. Cancel")
        mapping = {"1": "auto", "2": "flat", "3": "front", "4": "back", "5": "bell"}
        while True:
            selected = input("Select [1-6]: ").strip() or "1"
            if selected in mapping:
                return mapping[selected]
            if selected == "6":
                return None
            print("Please select 1 to 6")

    def review(self, output_file: Path, method: str) -> str:
        method_name = "Auto Distribution" if method == "auto" else get_distribution(method).name
        print("\n" + "=" * 72)
        print("Review Distribution")
        print("=" * 72)
        print(f"METHOD : {method_name}")
        print(f"OUTPUT : {output_file}")
        print("1. Accept this distribution")
        print("2. Try another distribution")
        print("3. Exit and keep current file")
        while True:
            selected = input("Select [1-3]: ").strip()
            if selected == "1": return "accept"
            if selected == "2": return "retry"
            if selected == "3": return "exit"
            print("Please select 1, 2 or 3")
