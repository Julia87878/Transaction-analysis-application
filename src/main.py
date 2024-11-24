import os
from typing import Any

from src.reports import spending_by_category
from src.services import get_transactions_by_search_string
from src.utils import get_transactions_excel, get_transactions_excel_dataframe
from src.views import get_data_json

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
path_to_excel_file = os.path.join(project_root, "data", "operations.xlsx")
transactions = get_transactions_excel(path_to_excel_file)
transactions_2 = get_transactions_excel_dataframe(path_to_excel_file)


def main(date_time_str: Any, search_str: str, category: str) -> Any:
    print(get_data_json(date_time_str))
    print(get_transactions_by_search_string(transactions, search_str))
    print(spending_by_category(transactions_2, category, date_time_str))


if __name__ == "__main__":
    main("2021-07-30 12:10:24", "Перевод", "Супермаркеты")
