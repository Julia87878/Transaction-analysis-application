import datetime
import json
import logging
import os
from functools import wraps
from typing import Optional

import pandas as pd

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(filename)s - %(levelname)s: - %(message)s",
    filename="../logs/reports.log",
    filemode="w",
)

logger = logging.getLogger("reports")


def log(filename="spending"):
    """Декоратор для функций-отчетов, который записывает в файл результат,
    который возвращает функция, формирующая отчет."""

    def my_decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            result = func(*args, **kwargs)
            with open(filename, "w", encoding="utf-8") as file:
                file.write(result)
            return result

        return wrapper

    return my_decorator


project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
path_to_excel_file = os.path.join(project_root, "data", "operations.xlsx")


@log()
def spending_by_category(transactions_2: pd.DataFrame, category: str, date_time_str: Optional[str] = None) -> str:
    """Функция возвращает траты по заданной категории за последние три месяца (от переданной даты)."""
    if date_time_str is None:
        date_time_str = datetime.datetime.now()
    else:
        date_time_str = datetime.datetime.strptime(date_time_str, "%Y-%m-%d %H:%M:%S")
    start_date = date_time_str - pd.DateOffset(months=3)
    transactions_by_category = transactions_2.loc[
        (pd.to_datetime(transactions_2["Дата операции"], dayfirst=True) <= date_time_str)
        & (pd.to_datetime(transactions_2["Дата операции"], dayfirst=True) >= start_date)
        & (transactions_2["Категория"] == category)
    ]
    transactions_by_category_dict = transactions_by_category.to_dict(orient="records")
    if not transactions_by_category_dict:
        logger.warning("Трат по заданной категории за последние три месяца нет.")
        return "Трат по заданной категории за последние три месяца нет."
    else:
        logger.info("Возвращаем траты по заданной категории за последние три месяца.")
        json_str = json.dumps(transactions_by_category_dict, ensure_ascii=False)
        return json_str
