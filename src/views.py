import json
import os

from src.utils import (filter_transactions_by_date, get_currency_rates, get_expenses_by_cards, get_greeting,
                       get_last_digits, get_stock_prices, get_top_five_transactions, get_transactions_excel)

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
path_to_excel_file = os.path.join(project_root, "data", "operations.xlsx")
path_to_user_settings = os.path.join(project_root, "user_settings.json")


def get_data_json(date_time_str: str) -> str:
    """Функция, принимающая на вход строку с датой и временем в формате YYYY-MM-DD HH:MM:SS и
    возвращающую JSON - ответ с данными"""
    result_greeting = get_greeting(date_time_str)
    res_cards = get_last_digits(
        get_expenses_by_cards(filter_transactions_by_date(date_time_str, get_transactions_excel(path_to_excel_file)))
    )
    transactions_5 = get_top_five_transactions(
        filter_transactions_by_date(date_time_str, get_transactions_excel(path_to_excel_file))
    )
    result_currency_rates = get_currency_rates(path_to_user_settings)
    result_stock_prices = get_stock_prices()

    data = {
        "greeting": result_greeting,
        "cards": res_cards,
        "top_transactions": transactions_5,
        "currency_rates": result_currency_rates,
        "stock_prices": result_stock_prices,
    }

    data_json = json.dumps(data, ensure_ascii=False)
    return data_json
