import datetime
import json
import logging
import os
from typing import Any

import pandas as pd
import requests
from dotenv import load_dotenv

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(filename)s - %(levelname)s: - %(message)s",
    filename="../logs/utils.log",
    filemode="w",
)

logger = logging.getLogger("utils")


project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
path_to_excel_file = os.path.join(project_root, "data", "operations.xlsx")
path_to_user_settings = os.path.join(project_root, "user_settings.json")


def get_greeting(date_time_str: str) -> str:
    """Функция, которая определяет тип приветствия пользователя
    в зависимости от текущего времени."""
    try:
        logger.info("Приветствуем пользователя.")
        date_obj = datetime.datetime.strptime(date_time_str, "%Y-%m-%d %H:%M:%S")
        if 4 <= date_obj.hour < 12:
            return "Доброе утро"
        elif 12 <= date_obj.hour < 17:
            return "Добрый день"
        elif 17 <= date_obj.hour <= 23:
            return "Добрый вечер"
        elif 0 <= date_obj.hour < 4:
            return "Доброй ночи"
    except ValueError:
        logger.error("Некорректная дата")
        return "Некорректная дата"


def get_transactions_excel(path_to_file: str) -> list:
    """Функция для считывания финансовых операций из Excel, которая возращает список транзакций"""
    try:
        logger.info("Считываем финансовые операции из Excel и возращаем список транзакций.")
        transactions_excel_frame = pd.read_excel(path_to_file)
        transactions_excel_list_dict = transactions_excel_frame.to_dict(orient="records")
        return transactions_excel_list_dict
    except FileNotFoundError:
        logger.error("Ошибка: фaйл не найден.")
        print("Ошибка: фaйл не найден.")
        return []


def get_transactions_excel_dataframe(path_to_file: str) -> Any:
    """Функция для чтения данных из Excel, которая возращает pd.DataFrame"""
    try:
        logger.info("Читаем данные из Excel и получаем DataFrame.")
        transactions_excel_df = pd.read_excel(path_to_file)
        return transactions_excel_df
    except FileNotFoundError:
        logger.error("Ошибка: фaйл не найден.")
        print("Ошибка: фaйл не найден.")
        return ""


def filter_transactions_by_date(date_time_str: str, transactions_excel_list_dict: list) -> Any:
    """Функция, которая фильтрует транзакции по дате и возвращает список транзакций с начала месяца, на который
    выпадает входящая дата, по входящую дату."""
    try:
        logger.info("Фильтруем транзакции по дате и возвращаем список транзакций с начала месяца по входящую дату.")
        transactions_for_certain_period = []
        date_obj = datetime.datetime.strptime(date_time_str, "%Y-%m-%d %H:%M:%S")
        the_first_month_day = date_obj.replace(day=1)
        for transaction in transactions_excel_list_dict:
            date_2_string = transaction.get("Дата операции")
            date_obj_2 = datetime.datetime.strptime(date_2_string, "%d.%m.%Y %H:%M:%S")
            if the_first_month_day <= date_obj_2 <= date_obj:
                transactions_for_certain_period.append(transaction)
        if not transactions_for_certain_period:
            logger.warning("Нет транзакций в данном периоде.")
            print("Нет транзакций в данном периоде.")
            return []
        return transactions_for_certain_period
    except ValueError:
        logger.error("Некорректная дата")
        return "Некорректная дата"


def get_expenses_by_cards(transactions_for_certain_period: list) -> list:
    """Функция,которая возвращает расходы и кэшбэк по каждой карте"""
    df = pd.DataFrame(transactions_for_certain_period)
    cards_dict = df.loc[df["Сумма платежа"] < 0].groupby(by="Номер карты").agg("Сумма платежа").sum().to_dict()
    logger.info("Возвращаем расходы и кэшбэк по каждой карте.")
    expenses_by_cards = []
    for card_number, expenses in cards_dict.items():
        expenses_by_cards.append(
            {"last_digits": card_number, "total spent": abs(expenses), "cashback": abs(round(expenses / 100, 2))}
        )
    return expenses_by_cards


def get_last_digits(expenses_by_cards: list) -> list:
    """Функция, которая возвращает 4 последние цифры карты в списке расходов и кэшбэка."""
    new_list = []
    logger.info("Возвращаем 4 последние цифры карты в списке расходов и кэшбэка.")
    for exp in expenses_by_cards:
        last_digits_value = exp.get("last_digits", None)
        if not isinstance(last_digits_value, str):
            continue
        elif isinstance(last_digits_value, str):
            last_digits_value_new = last_digits_value.replace("*", "")
            new_list.append(
                {
                    "last_digits": last_digits_value_new,
                    "total spent": exp.get("total spent", 0),
                    "cashback": exp.get("cashback", 0),
                }
            )
    return new_list


def get_top_five_transactions(transactions_for_certain_period: list) -> list:
    """Функция, которая предоставляет топ-5 транзакций по сумме платежа."""
    list_of_transactions = []
    logger.info("Предоставляем топ-5 транзакций по сумме платежа.")
    df = pd.DataFrame(transactions_for_certain_period)
    df["Сумма платежа"] = df["Сумма платежа"].abs()
    sorted_df = df.sort_values(by="Сумма платежа", ascending=False)
    top_five_df = sorted_df.head()
    top_five_dict = top_five_df.to_dict(orient="records")
    for transaction in top_five_dict:
        info_by_transaction = {
            "date": transaction.get("Дата платежа", ""),
            "amount": transaction.get("Сумма платежа", 0),
            "category": transaction.get("Категория", ""),
            "description": transaction.get("Описание", ""),
        }
        list_of_transactions.append(info_by_transaction)
    return list_of_transactions


def get_currency_rates(path_to_file: str) -> Any:
    """Функция, которая возвращает курс валют"""
    load_dotenv()
    apikey_1 = os.getenv("apikey_rates")
    with open(path_to_file, "r", encoding="utf-8") as file:
        logger.info("Успешное чтение файла с данными.")
        settings = json.load(file)
        user_currencies_list = settings.get("user_currencies", "")
        currency_1 = user_currencies_list[0]
        currency_2 = user_currencies_list[1]
    url_1 = f"https://api.apilayer.com/exchangerates_data/convert?to=RUB&from={currency_1}&amount=1"
    url_2 = f"https://api.apilayer.com/exchangerates_data/convert?to=RUB&from={currency_2}&amount=1"
    headers = {"apikey": f"{apikey_1}"}
    response_1 = requests.get(url_1, headers=headers)
    response_2 = requests.get(url_2, headers=headers)
    logger.info("Возвращаем курс валют.")
    if response_1.status_code == 200 and response_2.status_code == 200:
        result_1 = response_1.json()
        result_2 = response_2.json()
        list_responses_json = [result_1, result_2]
        list_currency_rates = []
        for r in list_responses_json:
            currency_rate = {"currency": r.get("query", "").get("from", ""), "rate": round(r.get("result", 0), 2)}
            list_currency_rates.append(currency_rate)
        return list_currency_rates
    else:
        logger.error("Ошибка.")
        return "Ошибка."


def get_stock_prices() -> Any:
    """Функция, которая получает стоимость акций из S&P500"""
    list_stock_prices = []
    load_dotenv()
    apikey_2 = os.getenv("apikey_stocks")
    url = f"https://financialmodelingprep.com/api/v3/quote/AAPL,AMZN,GOOGL,MSFT,TSLA?apikey={apikey_2}"
    response_2 = requests.get(url)
    logger.info("Получаем стоимость акций из S&P500.")
    if response_2.status_code == 200:
        results = response_2.json()
        for i in results:
            list_stock_prices.append({"stock": i.get("symbol"), "price": round(i.get("price"), 2)})
        return list_stock_prices
    else:
        logger.error("Запрос не был успешным.")
        return "Запрос не был успешным."
