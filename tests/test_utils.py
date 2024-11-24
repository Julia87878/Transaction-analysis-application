import os
from unittest.mock import Mock, patch

import pandas as pd
import pytest

from src.utils import (filter_transactions_by_date, get_currency_rates, get_expenses_by_cards, get_greeting,
                       get_last_digits, get_stock_prices, get_top_five_transactions, get_transactions_excel,
                       get_transactions_excel_dataframe)


@pytest.fixture
def transactions():
    return [
        {
            "Дата операции": "12.01.2018 16:04:18",
            "Дата платежа": "15.01.2018",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -6.96,
            "Валюта операции": "EUR",
            "Сумма платежа": -488.59,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Одежда и обувь",
            "MCC": 5641.0,
            "Описание": "Varaviksne-veikals",
            "Бонусы (включая кэшбэк)": 9,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 488.59,
        },
        {
            "Дата операции": "12.01.2018 15:58:01",
            "Дата платежа": "15.01.2018",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -28.29,
            "Валюта операции": "EUR",
            "Сумма платежа": -1985.96,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Одежда и обувь",
            "MCC": 5641.0,
            "Описание": "Varaviksne-veikals",
            "Бонусы (включая кэшбэк)": 39,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 1985.96,
        },
        {
            "Дата операции": "12.01.2018 11:08:52",
            "Дата платежа": "15.01.2018",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -22.43,
            "Валюта операции": "EUR",
            "Сумма платежа": -1574.59,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Супермаркеты",
            "MCC": 5411.0,
            "Описание": "Rimi Hm Gramzdas",
            "Бонусы (включая кэшбэк)": 31,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 1574.59,
        },
        {
            "Дата операции": "11.01.2018 17:21:40",
            "Дата платежа": "12.01.2018",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -44.46,
            "Валюта операции": "EUR",
            "Сумма платежа": -3089.97,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Одежда и обувь",
            "MCC": 5631.0,
            "Описание": "Veikals Rosme, Mukusalas",
            "Бонусы (включая кэшбэк)": 61,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 3089.97,
        },
        {
            "Дата операции": "11.01.2018 00:00:00",
            "Дата платежа": "13.01.2018",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -94.0,
            "Валюта операции": "RUB",
            "Сумма платежа": -94.0,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Транспорт",
            "MCC": 4121.0,
            "Описание": "Яндекс Такси",
            "Бонусы (включая кэшбэк)": 1,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 94.0,
        },
        {
            "Дата операции": "10.01.2018 21:31:46",
            "Дата платежа": "11.01.2018",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -2.75,
            "Валюта операции": "EUR",
            "Сумма платежа": -191.13,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Топливо",
            "MCC": 5541.0,
            "Описание": "Circle K",
            "Бонусы (включая кэшбэк)": 3,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 191.13,
        },
        {
            "Дата операции": "10.01.2018 19:10:56",
            "Дата платежа": "12.01.2018",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -1041.74,
            "Валюта операции": "RUB",
            "Сумма платежа": -1041.74,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Различные товары",
            "MCC": 5311.0,
            "Описание": "DutyFree",
            "Бонусы (включая кэшбэк)": 20,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 1041.74,
        },
        {
            "Дата операции": "10.01.2018 16:53:21",
            "Дата платежа": "11.01.2018",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -634.0,
            "Валюта операции": "RUB",
            "Сумма платежа": -634.0,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Рестораны",
            "MCC": 5812.0,
            "Описание": "Stedroll",
            "Бонусы (включая кэшбэк)": 12,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 634.0,
        },
    ]


@pytest.fixture
def date_time_str():
    return "2019-07-15 23:37:42"


def test_get_greeting(date_time_str: str) -> None:
    result = get_greeting(date_time_str)
    expected = "Добрый вечер"
    assert result == expected


@pytest.fixture
def date_time_str_2():
    return "2019-07-15 09:37:42"


def test_get_greeting_2(date_time_str_2: str) -> None:
    result = get_greeting(date_time_str_2)
    expected = "Доброе утро"
    assert result == expected


@pytest.fixture
def date_time_str_3():
    return "15-07-2019 09:37:42"


def test_get_greeting_3(date_time_str_3: str) -> None:
    result = get_greeting(date_time_str_3)
    expected = "Некорректная дата"
    assert result == expected


def test_get_greeting_4(date_time_str: str) -> None:
    result = get_greeting("2019-07-15 13:50:45")
    expected = "Добрый день"
    assert result == expected


def test_get_greeting_5(date_time_str: str) -> None:
    result = get_greeting("2019-07-15 00:57:45")
    expected = "Доброй ночи"
    assert result == expected


@pytest.mark.parametrize(
    "date_time_str, expected",
    [
        ("2019-07-15 09:12:27", "Доброе утро"),
        ("2019-07-15 14:10:20", "Добрый день"),
        ("2019-07-15 22:15:24", "Добрый вечер"),
        ("2019-07-15 02:32:07", "Доброй ночи"),
        ("15-07-2019 02:32:07", "Некорректная дата"),
    ],
)
def test_get_greeting_7(date_time_str: str, expected: str) -> None:
    assert get_greeting(date_time_str) == expected


@pytest.fixture
def test_df():
    test_dict = {
        "Дата операции": ["05.01.2018 14:58:38", "04.01.2018 15:00:41"],
        "Дата платежа": ["07.01.2018", "05.01.2018"],
        "Номер карты": ["*7197", "*7197"],
        "Статус": ["OK", "OK"],
        "Сумма операции": [-120.0, -1025.0],
        "Валюта операции": ["RUB", "RUB"],
        "Сумма платежа": [-120.0, -1025.0],
        "Валюта платежа": ["RUB", "RUB"],
        "Кэшбэк": [0, 0],
        "Категория": ["Цветы", "Топливо"],
        "MCC": [5992.0, 5541.0],
        "Описание": ["Magazin  Prestizh", "Pskov AZS 12 K2"],
        "Бонусы (включая кэшбэк)": [2, 20],
        "Округление на инвесткопилку": [0, 0],
        "Сумма операции с округлением": [120.0, 1025.0],
    }

    return pd.DataFrame(test_dict)


@pytest.fixture
def path_name_1():
    path_to_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "operations.xlsx")
    return path_to_file


@patch("src.utils.pd.read_excel")
def test_get_transactions_excel(mock_read: Mock, test_df, path_name_1: str) -> None:
    mock_read.return_value = test_df
    result = get_transactions_excel(path_name_1)
    expected = test_df.to_dict(orient="records")
    assert result == expected


@pytest.fixture
def path_name_2():
    path_to_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "operations_2.xlsx")
    return path_to_file


def test_get_transactions_excel_file_not_found(path_name_2: str) -> None:
    assert get_transactions_excel(path_name_2) == []


def test_get_transactions_excel_dataframe_file_not_found(path_name_2: str) -> None:
    assert get_transactions_excel_dataframe(path_name_2) == ""


@pytest.fixture
def transactions_excel():
    return [
        {
            "Дата операции": "15.07.2019 23:37:42",
            "Дата платежа": "17.07.2019",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -77.36,
            "Валюта операции": "RUB",
            "Сумма платежа": -77.36,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Супермаркеты",
            "MCC": 5411.0,
            "Описание": "SPAR",
            "Бонусы (включая кэшбэк)": 1,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 77.36,
        },
        {
            "Дата операции": "15.07.2019 20:00:31",
            "Дата платежа": "15.07.2019",
            "Номер карты": 0,
            "Статус": "OK",
            "Сумма операции": -50000.0,
            "Валюта операции": "RUB",
            "Сумма платежа": -50000.0,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Переводы",
            "MCC": 0,
            "Описание": 'На р/с ООО "ФОРТУНА"',
            "Бонусы (включая кэшбэк)": 0,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 50000.0,
        },
        {
            "Дата операции": "15.07.2019 16:55:34",
            "Дата платежа": "16.07.2019",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -120.0,
            "Валюта операции": "RUB",
            "Сумма платежа": -120.0,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Фастфуд",
            "MCC": 5814.0,
            "Описание": "IP Yakubovskaya M. V.",
            "Бонусы (включая кэшбэк)": 2,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 120.0,
        },
        {
            "Дата операции": "15.07.2019 16:50:10",
            "Дата платежа": "17.07.2019",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -114.19,
            "Валюта операции": "RUB",
            "Сумма платежа": -114.19,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Супермаркеты",
            "MCC": 5499.0,
            "Описание": "Колхоз",
            "Бонусы (включая кэшбэк)": 2,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 114.19,
        },
        {
            "Дата операции": "15.07.2019 13:43:11",
            "Дата платежа": "17.07.2019",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -200.0,
            "Валюта операции": "RUB",
            "Сумма платежа": -200.0,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Ж/д билеты",
            "MCC": 4111.0,
            "Описание": "Метро Санкт-Петербург",
            "Бонусы (включая кэшбэк)": 4,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 200.0,
        },
    ]


def test_filter_transactions_by_date_1(date_time_str: str, transactions_excel: list) -> None:
    result = filter_transactions_by_date(date_time_str, transactions_excel)
    expected = [
        {
            "Дата операции": "15.07.2019 23:37:42",
            "Дата платежа": "17.07.2019",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -77.36,
            "Валюта операции": "RUB",
            "Сумма платежа": -77.36,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Супермаркеты",
            "MCC": 5411.0,
            "Описание": "SPAR",
            "Бонусы (включая кэшбэк)": 1,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 77.36,
        },
        {
            "Дата операции": "15.07.2019 20:00:31",
            "Дата платежа": "15.07.2019",
            "Номер карты": 0,
            "Статус": "OK",
            "Сумма операции": -50000.0,
            "Валюта операции": "RUB",
            "Сумма платежа": -50000.0,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Переводы",
            "MCC": 0,
            "Описание": 'На р/с ООО "ФОРТУНА"',
            "Бонусы (включая кэшбэк)": 0,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 50000.0,
        },
        {
            "Дата операции": "15.07.2019 16:55:34",
            "Дата платежа": "16.07.2019",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -120.0,
            "Валюта операции": "RUB",
            "Сумма платежа": -120.0,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Фастфуд",
            "MCC": 5814.0,
            "Описание": "IP Yakubovskaya M. V.",
            "Бонусы (включая кэшбэк)": 2,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 120.0,
        },
        {
            "Дата операции": "15.07.2019 16:50:10",
            "Дата платежа": "17.07.2019",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -114.19,
            "Валюта операции": "RUB",
            "Сумма платежа": -114.19,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Супермаркеты",
            "MCC": 5499.0,
            "Описание": "Колхоз",
            "Бонусы (включая кэшбэк)": 2,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 114.19,
        },
        {
            "Дата операции": "15.07.2019 13:43:11",
            "Дата платежа": "17.07.2019",
            "Номер карты": "*7197",
            "Статус": "OK",
            "Сумма операции": -200.0,
            "Валюта операции": "RUB",
            "Сумма платежа": -200.0,
            "Валюта платежа": "RUB",
            "Кэшбэк": 0,
            "Категория": "Ж/д билеты",
            "MCC": 4111.0,
            "Описание": "Метро Санкт-Петербург",
            "Бонусы (включая кэшбэк)": 4,
            "Округление на инвесткопилку": 0,
            "Сумма операции с округлением": 200.0,
        },
    ]

    assert result == expected


def test_filter_transactions_by_date_2(date_time_str_3, transactions_excel) -> None:
    result = filter_transactions_by_date(date_time_str_3, transactions_excel)
    expected = "Некорректная дата"
    assert result == expected


@pytest.fixture
def date_time_str_4():
    return "2020-07-15 09:37:42"


def test_filter_transactions_by_date_3(date_time_str_4, transactions_excel) -> None:
    result = filter_transactions_by_date(date_time_str_4, transactions_excel)
    expected = []
    assert result == expected


def test_get_expenses_by_cards(transactions_excel) -> None:
    result = get_expenses_by_cards(transactions_excel)
    expected = [
        {"last_digits": 0, "total spent": 50000.0, "cashback": 500.0},
        {"last_digits": "*7197", "total spent": 511.55, "cashback": 5.12},
    ]
    assert result == expected


@pytest.fixture
def expenses():
    return [
        {"last_digits": "*7295", "total spent": 50000.0, "cashback": 500.0},
        {"last_digits": "*7197", "total spent": 511.55, "cashback": 5.12},
    ]


def test_get_last_digits(expenses) -> None:
    result = get_last_digits(expenses)
    expected = [
        {"last_digits": "7295", "total spent": 50000.0, "cashback": 500.0},
        {"last_digits": "7197", "total spent": 511.55, "cashback": 5.12},
    ]
    assert result == expected


def test_get_top_five_transactions(transactions) -> None:
    result = get_top_five_transactions(transactions)
    expected = [
        {
            "date": "12.01.2018",
            "amount": 3089.97,
            "category": "Одежда и обувь",
            "description": "Veikals Rosme, Mukusalas",
        },
        {"date": "15.01.2018", "amount": 1985.96, "category": "Одежда и обувь", "description": "Varaviksne-veikals"},
        {"date": "15.01.2018", "amount": 1574.59, "category": "Супермаркеты", "description": "Rimi Hm Gramzdas"},
        {"date": "12.01.2018", "amount": 1041.74, "category": "Различные товары", "description": "DutyFree"},
        {"date": "11.01.2018", "amount": 634.0, "category": "Рестораны", "description": "Stedroll"},
    ]

    assert result == expected


@pytest.fixture
def path_user_settings():
    path_to_settings = os.path.join(os.path.dirname(os.path.dirname(__file__)), "user_settings.json")
    return path_to_settings


@patch("src.utils.requests.get")
def test_get_currency_rates(mock_get, path_user_settings) -> None:
    mock_get.return_value.status_code = 200
    mock_get.return_value.json.return_value = {
        "success": True,
        "query": {"from": "USD", "to": "RUB", "amount": 1},
        "info": {"timestamp": 1732051636, "rate": 100.572674},
        "date": "2024-11-19",
        "result": 100.572674,
    }
    assert get_currency_rates(path_user_settings) == [
        {"currency": "USD", "rate": 100.57},
        {"currency": "USD", "rate": 100.57},
    ]


@patch("src.utils.requests.get")
def test_get_stock_prices(mock_get: Mock) -> None:
    mock_get.return_value.status_code = 200
    mock_get.return_value.json.return_value = [{"symbol": "AAPL", "name": "Apple Inc.", "price": 228.28}]
    assert get_stock_prices() == [{"stock": "AAPL", "price": 228.28}]
