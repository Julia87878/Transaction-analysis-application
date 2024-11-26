import json
import logging
import os

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(filename)s - %(levelname)s: - %(message)s",
    filename="../logs/services.log",
    filemode="w",
)

logger = logging.getLogger("services")


project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
path_to_excel_file = os.path.join(project_root, "data", "operations.xlsx")


def get_transactions_by_search_string(transactions: list, search_str: str) -> str:
    """Функция, которая принимает на вход строку для поиска и возвращает JSON-ответ со всеми транзакциями,
    содержащими запрос в описании или категории."""
    new_list = []
    search_string = search_str.lower()
    for transaction in transactions:
        description_value = transaction.get("Описание", "").lower()
        category_value = str(transaction.get("Категория", "")).lower()
        if search_string in description_value or search_string in category_value:
            new_list.append(transaction)
    result_json = json.dumps(new_list, ensure_ascii=False)
    logger.info("Возвращаем JSON-ответ со всеми транзакциями, содержащими запрос в описании или категории.")
    return result_json
