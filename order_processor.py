from data_loader import DataLoader
from info_extractor import InfoExtractor

class OrderProcessor:
    def __init__(self, data_loader:DataLoader):
        self.data_loader = data_loader

    def process_order(self, order_number):
        """
        Обрабатывает заказ и извлекает информацию.

        :param order_number: Номер заказа для поиска
        :return: Извлеченная информация о заказе или сообщение об ошибке.
        """
        df = self.data_loader.load_data()
        filtered_rows = df[df['№ Заказа'].astype(str) == str(order_number)]

        if filtered_rows.empty:
            return f"Заказ № {order_number} не найден!"
        first_row = filtered_rows.iloc[0]
        info_extractor = InfoExtractor(first_row)
        extracted_info = info_extractor.extract()
        return extracted_info
