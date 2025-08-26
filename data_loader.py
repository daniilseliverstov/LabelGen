import pandas as pd
from abc import ABC, abstractmethod


class DataLoader(ABC):
    @abstractmethod
    def load_data(self, filename):
        """Загрузка данных из файла"""
        pass

class ExcelDataLoader(DataLoader):
    def __init__(self):
        self.filename = None

    def load_data(self, filename=None):
        """
        Загружает данные из Excel файла.

        :param filename: Путь к файлу для загрузки. Если не указан используется сохраненный путь.
        :return: DataFrame с загруженными данными.
        """
        try:
            file_to_load = filename or self.filename
            if not file_to_load:
                raise ValueError("Не указан файл Раскроя для загрузки")

            # Определение движка в зависимости от расширения файла
            if file_to_load.endswith('.xls'):
                return pd.read_excel(file_to_load, engine='xlrd')
            else:
                return pd.read_excel(file_to_load, engine='openpyxl')

        except FileNotFoundError:
            raise ValueError(f'Файл "{file_to_load}" не найден')
        except Exception as e:
            raise RuntimeError(f'Ошибка при загрузке данных {e}')
