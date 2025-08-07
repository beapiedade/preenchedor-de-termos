import pandas as pd
import numpy as np

class Extractor:
    @staticmethod
    def from_path(path):
        data_frame = pd.read_excel(path, dtype=str)
        data_frame.replace(np.nan, '', inplace=True)
        return data_frame.to_records(index=False).tolist()
    
    @staticmethod
    def check_data(data):
        errors = []
        for row in data:
            for value in row:
                if type(value) != str:
                    data.remove(row)
                    errors.append(row)
                    break
        return data, errors