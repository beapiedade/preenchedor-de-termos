import pandas as pd

class Extractor:
    @staticmethod
    def from_path(path):
        data_frame = pd.read_excel(path)
        tuples = data_frame.to_records(index=False)
        data = []
        for row in tuples:
            data.append(row.tolist())
        return data
    
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