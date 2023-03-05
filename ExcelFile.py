import pandas as pd
class ExcelFile():
    def __init__(self, file_path: str) -> None:
        self._path = file_path

    def convert_xls_to_xlsx(self) -> None:
        pd.read_excel(self._path).to_excel(self._path+'x', index=False, header=False)

    def get_name(self, extension: bool) -> str:
        if extension:
            return self._path.split('\\')[-1]
        else:
            return '.'.join(self._path.split('\\')[-1].split('.')[0:-1])