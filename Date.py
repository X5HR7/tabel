from data import week_days
from calendar import weekday, monthrange
class Date():
    def __init__(self, date: str) -> None:
        self._date = date
        self.day = self.get_day()
        self.month = self.get_month()
        self.year = self.get_year()

    def _get_days_amount(self) -> int:
        return monthrange(year=int(self.year), month=int(self.month))[-1]

    def get_day_name(self) -> str:
        #print(self.day)
        day = week_days[weekday(year=int(self.year), month=int(self.month), day=int(self.day))]
        return day

    def get_day(self, type='str'):
        day = self._date.split('.')[0]
        if type == 'int':
            return int(day)
        else:
            return day
    
    def get_month(self, type='str'):
        month = self._date.split('.')[1]
        if type == 'int': 
            return int(month)
        else: 
            return month
    
    def get_year(self, type='str'):
        year = self._date.split('.')[2]
        if type == 'int': 
            return year
        else: 
            return int(year)
        
    def increase_day(self, value: int) -> str:
        if self.day[0] == '0':
            new_day = int(self.day[1:])+value
            return f'{new_day+value}.{self.month}.{self.year}'
                
        else:
            new_day = int(self.day)+value
            if new_day > self._get_days_amount():
                return f'{new_day-self._get_days_amount()}.{int(self.month)+1}.{self.year}'
            else:
                return f'{new_day}.{self.month}.{self.year}'