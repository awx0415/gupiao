import datetime as dt

now = dt.datetime.now()
before_10year = now.year - 10
before_10year_day = f'{before_10year}-{now.month}-{now.day}'
print(before_10year_day)
