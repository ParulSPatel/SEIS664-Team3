class Feature():
    def __init__(self, feature_num, days, order, value):
        self.feature_num = feature_num
        self.days = days
        self.order = order
        self.value = self.set_dollars(value)

    def set_dollars(self, dollar_amount):
        return int(dollar_amount.split("/")[0].replace("$", ""))

    def run(self, earnings, start_day, daily_earn):
        for day in range(self.days):
            earnings[start_day] = daily_earn
            start_day += 1
        daily_earn += self.value
        return earnings, start_day, daily_earn
