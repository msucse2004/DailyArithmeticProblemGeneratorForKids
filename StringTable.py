

class StringTable:
    def __init__(self):
        self.tbl: dict = {"windowtitle": {"EN": "Daily Arithmetic", "KR": "매일연산"},\
                          "option_negative": {"EN": "Negative Number", "KR": "음수 연산"}, \
                          "option_negative_tip": {"EN": "If you check, it will make the answer negative when creating problems.", "KR": "체크하면 문제를 만들때 정답이 음수도 나오게 만들어요"}, \
                          "option_decimal": {"EN": "Decimal Number", "KR": "소수점 연산"}, \
                          "option_decimal_tip": {"EN": "Checking this box will generate problems that include decimals. If you want to generate integers only, do not check this box.", "KR": "체크하면 소수점까지 포함하는 문제를 만들어요. 정수만 만드려면 체크하지 마세요"}, \
                          "option_fraction": {"EN": "Fraction", "KR": "분수"}, \
                          "option_fraction_tip": {
                              "EN": "Checking this box will generate fraction answers. If you want to generate decimal answers, do not check this box.",
                              "KR": "체크하면 정답이 분수로 나와요. 소수점 표기 정답을 만드려면 체크하지 마세요"}, \
                          "option_bold": {"EN": "Bold", "KR": "볼드체"}, \
                          }

    def findString(self, key, language) -> str:
        return self.tbl.get(key).get(language)