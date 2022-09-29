from collections import defaultdict
import xlwings
import numpy
import os
import constraint
import math

ws = xlwings.Book(os.getcwd() + "/KSAT Statistics.xlsm").sheets("Input")
scores = {
    "생윤": "AH04:AH52",
    "윤사": "AS04:AS52",
    "한지": "BD04:BD52",
    "세지": "AH59:AH107",
    "동사": "AS59:AS107",
    "세사": "BD59:BD107",
    "정법": "AH114:AH162",
    "경제": "AS114:AS162",
    "사문": "BD114:BD162",
    "물1": "BO4:BO52",
    "화1": "BZ4:BZ52",
    "생1": "CK4:CK52",
    "지1": "CV4:CV52",
    "물2": "BO59:BO107",
    "화2": "BZ59:BZ107",
    "생2": "CK59:CK107",
    "지2": "CV59:CV107"
}

subject = input("탐구 선택과목을 입력하세요. : ")

tscore_list = ws.range(scores[subject]).value
tscore_list = [x for x in tscore_list if x != None]  # 공백 셀 제거
pscore_list = list(range(51))  # 원점수 범위
pscore_list.remove(1)  # 존재하지 않는 원점수 제거
pscore_list.remove(49)
pscore_list.reverse()  # 순서 뒤집기

# 만점 표준점수, 0점 표준점수
max_score = tscore_list[0]
min_score = tscore_list[-1]

# 범위 제한
x_max = math.ceil(50000 / (max_score - min_score - 1)) / 100
x_min = math.trunc(50000 / (max_score - min_score + 1)) / 100
y_max = math.ceil(100 * (2525 - 50 * min_score) / (max_score - min_score)) / 100
y_min = math.trunc(100 * (2475 - 50 * min_score) / (max_score - min_score)) / 100
x_range = str(x_min) + " ~ " + str(x_max)
y_range = str(y_min) + " ~ " + str(y_max)
print("평균 범위: " + y_range)
print("표준편차 범위: " + x_range)

while True:
    duplicate_score_cnt = len(pscore_list) - len(tscore_list)
    msg = "표점 증발을 모두 입력하세요. 표점 증발 수는 " + str(duplicate_score_cnt) + "입니다. : "
    duplicate_assume = list(map(int, input(msg).strip().split()))

    # 부등식
    def inequalities(x, y):
        if (
            (50 - ((max_score - 50.5) / 10 * x)) >= y
            and (50 - ((max_score - 49.5) / 10 * x)) < y
            and ((50.5 - min_score) / 10 * x) >= y
            and ((49.5 - min_score) / 10 * x) < y
        ):
            return True

    # 연립부등식 problem 생성
    problem = constraint.Problem()
    problem.addVariable("x", numpy.arange(x_min, x_max, 0.01))  # 표준편차 = x
    problem.addVariable("y", numpy.arange(y_min, y_max, 0.01))  # 평균 = y
    problem.addConstraint(inequalities, ["x", "y"])
    solutions = problem.getSolutions()
    total_list = []  # 가능한 모든 경우의 수

    for solution in solutions:
        value_list = []

        for key, value in solution.items():
            value = round(value, 2)  # 반올림
            value_list.append(value)

        total_list.append(value_list)

    # 표점 증발이 실제와 일치하는지 확인
    def duplicates_check(source):
        out = defaultdict(set)
        for key, value in source.items():
            out[value].add(key)
        out = dict(out)

        flg = True

        for item in duplicate_assume:
            if len(out[item]) < 2:
                flg = False
                break

        if flg is True:
            return True
        else:
            return False

    final_dict = {}

    for solution in total_list:
        stdev = solution[0]
        mean = solution[1]
        tscore_list_new = []

        for i in range(0, len(pscore_list)):
            tscore = int(round(10 * (pscore_list[i] - mean) / stdev + 50, 0))
            tscore_list_new.append(tscore)

        duplicate_list = []
        tscore_list_new_remove_duplicates = [*set(tscore_list_new)]
        for score in tscore_list_new_remove_duplicates:
            if tscore_list_new.count(score) > 1:
                i = 1
                while i < tscore_list_new.count(score):
                    duplicate_list.append(score)
                    i = i + 1

        tscore_pscore_dict = dict(zip(pscore_list, tscore_list_new))

        if duplicates_check(tscore_pscore_dict) is True and len(duplicate_list) == duplicate_score_cnt:
            final_dict[(solution[1], solution[0])] = tuple(duplicate_list)

    output = defaultdict(set)
    for key, value in final_dict.items():
        output[value].add(key)
    output = dict(output)

    for key, value in output.items():
        print(key, ':', len(value), '가지-', value, '\n')
