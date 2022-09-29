import xlwings
import os

ws0 = xlwings.Book(os.getcwd() + "/KSAT Statistics.xlsm").sheets("Input")
ws1 = xlwings.Book(os.getcwd() + "/2306_도수분포.xlsx").sheets("국수")
ws2 = xlwings.Book(os.getcwd() + "/2306_도수분포.xlsx").sheets("사과탐")

def get_list(ws, rng):
    score_temp = ws.range(rng).value
    score = []
    for item in score_temp:
        if type(item) is float:
            score.append(int(item))

    return score

origin = {
    "국어": ["A8:A154", "D8:D154", ws1],
    "수학": ["F8:F154", "I8:I154", ws1],
    "생윤": ["A5:A53", "D5:D53", ws2],
    "윤사": ["F5:F53", "I5:I53", ws2],
    "한지": ["K5:K53", "N5:N53", ws2],
    "세지": ["P5:P53", "S5:S53", ws2],
    "동사": ["A58:A106", "D58:D106", ws2],
    "세사": ["F58:F106", "I58:I106", ws2],
    "경제": ["K58:K106", "N58:N106", ws2],
    "정법": ["P58:P106", "S58:S106", ws2],
    "사문": ["A111:A158", "D111:D158", ws2],
    "물1": ["A163:A212", "D163:D212", ws2],
    "화1": ["F163:F212", "I163:I212", ws2],
    "생1": ["K163:K212", "N163:N212", ws2],
    "지1": ["P163:P212", "S163:S212", ws2],
    "물2": ["A216:A264", "D216:D264", ws2],
    "화2": ["F216:F264", "I216:I264", ws2],
    "생2": ["K216:K264", "N216:N264", ws2],
    "지2": ["P216:P264", "S216:S264", ws2],
}

destination = {
    "국어": ["O4:O150", "R4:R150"],
    "수학": ["W4:W150", "Z4:Z150"],
    "생윤": ["AH04:AH52", "AK04:AK52"],
    "윤사": ["AS04:AS52", "AV04:AV52"],
    "한지": ["BD04:BD52", "BG04:BG52"],
    "세지": ["AH59:AH107", "AK9:AK107"],
    "동사": ["AS59:AS107", "AV59:AV107"],
    "세사": ["BD59:BD107", "BG59:BG107"],
    "정법": ["AH114:AH162", "AK114:AK162"],
    "경제": ["AS114:AS162", "AV114:AV162"],
    "사문": ["BD114:BD162", "BG114:BG162"],
    "물1": ["BO4:BO52", "BR4:BR52"],
    "화1": ["BZ4:BZ52", "CC4:CC52"],
    "생1": ["CK4:CK52", "CN4:CN52"],
    "지1": ["CV4:CV52", "CY4:CY52"],
    "물2": ["BO59:BO107", "BR59:BR107"],
    "화2": ["BZ59:BZ107", "CC59:CC107"],
    "생2": ["CK59:CK107", "CN59:CN107"],
    "지2": ["CV59:CV107", "CY59:CY107"],
}

for key, value in origin.items():
    loc = destination[key]
    ws0.range(loc[0]).options(transpose=True).clear_contents()
    ws0.range(loc[1]).options(transpose=True).clear_contents()
    ws0.range(loc[0]).options(transpose=True).value = get_list(value[2], value[0])
    ws0.range(loc[1]).options(transpose=True).value = get_list(value[2], value[1])