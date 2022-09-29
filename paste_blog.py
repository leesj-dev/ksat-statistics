import xlwings
import os

ws0 = xlwings.Book(os.getcwd() + "/KSAT Statistics.xlsm").sheets("Input")
ws1 = xlwings.Book(os.getcwd() + "/2309_도수분포.xlsx").sheets("Table 1")
ws2 = xlwings.Book(os.getcwd() + "/2309_도수분포.xlsx").sheets("Table 2")
ws3 = xlwings.Book(os.getcwd() + "/2309_도수분포.xlsx").sheets("Table 3")
ws4 = xlwings.Book(os.getcwd() + "/2309_도수분포.xlsx").sheets("Table 4")
ws5 = xlwings.Book(os.getcwd() + "/2309_도수분포.xlsx").sheets("Table 5")
ws6 = xlwings.Book(os.getcwd() + "/2309_도수분포.xlsx").sheets("Table 6")


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
    "생윤": ["A4:A52", "D4:D52", ws2],
    "윤사": ["F4:F52", "I4:I52", ws2],
    "한지": ["K4:K52", "N4:N52", ws2],
    "세지": ["P4:P52", "S4:S52", ws2],
    "동사": ["A3:A51", "D3:D51", ws3],
    "세사": ["F3:F51", "I3:I51", ws3],
    "경제": ["K3:K51", "N3:N51", ws3],
    "정법": ["P3:P51", "S3:S51", ws3],
    "사문": ["A3:A51", "D3:D51", ws4],
    "물1": ["A3:A51", "D3:D51", ws5],
    "화1": ["F3:F51", "I3:I51", ws5],
    "생1": ["K3:K51", "N3:N51", ws5],
    "지1": ["P3:P51", "S3:S51", ws5],
    "물2": ["A3:A51", "D3:D51", ws6],
    "화2": ["F3:F51", "I3:I51", ws6],
    "생2": ["K3:K51", "N3:N51", ws6],
    "지2": ["P3:P51", "S3:S51", ws6],
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