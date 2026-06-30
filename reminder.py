#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
淡信每月作業 LINE 提醒小幫手
------------------------------------------------
每天執行一次。當天若有對應待辦，就整理成一則訊息推播到你的 LINE。

假日判斷優先序：
  1. 同資料夾的 holidays.json（官方辦公日曆表轉出，最準）涵蓋的年份用它。
  2. 沒涵蓋的年份用內建後備：週末＋固定國定假日＋春節區間（2026 起已無補班/彈性放假）。
  新年度官方檔出來後，把 xlsx 交給我重產 holidays.json 即可，程式不用動。

假日順延：
  ・「X 號前要完成」的死線（6、7、15、25 號、月底）→ 遇假日「提前」到前一個上班日。
  ・「月初」（1～5 號）→ 遇假日「往後」順延，避免跨到上個月。

執行方式：
  python reminder.py            正式推播
  python reminder.py --dry      只在畫面預覽，不推播（測試用）
  python reminder.py --dry --date 20260213   假裝某天來預覽
"""
import os
import sys
import json
import datetime
from calendar import monthrange
import urllib.request
import urllib.error

# ── LINE 設定：用環境變數帶入，不要把金鑰寫死在程式裡 ──
LINE_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
USER_ID = os.environ.get("LINE_USER_ID", "")

# ── 台灣時間 ──
TW = datetime.timezone(datetime.timedelta(hours=8))

HERE = os.path.dirname(os.path.abspath(__file__))


# ============================================================
#  待辦清單：要改時間或內容，只動這一段就好
#  when 可填：
#    數字 15        → 每月 15 號
#    多個 [6, 7]    → 6 號和 7 號
#    "end"          → 當月最後一天
#    "end-2"        → 月底前 2 天
#  順延方向自動判斷（月初往後、其餘提前）。
#  若想強制某條的方向，加 "roll": "fwd"（往後）或 "back"（提前）。
# ============================================================
TASKS = [
    {
        "when": [1],
        "title": "月初：放款統計表",
        "body": "志偉 3 號前會給放款資料。\n"
                "整理到 D:\\企劃業務作業\\放款統計表\\115，做好後交張副理。",
    },
    {
        "when": [1],
        "title": "月初：聯名卡推廣數量",
        "body": "依各分社統計聯名卡推廣數量，交張副理。\n"
                "來源 D:\\企劃業務作業\\聯名卡\\進件彙總表\\115年",
    },
    {
        "when": [3],
        "title": "放款資料確認",
        "body": "志偉的放款資料今天前應到位，整理後交張副理。",
    },
    {
        "when": [6],
        "title": "查平均利率",
        "body": "上央行查平均利率：\n"
                "https://cpx.cbc.gov.tw/BIRWeb/Range/RangeSelect",
    },
    {
        "when": [6, 7],
        "title": "寄聯邦收據",
        "body": "7 號前要寄聯邦收據。\n"
                "開 D:\\企劃業務作業\\聯名卡\\收據\\消費回饋金收據_115年開始.xlsx\n"
                "在「紀錄」填寫，回填「消費回饋金收據」的對應數字，再印出來。",
    },
    {
        "when": [13],
        "title": "提前提醒：15 號待辦",
        "body": "兩天後（15 號）要做：\n"
                "・上傳銀行局數字（子儀會寄資料），並確認副理審核\n"
                "・查 16-11-3891010 餘額轉帳給聯邦",
    },
    {
        "when": [15],
        "title": "上傳銀行局數字",
        "body": "子儀會寄「金融機構行動金融卡辦理情形」。\n"
                "上傳到 ebank.banking.gov.tw，並確認副理也審核過數字。\n"
                "（15 號前要完成）",
    },
    {
        "when": [15],
        "title": "餘額轉帳聯邦（月中）",
        "body": "查 16-11-3891010 餘額，轉帳給聯邦（內扣 20 元手續費）。\n"
                "・記錄在 D:\\企劃業務作業\\聯名卡\\轉帳紀錄\n"
                "・傳真給財務管理科 許小姐 87526139、發卡帳務科 藍小姐 87526256\n"
                "・匯款申請書第二聯留著，月底跟月中那張一起寄給聯邦",
    },
    {
        "when": [15],
        "title": "富保佣金清單作業",
        "body": "富保約這幾天寄佣金清單，存到\n"
                "D:\\企劃業務作業\\保險\\淡信每月佣金明細\\銀保保代每月佣金明細\n"
                "接著：\n"
                "・獎金明細印 2 張\n"
                "・寫發票號碼那張交給思瑜\n"
                "・會計向銀行局申報辦理保險資料\n"
                "・保險存摺 16-11-3087607 開取款條兩張（銀保、富保），大章用「金融商品專用章」\n"
                "・25 號確認佣金有入帳後，思瑜會開兩張發票，再寄給瑜璇",
    },
    {
        "when": [25],
        "title": "佣金入帳與轉付",
        "body": "25 號保代和聯名卡會把佣金入帳。\n"
                "・確認入帳後，思瑜開兩張發票寄給瑜璇\n"
                "・若有車險、其他險，把佣金轉給相對的同事",
    },
    {
        "when": [25],
        "title": "寄發票＋簽到表給富保",
        "body": "上課簽到表這時應已收齊（紀錄在 LINE 群組記事本「合作推廣」）。\n"
                "把思瑜開的發票連同簽到表一起寄給富保。\n"
                "・寄出前登記在 D:\\企劃業務作業\\保險\\保代送件紀錄表.xlsx\n"
                "・用便利袋，貼上「銀保保代」貼紙\n"
                "・通知便利帶公司取件：member.25431010.tw，帳號 138730",
    },
    {
        "when": ["end"],
        "title": "餘額轉帳聯邦（月底）",
        "body": "查 16-11-3891010 餘額，轉帳給聯邦（內扣 20 元手續費）。\n"
                "・記錄在 D:\\企劃業務作業\\聯名卡\\轉帳紀錄\n"
                "・傳真給財務管理科 許小姐 87526139、發卡帳務科 藍小姐 87526256\n"
                "・把月中＋月底兩張匯款申請書第二聯，一起用郵局寄給聯邦 蔡玉婷小姐（有貼紙可用）",
    },
]


# ============================================================
#  以下是程式邏輯，平常不用動
# ============================================================
def load_holidays():
    """讀同資料夾的 holidays.json（官方辦公日曆表轉出）。
    回傳 (放假日集合, 有涵蓋的年份集合)。讀不到則回 (None, set())，改用週末判斷。"""
    path = os.path.join(HERE, "holidays.json")
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        offdays, years = set(), set()
        for year, days in data.items():
            years.add(int(year))
            for ds in days:
                y, m, d = ds.split("-")
                offdays.add(datetime.date(int(y), int(m), int(d)))
        return offdays, years
    except Exception as e:
        print("（讀不到 holidays.json，改用週末判斷假日）", e)
        return None, set()


# 各年「大年初一」國曆日期（後備推算用，已查證）
LUNAR_NEW_YEAR = {
    2026: (2, 17), 2027: (2, 6), 2028: (1, 26), 2029: (2, 13),
    2030: (2, 3), 2031: (1, 23), 2032: (2, 11), 2033: (1, 31), 2034: (2, 19),
}
# 固定國定假日（2025 修法後實際放假者）：元旦、和平、兒童、清明、勞動、教師、
# 國慶、光復、行憲（清明取 4/4、4/5 兩天涵蓋）
FIXED_HOLIDAYS = {(1, 1), (2, 28), (4, 4), (4, 5), (5, 1),
                  (9, 28), (10, 10), (10, 25), (12, 25)}


def fallback_is_holiday(d):
    """沒有官方檔年份的後備判斷：週末＋固定國定假日＋春節區間"""
    if d.weekday() >= 5:
        return True
    if (d.month, d.day) in FIXED_HOLIDAYS:
        return True
    if d.year in LUNAR_NEW_YEAR:
        m, day = LUNAR_NEW_YEAR[d.year]
        cny = datetime.date(d.year, m, day)
        # 除夕(初一-1) 到 初三後一天(初一+3) 視為春節假
        if cny - datetime.timedelta(1) <= d <= cny + datetime.timedelta(3):
            return True
    return False


def is_holiday(d, holidays, years):
    # 該年份有官方資料就用官方的；沒有就用後備推算
    if holidays is not None and d.year in years:
        return d in holidays
    return fallback_is_holiday(d)


def roll_step(token):
    """自動判斷順延方向：月底家族與一般死線往前(-1)，月初(1~5號)往後(+1)"""
    if token == "end" or (isinstance(token, str) and token.startswith("end-")):
        return -1
    if isinstance(token, int) and token <= 5:
        return +1
    return -1


def roll(d, holidays, years, step):
    guard = 0
    while is_holiday(d, holidays, years) and guard < 25:
        d += datetime.timedelta(days=step)
        guard += 1
    return d


def effective_dates(task, today, last_day, holidays, years):
    y, m = today.year, today.month
    override = {"fwd": +1, "back": -1}.get(task.get("roll"))
    out = set()
    for token in task["when"]:
        step = override if override is not None else roll_step(token)
        if token == "end":
            base = datetime.date(y, m, last_day)
        elif isinstance(token, str) and token.startswith("end-"):
            base = datetime.date(y, m, last_day - int(token.split("-")[1]))
        else:
            base = datetime.date(y, m, int(token))
        out.add(roll(base, holidays, years, step))
    return out


def build_message(today, holidays, years):
    last_day = monthrange(today.year, today.month)[1]
    fired = [t for t in TASKS
             if today in effective_dates(t, today, last_day, holidays, years)]
    if not fired:
        return None
    roc = today.year - 1911
    lines = [f"📋 淡信作業提醒  {today.month}/{today.day}",
             f"（民國 {roc} 年 {today.month} 月）", ""]
    for t in fired:
        lines.append(f"🔸 {t['title']}")
        lines.append(t["body"])
        lines.append("")
    return "\n".join(lines).strip()


def _post(url, payload):
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode("utf-8"), method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Authorization", f"Bearer {LINE_TOKEN}")
    try:
        with urllib.request.urlopen(req) as resp:
            print("已送出，HTTP", resp.status)
    except urllib.error.HTTPError as e:
        print("送出失敗：", e.code, e.read().decode())
        sys.exit(1)


def push(text):
    """收件人由環境變數 LINE_USER_ID 決定：
       ・單一 ID          → 發給一個人
       ・多個 ID 用逗號分隔 → 發給多個人（multicast）
       ・設成 ALL         → 廣播給所有加好友的人（broadcast）"""
    if not LINE_TOKEN:
        print("缺少 LINE_CHANNEL_ACCESS_TOKEN 環境變數")
        sys.exit(1)
    msg = [{"type": "text", "text": text}]
    target = USER_ID.strip()
    if target.upper() == "ALL":
        _post("https://api.line.me/v2/bot/message/broadcast", {"messages": msg})
        return
    ids = [x.strip() for x in target.split(",") if x.strip()]
    if not ids:
        print("缺少 LINE_USER_ID 環境變數")
        sys.exit(1)
    if len(ids) == 1:
        _post("https://api.line.me/v2/bot/message/push",
              {"to": ids[0], "messages": msg})
    else:
        _post("https://api.line.me/v2/bot/message/multicast",
              {"to": ids, "messages": msg})


def main():
    dry = "--dry" in sys.argv
    has_date = "--date" in sys.argv
    today = datetime.datetime.now(TW).date()
    if has_date:
        s = sys.argv[sys.argv.index("--date") + 1]
        today = datetime.date(int(s[:4]), int(s[4:6]), int(s[6:8]))

    holidays, years = load_holidays()
    msg = build_message(today, holidays, years)

    if msg is None:
        print(f"{today} 今天沒有待辦，不推播。")
        return
    if dry:
        print("----- 預覽 -----")
        print(msg)
        print("----------------")
        return

    # 真實推播。為了配合雲端「一天排多個備援時段」，這裡用 last_sent.txt
    # 記住今天是否已發過，確保一天只發一次、不會重複。
    # （--date 是測試用，不做去重、也不寫紀錄）
    state_path = os.path.join(HERE, "last_sent.txt")
    if not has_date:
        try:
            last = open(state_path, encoding="utf-8").read().strip()
        except Exception:
            last = ""
        if last == today.isoformat():
            print(f"{today} 今天已經發送過，略過（避免重複）。")
            return

    push(msg)

    if not has_date:
        try:
            with open(state_path, "w", encoding="utf-8") as f:
                f.write(today.isoformat())
        except Exception as e:
            print("（寫入發送紀錄失敗，不影響推播）", e)


if __name__ == "__main__":
    main()
