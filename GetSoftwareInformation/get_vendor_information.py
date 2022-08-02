import requests
import re
import const
from bs4 import BeautifulSoup
from lxml import html
import sqlite3
import datetime

DB_PATH = "software_information.sqlite3"
ERROR_MSG = "[ERROR] バージョン情報を取得出来ませんでした"

# Lhaplus
def get_lhaplus():
    try:
        target_url = "http://www7a.biglobe.ne.jp/~schezo/"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        all_p = soup.find_all("p")
        for p in all_p:
            if re.match("Lhaplus Version", p.text):
                return p.text
    except:
        return ERROR_MSG

# 7-zip
def get_7zip():
    try:
        target_url = "https://sevenzip.osdn.jp/sdk.html"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        release_date = soup.find_all("table")[0].find_all("tr")[1].find_all("td")[2]
        version = soup.find_all("table")[0].find_all("tr")[1].find_all("td")[3]
        return version.text + "[ " + release_date.text + " ]"
    except:
        return ERROR_MSG

# Becky!
def get_becky():
    try:
        target_url = "https://forest.watch.impress.co.jp/library/software/becky/"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        version_release_date = soup.find_all("dd")[0]
        return version_release_date.text
    except:
        return ERROR_MSG

# ATOK
def get_atok():
    try:
        #Pro 4
        target_url = "https://support.justsystems.com/faq/1032/app/servlet/qasearchtop?MAIN=002001003006004"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        pro4 = soup.find_all("table", summary="resultlist")
        #Pro4 と Medical2の更新情報が混在している為、Pro4の情報を絞り込み
        for i in pro4:
            if ("ATOK Pro 4" in i.text):
                pro4_info = i.find("a", target="faq_win_contents").text
                pro4_date = i.find(class_="day").text
                break
        pro4_result = "Pro 4" + " [ " + pro4_date + " ] " + "\n  " + pro4_info

        #Pro 3
        target_url = "https://support.justsystems.com/faq/1032/app/servlet/qasearchtop?MAIN=002001003006003"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        pro3 = soup.find_all("table", summary="resultlist")[0]
        pro3_info = pro3.find("a", target="faq_win_contents").text
        pro3_date = pro3.find(class_="day").text
        pro3_result = "Pro 3" + " [ " + pro3_date + " ] " + "\n  " + pro3_info
    
        return pro4_result + "\n\n" + pro3_result
    except:
        return ERROR_MSG

# 一太郎ビューア 
def get_ichitaro_viewer():
    try:
        target_url = "https://www.justsystems.com/jp/download/viewer/ichitaro"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        release_date = soup.find_all("table", class_="dl")[0].find_all("tr")[1].find_all("td")[0]
        version = soup.find_all("table", class_="dl")[0].find_all("tr")[1].find_all("td")[3]
        return version.text + " [ " + release_date.text + " ]"
    except:
        return ERROR_MSG

# MarkDiff
def get_markdiff():
    try:
        target_url = "http://rulebook.biz/markdiffj/contents/passcode_r.htm"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        table = soup.find_all("table")[1]
        version = table.find_all("tr")[len(table.find_all("tr")) - 1].find_all("td")[0]
        return version.text
    except:
        return ERROR_MSG

# Docuworks Viewer Light
def get_docuworksviewerlight():
    try:
        target_url = "https://www.fujifilm.com/fb/download/software/docuworks/download101"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        version = soup.find_all("table", class_="tbl")[1].find_all("tr")[1].find_all("td")[0]
        release_date = soup.find_all("table", class_="tbl")[1].find_all("tr")[2].find_all("td")[0]
        return version.text + " [ " + release_date.text + " ]"
    except:
        return ERROR_MSG

# Java (JRE)
def get_jre():
    try:
        # Ver 8
        target_url = "https://www.oracle.com/technetwork/java/javase/8u-relnotes-2225394.html"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        version8 = soup.find_all("table", class_="innerPgSignpost")[0].find_all("li")[0]

        # Ver 11
        target_url = "https://www.oracle.com/technetwork/java/javase/11u-relnotes-5093844.html"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        version11 = soup.find_all("table", class_="innerPgSignpost")[0].find_all("li")[0]

        return version8.text + "\n" + version11.text
    except:
        return ERROR_MSG

# Adobe Flash Player
def get_flashplayer():
    try:
        target_url = "https://get.adobe.com/jp/flashplayer/about/"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        version = soup.find_all("table", class_="data-bordered max")[0].find_all("tr")[1].find_all("td")[2]
        return version.text
    except:
        return ERROR_MSG

# Adobe Acrobat
def get_acrobat():
    try:
        version_info = ""

        # リリースノートサイトより取得
        target_url = "https://helpx.adobe.com/jp/acrobat/release-note/release-notes-acrobat-reader.html"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        table_2020 = soup.find_all("table")[1]
        tr_2020 = table_2020.find_all("tr")[1]
        date_2020 = tr_2020.find_all("td")[0]
        version_2020 = tr_2020.find_all("td")[1]
        table_2017 = soup.find_all("table")[2]
        tr_2017 = table_2017.find_all("tr")[1]
        date_2017 = tr_2017.find_all("td")[0]
        version_2017 = tr_2017.find_all("td")[1]
        version_info += "【Release Noteより取得】\n"
        version_info += version_2020.text.replace( '\n' , '' ) + " [ " + date_2020.text.replace( '\n' , '' ) + " ]\n" 
        version_info += version_2017.text.replace( '\n' , '' ) + " [ " + date_2017.text.replace( '\n' , '' ) + " ]"

        return version_info
    except:
        return ERROR_MSG

# Trend VirusBuster Corp 
def get_virusbustercorp():
    try:
        target_url = "http://downloadcenter.trendmicro.com/index.php?regs=jp&clk=latest&clkval=4634&lang_loc=13"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        all_version = ""
        for table in soup.find_all("table", class_="file_results"):
            # プロダクト・アップデート、Service Pack、Patch、検索エンジン
            version = table.find_all("tr")[0].find_all("td")[0]
            release_date = table.find_all("tr")[0].find_all("td")[1]
            all_version += "".join(version.text.split()) + " [ " + release_date.text + " ]\n"
        
        target_url = "https://downloadcenter.trendmicro.com/index.php?clk=tbl&clkval=5004&regs=jp&lang_loc=13"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        version_12 = ""
        for table in soup.find_all("table", class_="file_results"):
            # プロダクト・アップデート、Service Pack、Patch、検索エンジン
            version = table.find_all("tr")[0].find_all("td")[0]
            release_date = table.find_all("tr")[0].find_all("td")[1]
            version_12 += "".join(version.text.split()) + " [ " + release_date.text + " ]\n"
        
        return "Version 11" + "\n" + all_version + "\n" + "Version XG(12)" + "\n" + version_12
    except:
        return ERROR_MSG


# Microsoft Edge (Chromium-based)  
def get_edge():
    try:
        target_url = "https://docs.microsoft.com/en-us/deployedge/microsoft-edge-relnotes-security"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        security_version = ""
        security_version_info = soup.find("main", id="main")
        version = security_version_info.find(text=re.compile("Version"))
        release_date = security_version_info.find_all("h2")[0]
        version = str(version).split('(')[1].split(')')[0]
        security_version += version + " [ " + release_date.text + " ]\n"
        
        target_url = "https://docs.microsoft.com/en-us/deployedge/microsoft-edge-relnote-stable-channel"
        response = requests.get(target_url, proxies=const.PROXIES)
        soup = BeautifulSoup(response.content, "lxml")
        latest_version = ""
        latest_version_info = soup.find("main", id="main").find_all("h2")[0]
        latest_version = latest_version_info.text
        return "【セキュリティアップデート】" + "\n" + security_version + "\n" + "【最新バージョン】" + "\n" + latest_version
    except:
        return ERROR_MSG


# ベンダーヴァージョン情報を返す
def get_information():
    return {
        "lhaplus": get_lhaplus(),
        "seven_zip": get_7zip(),
        "becky":  get_becky(),
        "atok": get_atok(),
        "ichitaro_viewer": get_ichitaro_viewer(),
        "markdiff": get_markdiff(),
        "docuworksviewerlight": get_docuworksviewerlight(),
        "jre": get_jre(),
        "flashplayer": get_flashplayer(),
        "acrobat": get_acrobat(),
        "virusbustercorp": get_virusbustercorp(),
        "edge": get_edge()
    }

# 前回実施日の取得
def get_last_date():
    try:
        cn = sqlite3.connect(DB_PATH)
        cursor = cn.cursor()
        cursor.execute("SELECT * FROM vendor_data ORDER BY created_date DESC")
        row = cursor.fetchone()
        if row == None:
            cursor.close()
            cn.close()
            return ""

        return row[0]

        cursor.close()
        cn.close()
    except:
        return ""

# 前回データと比較しアップデートフラグを返す
def compare_last_data(current_data_dictionary):
    try:
        # 前回より更新されているか判別用フラグ
        update_flag_dictionary = {
            "lhaplus": True,
            "seven_zip": True,
            "becky": True,
            "atok": True,
            "ichitaro_viewer": True,
            "markdiff": True,
            "docuworksviewerlight": True,
            "jre": True,
            "flashplayer": True,
            "acrobat": True,
            "virusbustercorp": True,
            "edge": True
        }

        cn = sqlite3.connect(DB_PATH)
        cursor = cn.cursor()

        # 前回データが存在する場合は取得しアップデートフラグをセットする
        last_date = get_last_date()
        if last_date != "":
            cursor.execute("SELECT * FROM vendor_data WHERE created_date = ? ORDER BY created_date DESC", [last_date])
            for row in cursor.fetchall():
                update_flag_dictionary[row[1]] = (current_data_dictionary[row[1]] != row[2])
            cursor.close()

        # データの更新
        created_date = "{0:%Y%m%d}".format(datetime.datetime.now())
        update_vendor_data = [
            (created_date, "lhaplus", current_data_dictionary["lhaplus"]),
            (created_date, "seven_zip", current_data_dictionary["seven_zip"]),
            (created_date, "becky", current_data_dictionary["becky"]),
            (created_date, "atok", current_data_dictionary["atok"]),
            (created_date, "ichitaro_viewer", current_data_dictionary["ichitaro_viewer"]),
            (created_date, "markdiff", current_data_dictionary["markdiff"]),
            (created_date, "docuworksviewerlight", current_data_dictionary["docuworksviewerlight"]),
            (created_date, "jre", current_data_dictionary["jre"]),
            (created_date, "flashplayer", current_data_dictionary["flashplayer"]),
            (created_date, "acrobat", current_data_dictionary["acrobat"]),
            (created_date, "virusbustercorp", current_data_dictionary["virusbustercorp"]),
            (created_date, "edge", current_data_dictionary["edge"])
        ]
        cursor = cn.cursor()
        cursor.executemany("INSERT OR REPLACE INTO vendor_data(created_date, vendor_id, version_text) VALUES (?,?,?)", update_vendor_data)
        cn.commit()

        cursor.close()
        cn.close()

        return update_flag_dictionary
    except:
        return {
            "lhaplus": True,
            "seven_zip": True,
            "becky": True,
            "atok": True,
            "ichitaro_viewer": True,
            "markdiff": True,
            "docuworksviewerlight": True,
            "jre": True,
            "flashplayer": True,
            "acrobat": True,
            "virusbustercorp": True,
            "edge": True
        }
