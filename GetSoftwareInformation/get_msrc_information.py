import datetime
import requests
import re
import const
from lxml import etree

# WebAPI関連
API_KEY = "8492a975d24a4ddb8af754e36792f76b"
MSRC_HEADERS = {"Api-Key": API_KEY}
MSRC_URL = "https://api.msrc.microsoft.com/cvrf/%s?api-Version=2017"
KB_LINK_URL = "https://support.microsoft.com/ja-jp/help/"

# MSRCFNAMESPACE
MSRC_NAMESPACE = {
    "cpe-lang": "http://cpe.mitre.org/language/2.0",
    "cvrf-common": "http://www.icasi.org/CVRF/schema/common/1.1",
    "dc": "http://purl.org/dc/elements/1.1/",
    "prod": "http://www.icasi.org/CVRF/schema/prod/1.1",
    "cvssv2": "https://scap.nist.gov/schema/cvss-v2/1.0",
    "vuln": "http://www.icasi.org/CVRF/schema/vuln/1.1",
    "sch": "http://purl.oclc.org/dsdl/schematron",
    "cvrf": "http://www.icasi.org/CVRF/schema/cvrf/1.1"
}

# 対象プロダクト(前rは正規表現使用)
TARGET_PRODUCTS = {
    "Internet Explorer 11 on Windows 10 Version 1607 for x64-based Systems",
    "Internet Explorer 11 on Windows 7 for 32-bit Systems Service Pack 1",
    r"Microsoft.+2010 Service Pack 2 \(32-bit editions\)",
    r"Microsoft.+2016 \(32-bit edition\)",
    r"Microsoft.+2016 \(64-bit edition\)"
}

# セキュリティタイトル変換
SECURITY_TITLE_TRANSLATE = {
    "Security Update": "セキュリティ更新プログラム",
    "IE Cumulative": "累積的なセキュリティ更新プログラム",
    "Monthly Rollup": "セキュリティ マンスリー品質ロールアップ"
}

# 深刻度変換
SEVERITY_TRANSLATE = {
    "Critical": "緊急",
	"Important": "重要",
	"Moderate": "警告",
    "Low": "注意"
}

# レポート用文字列
REPORT_TEMPLATE = {
    "Office Security Update": "Microsoft Office %s用のセキュリティ更新プログラム (%s)",
    "IE Security Update": "%s用 Internet Explorer 11 のセキュリティ更新プログラム (%s)",
    "IE Cumulative": "%s用 Internet Explorer 11 の累積的なセキュリティ更新プログラム (%s)",
    "Monthly Rollup": "セキュリティ マンスリー品質ロールアップ  (%s)"
}

# 対象プロダクト判別
def is_target_product(product):
    for target_product in TARGET_PRODUCTS:
        if re.match(target_product, product):
            return True
    return False

# セキュリティタイトル変換
def traslate_security_title(title):
    if title not in SECURITY_TITLE_TRANSLATE.keys():
        return title
    return SECURITY_TITLE_TRANSLATE[title]

# 深刻度変換
def traslate_severity(severity):
    if severity not in SEVERITY_TRANSLATE.keys():
        return severity
    return SEVERITY_TRANSLATE[severity]

# レポート用文字列の取得（コピペ用）
def create_repot_string(title, product, kb):
    if is_excel_security_update(title, product):
        return REPORT_TEMPLATE["Office Security Update"] % ("Excel", kb)

    if is_word_security_update(title, product):
        return REPORT_TEMPLATE["Office Security Update"] % ("Word", kb)

    if is_powerpoint_security_update(title, product):
        return REPORT_TEMPLATE["Office Security Update"] % ("PowerPoint", kb)

    if is_access_security_update(title, product):
        return REPORT_TEMPLATE["Office Security Update"] % ("Access", kb)

    if is_outlook_security_update(title, product):
        return REPORT_TEMPLATE["Office Security Update"] % ("Outlook", kb)

    if is_office_security_update(title, product):
        return REPORT_TEMPLATE["Office Security Update"] % ("", kb)

    if is_ie_win7_security_update(title, product):
        return REPORT_TEMPLATE["IE Security Update"] % ("Windows 7", kb)

    if is_ie_win10_security_update(title, product):
        return REPORT_TEMPLATE["IE Security Update"] % ("Windows 10", kb)

    if is_win7_ie_cumulative(title, product):
        return REPORT_TEMPLATE["IE Cumulative"] % ("Windows 7", kb)

    if is_win10_ie_cumulative(title, product):
        return REPORT_TEMPLATE["IE Cumulative"] % ("Windows 10", kb)

    if is_ie_win7_monthly_rollup(title, product):
        return REPORT_TEMPLATE["Monthly Rollup"] % (kb)
    
    if is_ie_win10_monthly_rollup(title, product):
        return REPORT_TEMPLATE["Monthly Rollup"] % (kb)

    return ""

# Excelセキュリティアップデート判別
def is_excel_security_update(title, product):
    return (title == "Security Update" and "Excel" in product)

# Wordセキュリティアップデート判別
def is_word_security_update(title, product):
    return (title == "Security Update" and "Word" in product)

# PowerPointセキュリティアップデート判別
def is_powerpoint_security_update(title, product):
    return (title == "Security Update" and "PowerPoint" in product)

# Accessセキュリティアップデート判別
def is_access_security_update(title, product):
    return (title == "Security Update" and "Access" in product)

# Outlookセキュリティアップデート判別
def is_outlook_security_update(title, product):
    return (title == "Security Update" and "Outlook" in product)

# Officeセキュリティアップデート判別
def is_office_security_update(title, product):
    return (title == "Security Update" and "Office" in product)

# Win7用IEセキュリティアップデート判別
def is_ie_win7_security_update(title, product):
    return (title == "Security Update" and "Internet Explorer 11 on Windows 7" in product)

# Win10用IEセキュリティアップデート判別
def is_ie_win10_security_update(title, product):
    return (title == "Security Update" and "Internet Explorer 11 on Windows 10" in product)

# Win7用IE累積的なセキュリティ判別
def is_win7_ie_cumulative(title, product):
    return (title == "IE Cumulative" and "Internet Explorer 11 on Windows 7" in product)

# Win10用IE累積的なセキュリティ判別
def is_win10_ie_cumulative(title, product):
    return (title == "IE Cumulative" and "Internet Explorer 11 on Windows 10" in product)

# Win7用IEマンスリー品質ロールアップ判別
def is_ie_win7_monthly_rollup(title, product):
    return (title == "Monthly Rollup" and "Internet Explorer 11 on Windows 7" in product)

# Win10用IEマンスリー品質ロールアップ判別
def is_ie_win10_monthly_rollup(title, product):
    return (title == "Monthly Rollup" and "Internet Explorer 11 on Windows 10" in product)

# メイン
def get_information():
    try:
        # WebApiよりデータを取得
        response = requests.get(MSRC_URL % "{0:%Y-%b}".format(datetime.date.today()), headers=MSRC_HEADERS, proxies=const.PROXIES)
        # response = requests.get("https://api.msrc.microsoft.com/cvrf/2018-Nov?api-Version=2017", headers=MSRC_HEADERS, proxies=const.PROXIES)
        if response.status_code != requests.codes.ok:
            return []

        # XML取得
        root = etree.XML(response.content)

        # 対象ProductIDディクショナリの作成
        target_products_dictionary = {}
        for fullproductname in root.findall("prod:ProductTree/prod:FullProductName", MSRC_NAMESPACE):
            if is_target_product(fullproductname.text):
                target_products_dictionary[fullproductname.get("ProductID")] = fullproductname.text

        # 脆弱性情報の取得
        msrc_list = []
        descriptions_productids = []
        for vulnerability in root.findall("vuln:Vulnerability", MSRC_NAMESPACE):
           # 日付の取得
            data_text = vulnerability.findall("vuln:RevisionHistory/vuln:Revision/cvrf:Date", MSRC_NAMESPACE)[0].text
            tdatetime = datetime.datetime.strptime(data_text[0:19], "%Y-%m-%dT%H:%M:%S")
            release_date = datetime.date(tdatetime.year, tdatetime.month, tdatetime.day)

            # Severity情報(深刻度)ディクショナリの作成
            severity_dictionary = {}
            for threat in vulnerability.findall("vuln:Threats/vuln:Threat", MSRC_NAMESPACE):
                if threat.get("Type") == "Severity":
                    product_id = threat.find("vuln:ProductID", MSRC_NAMESPACE)
                    description = threat.find("vuln:Description", MSRC_NAMESPACE)
                    severity_dictionary[product_id.text] = description.text

            # 修復ループ(CVEベース)
            for remediation in vulnerability.findall("vuln:Remediations/vuln:Remediation", MSRC_NAMESPACE):
                description = remediation.find("vuln:Description", MSRC_NAMESPACE)
                sub_type = remediation.find("vuln:SubType", MSRC_NAMESPACE)

                # セキュリティタイトルがないものはスルー
                if sub_type is None or description is None:
                    continue

                # プロダクトループ
                for product_id in remediation.findall("vuln:ProductID", MSRC_NAMESPACE):
                    if product_id.text in target_products_dictionary.keys():
                        # 同じ記事はまとめる
                        if description.text + "_" + product_id.text not in descriptions_productids:
                            msrc_list.append(
                                {
                                    "release_date": release_date.strftime("%Y-%m-%d"),
                                    "kb": description.text,
                                    "kb_link": KB_LINK_URL + description.text,
                                    "title": traslate_security_title(sub_type.text),
                                    "product": target_products_dictionary[product_id.text],
                                    "severity": traslate_severity(severity_dictionary[product_id.text]),
                                    "report": create_repot_string(sub_type.text, target_products_dictionary[product_id.text], description.text)
                                }
                            )
                            # 重複チェック用に記事を保存
                            descriptions_productids.append(
                                description.text + "_" + product_id.text)
        return msrc_list
    except:
        return []
