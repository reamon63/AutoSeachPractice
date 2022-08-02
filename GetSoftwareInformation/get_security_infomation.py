import requests
import const
from bs4 import BeautifulSoup
from lxml import html

IPA_LINK_URL = "https://www.ipa.go.jp"
MAX_INFOMATION_COUNT = 15

# メイン
def get_information():
    try:
        target_url = "https://www.ipa.go.jp/security/announce/alert.html"
        response = requests.get(target_url, proxies=const.PROXIES)
        if response.status_code != requests.codes.ok:
            return []

        soup = BeautifulSoup(response.content, "lxml")
        ipar_list = []
        count = 0
        for ipar_table in soup.find_all("table", class_="ipar_newstable"):
            for tr in ipar_table.find_all("tr"):
                th = tr.find("th")
                td = tr.find("td")
                a = td.find("a")
                ipar_list.append(
                    {
                        "release_date": th.text,
                        "title": td.text,
                        "link": IPA_LINK_URL + a.attrs["href"]
                    }
                )
                count += 1
                if count >= MAX_INFOMATION_COUNT:
                    return ipar_list
        return ipar_list
    except:
        return []
        