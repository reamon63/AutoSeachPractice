import datetime
import get_vendor_information
import get_msrc_information
import get_security_infomation
from jinja2 import Environment, FileSystemLoader

RESULT_TEMPLATE_FILE = "result_template.html"
RESULT_FILE = "result.html"

def main():
    # 作成日の取得
    created_date = "{0:%Y/%m/%d}".format(datetime.datetime.now())

    # 各ベンダーのバージョン情報を取得
    vendor_version = get_vendor_information.get_information()
    vendor_update = get_vendor_information.compare_last_data(vendor_version)

    # MSのセキュリティ情報の取得
    msrc_list = get_msrc_information.get_information()
    is_msrc_success = (len(msrc_list) > 0)

    # 重要なセキュリティ情報の取得
    ipar_list = get_security_infomation.get_information()
    is_ipar_success = (len(ipar_list) > 0)

    # 情報収集結果HTMLファイルの作成
    env = Environment(loader=FileSystemLoader("./", encoding="utf8"))
    result_template = env.get_template(RESULT_TEMPLATE_FILE)
    html = result_template.render(
        {
            "created_date": created_date,
            "vendor_version": vendor_version,
            "vendor_update": vendor_update,
            "is_msrc_success": is_msrc_success,
            "msrc_list": msrc_list,
            "is_ipar_success": is_ipar_success,
            "ipar_list": ipar_list
        }
    )
    result_file = open(RESULT_FILE, "w", encoding="utf8")
    result_file.write(html)
    result_file.close()

if __name__ == "__main__":
    main()