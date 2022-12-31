import requests, logging
from datetime import date,datetime
from .models import pdfFile
logger_debug = logging.getLogger("bot_tele_log")

def telegram_bot_sendtext(bot_message,bot_token,bot_chatID):
    
    # bot_token = '1333097128:AAEiN8KpM8g1jVrQ7rruXCt5IpSPd60OyaE'
    # bot_chatID = '-212348088'
    send_text = 'https://api.telegram.org/bot' + bot_token + '/sendMessage?chat_id=' + bot_chatID + '&parse_mode=HTML&text=' + bot_message

    response = requests.get(send_text)
    # print(response.json())
    if not response.json()["ok"]:
        print(bot_chatID)
        print(response.json())
        logger_debug.debug(bot_chatID)
        logger_debug.debug(response.json())
    return response.json()


# send message by project_manager_bot
def send_message_alert_project(message, chatID):
    token = '1161874277:AAGRwz3QGRAYBtwxupecbOdjtHVL3tCd_qg' # pvs_project_manager_bot

    telegram_bot_sendtext(message,token,chatID)

def kc_project_alert():
    chatID = "-404313724"
    year = datetime.today().year
    last_month = datetime.today().month - 1
    first_date = datetime(year, last_month, 1)
    first_date_new_month = datetime(year, datetime.today().month, 1)
    nums_files_month = len(pdfFile.objects.filter(SignedTime__date__gte=first_date.strftime("%Y-%m-%d"),SignedTime__date__lte=first_date_new_month.strftime("%Y-%m-%d"),signed=True))
    total_files= len(pdfFile.objects.filter(signed=True))
    mess = f'''
        https://psign.kcvn.vn
        <b>Tháng {last_month}</b> : {nums_files_month} văn bản đã kí
        <b>Tổng số văn bản</b> : {total_files} văn bản
    '''
    #chatID = "1670125478"
    send_message_alert_project(mess, chatID)