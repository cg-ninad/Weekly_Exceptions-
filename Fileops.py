try:
    import logging
    import os
    import sys
    from datetime import datetime
    import shutil
    from config_file import *
    import time as t
    from logger_format import setup_logging
    import uuid
    from outlook_exch_lib import ExchangeMail
    import pytz

except ImportError as e:
    print(e)
    sys.exit()


class Fileops:
    def __init__(self, logFolder=None, logger=None, guid=None):
        """
            Configuration to start the process...
        """
        self.pstart_time = t.time()
        if guid:
            self.guid = guid
        else:
            self.guid = uuid.uuid4()
        if logFolder:
            self.logger = logger
        else:
            self.logger = setup_logging(__file__)
        self.logger.info("----------------------Start-----------------------")
        self.exchangeobj = ExchangeMail(mail_sender,mail_user)  # initializing ExchangeMail() from outlook_exch_lib library

    def copy_files_with_extension(self,src_folder=root_folder, txt_extension='.xlsx', xlsx_extension2='.zip'):
        if not os.path.exists('Archive'):
            os.mkdir('Archive')

        india_timezone = pytz.timezone('Asia/Kolkata')
        tdate = datetime.now(india_timezone)
        todays_date = tdate.strftime('%d-%m-%Y')

        if not os.path.exists('Archive\\' + todays_date):
            os.makedirs("Archive\\" + todays_date)

        real_dst = os.path.join(source, todays_date)  # this is target folder to move xlsx & zip files

        try:
            for file in os.listdir(src_folder):
                source_file = os.path.join(root_folder, file)
                if os.path.isfile(source_file) and file.endswith(txt_extension) or file.endswith(xlsx_extension2):
                    # shutil.move(source_file, target_folder)
                    shutil.move(source_file, os.path.join(real_dst, file))

        except Exception as e:
            # self.exchangeobj.SendMail(error_to, error_cc, bcc_recipients, error_mail_subject, error_mail_body, [],False, True)
            self.logger.exception("Exception occurred in Fileops {}".format(e))
            print(e)
            sys.exit()


fileobj = Fileops()
