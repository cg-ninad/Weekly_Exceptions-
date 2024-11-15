# NOTE:
# Both libraries are used in this UseCase Outlook_mail & outlook_exch_lib
# Make sure mailbody is in HTML format else it will come as single line in oulook

try:
    import sys
    import uuid
    from logger_format import setup_logging
    from config_file import *
    from db_connection import *
    import pandas as pd
    from datetime import datetime,timedelta
    import time as t
    import os
    import shutil
    #from outlook_exch_lib import ExchangeMail
    from zipfile import *
    import pytz
    import Fileops
    from Fileops import *
    from Outlook_Mail import *

except ImportError as e:
    print(e)
    sys.exit(str(e))


class WeeklyExceptions:

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
        # self.exchangeobj = ExchangeMail(mail_sender,mail_user)  # initializing ExchangeMail from outlook_exch_lib library
        self.outlookexchange = Outlook_Mails(__file__,self.logger,self.guid)  # initializing outlookexchange object from outlook_mail library/class

    def WeeklyException(self, query, k, v, filename=''):
        try:
            cursor = cnx.cursor()
            self.logger.info("Connected to GFS database successfully")
            cursor.execute(common_query1)
            cursor.execute(common_query2)
            updated_query = query.format(str(v), last_week_date, current_week)
            # updated_query = query.format(str(v))  # TODO: This is temp for to make it for custom dates
            df = pd.read_sql_query(updated_query, cnx)   # firing query here
            if query == Revenue_query:
                filename = 'Revenue Exception'
            elif query == ICB_query:
                filename = 'ICB Exception'
            elif query == BNL_query:
                filename = 'BNL Exception'
            india_tzone = pytz.timezone("Asia/Kolkata")
            current_date = datetime.now(india_tzone)
            strp = current_date.strftime('%d-%m-%Y')
            file_name = strp + " " + filename + ' (' + k + ')' + '.xlsx'
            df.to_excel(os.path.join(root_folder, file_name),index=False)
            self.logger.info("Data is inserted into Excel")
            print(filename, 'for', k)
##            if df.empty:  # This function is commented as business dont need it
##                self.empty_df_check(k,filename)
##                print('Empty df found on {} for {}'.format(filename,k))
##                self.logger.info('{} found empty for {} region'.format(filename,k))

            # mail_subject_1 = '{} Region: Exceptions on ICB Revenue and B&L - Expenditure Items for {} month'.format(k,ongoing_month)
            # self.exchangeobj.SendMail(to, None, None, mail_subject_1, mail_body, None, False,True)  # if error is coming in sendmail function please to check logs
            # print(df)  # this will print Excel data (dataframe) to console

            # print(updated_query)
            # file_name = current_date + " " + filename + ' (' + k + ')' + '.txt'
            # f = open(file_name, 'w')
            # f.write(updated_query)
            # f.close()
            global counter
            global new_zip
            if not os.path.exists(k + '.zip'):
                new_zip = ZipFile(k + '.zip', 'w')
                counter = 1
            if os.path.exists(k + '.zip'):
                new_zip.write(file_name)
                counter = counter + 1
            if counter == 4:
                new_zip.close()
                # zip_file_path = os.path.abspath(k + '.zip')
                dirname = os.getcwd()
                f_name = os.path.join(dirname, k + '.zip')
                fl_name = [f_name]  # putting this in list to pass it as list for mail sending attachment
                if k in region_spoc_dict:
                    # mail_subject_1 = '{} Region: Exceptions on ICB Revenue and B&L - Expenditure Items for {} month'.format(k, ongoing_month)
                    mail_subject_1 = '{} Region: Exceptions on ICB Revenue and B&L - Expenditure Items from {} to {}'.format(k,last_week_date,current_week)
                    to = region_spoc_dict[k]
                    #self.exchangeobj.SendMail(to, cc, bcc_recipients, mail_subject_1, mail_body, fl_name, False,True)  # if error is coming in sendmail function please to check logs
                    self.outlookexchange.SendMail(sender=mail_sender,to=to,subject=mail_subject_1,mailBody=mail_body,attachment=fl_name,cc_str=cc,bcc_str=bcc_recipients)
                    print('{}.zip sent over mail'.format(k))
                    print()
                fileobj.copy_files_with_extension()

        except Exception as e:
            # self.exchangeobj.SendMail(error_to, error_cc, bcc_recipients, error_mail_subject, error_mail_body, [],False, True)
            self.outlookexchange.SendMail(sender=mail_sender, to=error_to, subject=error_mail_subject, mailBody=error_mail_body,attachment=None, cc_str=error_cc, bcc_str=bcc_recipients)
            self.logger.exception("Exception occurred in Weekly Exception {}".format(e))
            self.logger.info("----------------------End--------------------")
            print(e)
            sys.exit()

    def process(self):
        try:
            L = [Revenue_query,ICB_query,BNL_query]
            for k, v in my_dict.items():
                for query in L:
                    self.WeeklyException(query=query, k=k, v=v)
            print('Code execution completed')
        except Exception as e:
            # self.exchangeobj.SendMail(error_to, error_cc, bcc_recipients, error_mail_subject, error_mail_body, [],False, True)
            self.outlookexchange.SendMail(mail_sender,error_to,error_mail_subject,error_mail_body,None,error_cc,bcc_recipients)
            self.logger.exception("Exception occurred in process function {}".format(e))
            self.logger.info("----------------------End--------------------")
            print(e)
            exit()

##    def empty_df_check(self,k,filename):
##        df_mail_subject = '{} Region: Exceptions on ICB Revenue and B&L - Expenditure Items from {} to {} - No Data found '.format(k,last_week_date,current_week)
##        df_mail_body = 'This is to inform you that {} query returned no records for {} region.\n\nRegards,\nAutobot'.format(filename,k)
##        if k in region_spoc_dict:
##            to = region_spoc_dict[k]  #TODO: enabled this is Prod only
##            # self.exchangeobj.SendMail(to, cc, bcc_recipients, df_mail_subject,df_mail_body, [],False, True)
##            self.outlookexchange.SendMail(sender=mail_sender, to=to, subject=df_mail_subject, mailBody=df_mail_body,attachment=None, cc_str=cc, bcc_str=bcc_recipients)


if __name__ == '__main__':
    obj = WeeklyExceptions()
    obj.process()
