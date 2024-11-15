from datetime import datetime, timedelta

# current_date = datetime.now().strftime("%d-%b-%Y")
# _last_week_date = datetime.now()-timedelta(weeks=1)
# last_week_date = _last_week_date.date().strftime('%d-%b-%Y')

# print(current_date)
# print(last_week_date)

# current_date = datetime.now().strftime("%d-%m-%Y")

current_week = datetime.now().strftime("%d-%B-%Y")

_last_week_date = datetime.now() - timedelta(weeks=1)  # for week calculations
last_week_date = _last_week_date.date().strftime('%d-%B-%Y')  # for week calculations

ongoing_month = datetime.now().strftime('%B')

common_query1 = """alter session set current_schema=apps;"""
common_query2 = """alter session set nls_language = 'AMERICAN';"""

# NOTE: DO NOT MODIFY BELOW QUERIES
Revenue_query = """select distinct a.REV_DIST_REJECTION_CODE, 
    a.creation_date ,
    a.expenditure_item_id,
    a.expenditure_item_date,
    a.expenditure_type,
    e.employee_number,
    e.full_name employee_name,
    a.quantity ,
    a.system_linkage_function ,
    a.net_zero_adjustment_flag ,
    a.CC_IC_PROCESSED_CODE "Intercompany_Billing_Processed",
    a.CC_CROSS_CHARGE_TYPE "Cross_Charge_Proc_Method",
    a.COST_DISTRIBUTED_FLAG,
    a.BILLABLE_FLAG,
    a.BILL_HOLD_FLAG,
    c.segment1 project_code,
    b.name Project_OU,
    hro.name proj_PU,
    exp_OU.name EMP_OU,
    hro_emp.attribute6 emp_PU ,
    c.project_status_code ,
    a.REVENUE_DISTRIBUTED_FLAG,
    a.BILL_RATE,
    c.start_date,
    c.completion_date , 
    c.CLOSED_DATE 
from 
apps.pa_expenditure_items_all a,
apps.pa_expenditures_all d,
apps.per_all_people_f e,
apps.pa_expenditures_all pea, 
apps.hr_operating_units b, 
apps.hr_operating_units exp_OU, 
apps.hr_all_organization_units hro,
apps.hr_all_organization_units hro_emp,
apps.pa_projects_all c 
where a.expenditure_id = d.expenditure_id
and e.person_id = d.incurred_by_person_id 
and c.org_id = b.organization_id
and pea.expenditure_id = a.expenditure_id
and c.CARRYING_OUT_ORGANIZATION_ID = hro.organization_id
and a.project_id = c.project_id
and b.organization_id in {}
and exp_OU.organization_id  = a.org_id
and a.COST_DISTRIBUTED_FLAG = 'Y'
and a.BILLABLE_FLAG = 'Y'
and a.BILL_HOLD_FLAG = 'N'
and nvl(a.revenue_hold_flag,'N') <> 'Y'
and c.revenue_accrual_method = 'WORK'
and nvl(hro.date_to, sysdate+1) > sysdate
and nvl(hro_emp.date_to, sysdate+1) > sysdate
and pea.incurred_by_organization_id = hro_emp.organization_id
and a.REVENUE_DISTRIBUTED_FLAG = 'N'
and a.REV_DIST_REJECTION_CODE is not null
and a.CREATION_DATE >= to_date('{}','dd-mon-yyyy')
and a.CREATION_DATE <= to_date('{}','dd-mon-yyyy');"""

ICB_query = """SELECT /*+ FULL(peia)*/
    peia.expenditure_item_id,
    peia.expenditure_item_date,
    peia.creation_date,
    peia.quantity,
    ppa.segment1,
    ppa.project_status_code,
    ppa.start_date,
    ppa.completion_date,
    peia.cc_cross_charge_code,
    peia.cc_cross_charge_type,
    peia.cc_bl_distributed_code,
    peia.CC_REJECTION_CODE, peia.raw_cost,
    peia.PROJECT_TRANSFER_PRICE,
    haou1.name   provider_org,
    haou2.name   receiver_org,
    peia.CONVERTED_FLAG,
    peia.NET_ZERO_ADJUSTMENT_FLAG
FROM
    apps.pa_expenditure_items_all    peia,
    apps.hr_operating_units          hou,
    apps.hr_all_organization_units   haou1,
    apps.hr_all_organization_units   haou2,
    apps.pa_projects_all             ppa
WHERE
    peia.org_id = hou.organization_id
    AND peia.project_id = ppa.project_id
    AND peia.org_id in {}
    and nvl(haou1.date_to, sysdate+1) > sysdate
    and nvl(haou2.date_to, sysdate+1) > sysdate
    AND peia.cc_prvdr_organization_id = haou1.organization_id
    AND peia.cc_recvr_organization_id = haou2.organization_id
     AND peia.cc_cross_charge_code = 'I'
   AND peia.CC_REJECTION_CODE is not null
   and peia.CREATION_DATE >= to_date('{}','dd-mon-yyyy')
   and peia.CREATION_DATE <= to_date('{}','dd-mon-yyyy');"""

BNL_query = """SELECT /*+ FULL(peia)*/
    peia.expenditure_item_id,
    peia.expenditure_item_date,
    peia.creation_date,
    peia.quantity,
    ppa.segment1,
    ppa.project_status_code,
    ppa.start_date,
    ppa.completion_date,
    peia.cc_cross_charge_code,
    peia.cc_cross_charge_type,
    peia.cc_bl_distributed_code,
    peia.CC_REJECTION_CODE,
    peia.creation_date,
    haou1.name   provider_org,
    haou2.name   receiver_org
FROM
    apps.pa_expenditure_items_all    peia,
    apps.hr_operating_units          hou,
    apps.hr_all_organization_units   haou1,
    apps.hr_all_organization_units   haou2,
    apps.pa_projects_all             ppa
WHERE
    peia.org_id = hou.organization_id
    AND peia.project_id = ppa.project_id
    AND peia.org_id in {}
    AND peia.cc_prvdr_organization_id = haou1.organization_id
    AND peia.cc_recvr_organization_id = haou2.organization_id
    AND peia.cc_cross_charge_code = 'B'
    and peia.CREATION_DATE >= to_date('{}','dd-mon-yyyy')
    and peia.CREATION_DATE <= to_date('{}','dd-mon-yyyy')
    and nvl(haou1.date_to, sysdate+1) > sysdate
    and nvl(haou2.date_to, sysdate+1) > sysdate
    and peia.CC_REJECTION_CODE is not null;"""

# below is testing dict

##my_dict = {'APAC': (68399,68359,142086,142082,142100,142020,142061,108259,86701,86700,86699),
##            'AUSTRIA': ((55687)),
##            'BELUX': (13299, 16855)}

# below is the original region dict
my_dict = {'APAC': (68399,68359,142086,142082,142100,142020,142061,108259,86701,86700,86699),
           'AUSTRIA': ((55687)),
           'BELUX': (13299, 16855),
           'FRANCE': (146256,146253,3715,6144,6147,6152,12544,117719,146249,146254,146260,146300),
           'FS': (58849,55787,30263,8758,5195,5219,12816,6456,8759,11169,14905,5198,5196,8760,5218,24067,18961),
           'GERMANY': (141983,141981,114740,74599,75099,31348,13067,7976),
           'IBERIA': (137820, 137780, 145620, 8116),
           'INDIA': (
               193621, 193620, 189261, 189260, 181643, 181642, 181641, 181640, 156061,
                              156060,146247,146246,146245,146244,146243,146242,146241,146240,
                              143420,120281,120280,120279,117105,117104,117103,117101,117100,
                              117099,117079,135860,115779,117117,117116,117115,5636,5637,5638,
                              117060,117059,73199,67639,64379,66099,61967,49967,45569,45067,
                              45167,34374,34373,34371,34370,34369,34367,34366,34365,34364,34363,
                              34362,34361,103,104,105,109,2715,7415,16011,19136,19507,19485,22687,
                              26688,24827,24828,25927,25987,26789,26791,34354,34355,34356,34358,34360),
           'IRELAND': ((75899)),
           'ITALY': (120139, 7977),
           'LATAM': (55847,8077,173141,173081,173080),
           'MOROCCO': (175342,175340,122879),
           'NAR AOS': (11039,8239,3942,3941,12323,8244,8242,3939,126142,59528),
           'NL': (114579,13298,13294,13297,13295),
           'NORDIC': (137720,3996,3995,3980,126144,3997,5055,177280,3979),
           'POLAND': ((12105)),
           'ROMANIA': ((55887)),
           'SWITZERLAND': ((11409)),
           'UK': (126140, 19712)}


region_spoc_dict = {'BELUX': 'jean-claude.maes@capgemini.com$^$gfsl1.belux@capgemini.com',
                    'FRANCE': 'laurent.he@capgemini.com$^$brice.ruffinoni@capgemini.com$^$gfs-support-l1-com.fr@capgemini.com',
                    'FS': 'swapnil.rathod@capgemini.com$^$sivannarayana@capgemini.com$^$udhayakumar.ravi@capgemini.com$^$gfsgovernanceteamfinance.fssbu@capgemini.com',
                    'SWITZERLAND': 'maria.loch-bednarczyk@capgemini.com$^$ewelina.gaska@capgemini.com$^$businesssupportswitzerland.pl@capgemini.com$^$expenses.zh.ch@capgemini.com',
                    'APAC': 'stella.tang@capgemini.com$^$sarah.magee@capgemini.com$^$catherine.costa@capgemini.com$^$mayuresh.kapre@capgemini.com$^$anant.danve@capgemini.com$^$biswajit.mohapatra@capgemini.com$^$geeta.rathod@capgemini.com$^$gfssupport.au@capgemini.com',
                    'INDIA': 'lalit.bagwe@capgemini.com$^$debtanu.chakrabortty@capgemini.com$^$funcindial1.in@capgemini.com',
                    'GERMANY': 'sebastian.sanner@capgemini.com$^$ulrike.baar@capgemini.com$^$harald.bornhardt@capgemini.com$^$financebusinesssupportl1.de@capgemini.com',
                    'LATAM': 'mariadelrocio.hernandez@capgemini.com$^$l1_support.mx@capgemini.com$^$gfsl1.ar@capgemini.com$^$byron.rivas@capgemini.com',
                    'IBERIA': 'alberto.oliva-gonzalez@capgemini.com$^$jorge.romano@capgemini.com$^$fernando-miguel.cano-torres@capgemini.com$^$gfsl1.ar@capgemini.com$^$l1.support-gfs.es@capgemini.com$^$ptsugfs.pl@capgemini.com',
                    'MOROCCO': 'brice.ruffinoni@capgemini.com$^$laurent.he@capgemini.com$^$gfs-support-l1-com.fr@capgemini.com',
                    'IRELAND': 'Esther.Prats@capgemini.com$^$urmi.sarkar@capgemini.com$^$finance.systemsteam.ie@capgemini.com$^$finance.systemsteam.i.e@capgemini.com',
                    'AUSTRIA': 'sabina.klejdysz@capgemini.com$^$atsu.pl@capgemini.com',
                    'ROMANIA': 'daniel.holda@capgemini.com$^$nica.pipota@capgemini.com$^$pa.ro@capgemini.com',
                    'NAR AOS': 'abhijit.guha@capgemini.com$^$andrea.fernandes@capgemini.com$^$gfsl1support.nar@capgemini.com',
                    'ITALY': 'alessandro.bottacchiari@capgemini.com$^$gfssupportl1.it@capgemini.com',
                    'NL': 'rico.collenburg@capgemini.com$^$rob.vanden.berg@capgemini.com$^$financespocs@capgemini.com$^$gfs_su.nl@capgemini.com',
                    'NORDIC': 'pawel.grzegorzek@capgemini.com$^$jimmy.munther@sogeti.se$^$sugfs.se@capgemini.com',
                    'POLAND': 'daniel.holda@capgemini.com$^$beata.katulka@capgemini.com$^$plsugfs.pl@capgemini.com',
                    'UK': 'Esther.Prats@capgemini.com$^$urmi.sarkar@capgemini.com'}

root_folder = r"E:\PythonDeployments\WeeklyException_ICB_B&L_Revenue"
source = r"E:\PythonDeployments\WeeklyException_ICB_B&L_Revenue\Archive"

# database = ''
# host = ''

# -------------Mailing Details--------------
mail_sender = 'gitauto1.agent1@capgemini.com'
mail_user = 'gitauto1.agent1@capgemini.com'


# --------------------------Success Mail------------------------------
#to = 'ninad-kiran.deshpande@capgemini.com$^$anshika.f.singh@capgemini.com'
#cc = ['ninad-kiran.deshpande@capgemini.com','anshika.f.singh@capgemini.com']
#cc = 'ninad-kiran.deshpande@capgemini.com$^$poonam.bajaj@capgemini.com'
bcc_recipients = 'ninad-kiran.deshpande@capgemini.com'

# Business people ID's
#to = ['poonam.bajaj@capgemini.com','neeta-pradeep.pawar@capgemini.com','r-bhavani.sree@capgemini.com']  # this is for outlook_exch_lib library
#cc = ['gaoraclep2c.in@capgemini.com','groupfinancept.in@capgemini.com','prashant.choudhari@capgemini.com','anita.yadav@capgemini.com']
#to = 'poonam.bajaj@capgemini.com$^$neeta-pradeep.pawar@capgemini.com$^$r-bhavani.sree@capgemini.com'  # $^$ is added for outlook_mail library
cc = 'gaoraclep2c.in@capgemini.com$^$groupfinancept.in@capgemini.com$^$prashant.choudhari@capgemini.com$^$anita.yadav@capgemini.com$^$ninad-kiran.deshpande@capgemini.com$^$anshika.f.singh@capgemini.com$^$poonam.bajaj@capgemini.com'

mail_body = '''<html>Hi team,<p>Please find the attached Exceptions on ICB,Revenue & B&L transactions</p>Regards,<br>AutoBot</html>'''

# --------------------------Error Mail---------------------------------
error_to = 'ninad-kiran.deshpande@capgemini.com'
error_cc = 'ninad-kiran.deshpande@capgemini.com'
error_bcc = 'ninad-kiran.deshpande@capgemini.com'
error_mail_subject = 'Failed: Weekly Exception Internal error.'
db_error_mail_subject = 'Failed: Weekly Exception empty database'
db_mail_body = '''<html>Hi team<p>This is to inform you that Database is empty</p>Regards,<br>Autobot</html>'''
error_mail_body = '''<html>Hi team<p>This is to inform you that Bot process failed due to some internal error</p>Regards,<br>AutoBot</html>'''


