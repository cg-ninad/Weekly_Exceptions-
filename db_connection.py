import pyodbc

# db will work only in Corp\Gagent1 login in server i.e. PRB01P or any other

try:
    # CON_STR1 = 'DSN=GFS_UAT;UID=apps;PWD=appsru'  # Verified UAT instance
    CON_STR1 = 'DSN=GDHPROD;USR_AUTOBOT;PWD=AutoboT##12345671'  # Verified Prod instance
    cnx = pyodbc.connect(CON_STR1)
    print(cnx)

except Exception as e:
    print(e)
