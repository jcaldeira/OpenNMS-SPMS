import datetime



datastr="30/04/2021"

date_ = datetime.datetime.strptime(datastr,'%d/%m/%Y')

date_.year
date_.month
date_.day
date_.minute
dir(date_)


date_ ==  datetime.datetime.strptime('30/04/2021' ,'%d/%m/%Y')
