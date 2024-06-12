import time
import datetime
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
time1='2022-12-2-3-1'
time2='2022-12-2-3-10'
date1 = time.strptime(time1, "%Y-%m-%d-%H-%M")
date2 = time.strptime(time2, "%Y-%m-%d-%H-%M")

print(tyh.time_differ_5(time1,time2))