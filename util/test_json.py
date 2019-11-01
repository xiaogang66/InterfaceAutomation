import re
import json


request_param1='{"bill_number":${test_03.data.orderId},"bill_number2":"aa"}'
data_depend_pt = r'\$\{.+?\..+?\}'

match1 = re.findall(data_depend_pt,request_param1)


print(match1)


