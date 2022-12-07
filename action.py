import json
import datetime
import pandas as pd
import xlwt


def read_json(file_path='/root/elastic/backup'):
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f.readlines():
            json_data = json.loads(line)
        return json_data


def edit_json(json_data):
    time_data = json_data['_source']['RequestReceivedTimestamp']
    time_data = datetime.datetime.strptime(time_data, "yyyy-MM-dd'T'HH:mm:ss+08:00")
    verb_data = json_data['_source']['Verb']
    code_data = json_data['_source']['ResponseStatus']['code']
    work_data = json_data['_source']['Workspace']
    prject_data = json_data['_source']['ObjectRef']["Namespace"]
    message_dta = json_data['_source']['Message']
    resource_data = json_data['_source']['ObjectRef']["Resource"]
    subresource_data = json_data['_source']['ObjectRef']["Subresource"]
    username_data = json_data['_source']["User"]["Username"]
    ip_data = json_data['_source']['SourceIPs']



    if code_data == 100:
        code_data = 'Continue'
    elif code_data == 101:
        code_data = 'Switching Protocols'
    elif code_data == 200:
        code_data = 'OK'
    elif code_data == 201:
        code_data = 'Created'
    elif code_data == 202:
        code_data = 'Accepted'
    elif code_data == 203:
        code_data = 'Non-Authoritative Information'
    elif code_data == 204:
        code_data = 'No Content'
    elif code_data == 205:
        code_data = 'Reset Content'
    elif code_data == 206:
        code_data = 'Partial Content'
    elif code_data == 300:
        code_data = 'Multiple Choices'
    elif code_data == 301:
        code_data = 'Moved Permanently'
    elif code_data == 302:
        code_data = 'Found'
    elif code_data == 303:
        code_data = 'See Other'
    elif code_data == 304:
        code_data = 'Not Modified'
    elif code_data == 305:
        code_data = 'Use Proxy'
    elif code_data == 307:
        code_data = 'Temporary Redirect'
    elif code_data == 400:
        code_data = 'Bad Request'
    elif code_data == 401:
        code_data = 'Unauthorized'
    elif code_data == 403:
        code_data = 'Forbidden'
    elif code_data == 404:
        code_data = 'Not Found'
    elif code_data == 405:
        code_data = 'Method Not Allowed'
    elif code_data == 406:
        code_data = 'Not Acceptable'
    elif code_data == 407:
        code_data = 'Proxy Authentication Required'
    elif code_data == 408:
        code_data = 'Request Time-out'
    elif code_data == 409:
        code_data = 'Conflict'
    elif code_data == 410:
        code_data = 'Gone'
    elif code_data == 411:
        code_data = 'Length Required'
    elif code_data == 412:
        code_data = 'Precondition Failed'
    elif code_data == 413:
        code_data = 'Request Entity Too Large'
    elif code_data == 414:
        code_data = 'Request-URI Too Large'
    elif code_data == 415:
        code_data = 'Unsupported Media Type'
    elif code_data == 416:
        code_data = 'Requested range not satisfiable'
    elif code_data == 417:
        code_data = 'Expectation Failed'
    elif code_data == 500:
        code_data = 'Internal Server Error'
    elif code_data == 501:
        code_data = 'Not Implemented'
    elif code_data == 502:
        code_data = 'Bad Gateway'
    elif code_data == 503:
        code_data = 'Service Unavailable'
    elif code_data == 504:
        code_data = 'Gateway Time-out'
    elif code_data == 505:
        code_data = 'HTTP Version not supported'


    excel_dic = {}
    excel_dic["时间"] = time_data
    excel_dic["操作行为"] = verb_data
    excel_dic["状态码"] = code_data
    excel_dic["企业空间"] = work_data
    excel_dic["项目"] = prject_data
    excel_dic["原因"] = message_dta
    excel_dic["资源类型与名称"] = resource_data
    excel_dic["子资源"] = subresource_data
    excel_dic["操作者"] = username_data
    excel_dic["源IP地址"] = ip_data
    return excel_dic

def to_excel(excel_dic):
    pf = pd.DataFrame(list(excel_dic))
    exce_time = datetime.datetime.now()
    filename = f'vxTEL_audit_<{exce_time}>' + '.xlx'
    file_path = pd.ExcelWriter(filename)
    pf.fillna(' ', inplace=True)
    pf.to_excel(file_path, encoding='utf-8', index=False)
    file_path.save()




