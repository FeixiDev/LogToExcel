import json
import datetime
import pandas as pd
import openpyxl


def read_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        time_list = []
        verb_list = []
        code_list = []
        work_list = []
        prject_list = []
        message_list = []
        resource_list = []
        subresource_list = []
        username_list = []
        ip_data_list = []
        for line in f.readlines():
            json_data = json.loads(line)
            time_list.append(json_data['_source']['RequestReceivedTimestamp'])
            verb_list.append(json_data['_source']['Verb'])
            work_list.append(json_data['_source']['Workspace'])
            prject_list.append(json_data['_source']['ObjectRef']["Namespace"])
            message_list.append(json_data['_source']['Message'])
            resource_list.append(json_data['_source']['ObjectRef']["Resource"])
            subresource_list.append(json_data['_source']['ObjectRef']["Subresource"])
            username_list.append(json_data['_source']["User"]["Username"])
            ip_data_list.append(json_data['_source']['SourceIPs'])
            code_list.append(json_data.get('_source', {}).get('ResponseStatus', ()).get('code', ''))
        for i in range(len(time_list)):
            time_list[i] = time_list[i][:19]
        for i in range(len(code_list)):
            if code_list[i] == 100:
                code_list[i] = 'Continue'
            elif code_list[i] == 101:
                code_list[i] = 'Switching Protocols'
            elif code_list[i] == 200:
                code_list[i] = 'OK'
            elif code_list[i] == 201:
                code_list[i] = 'Created'
            elif code_list[i] == 202:
                code_list[i] = 'Accepted'
            elif code_list[i] == 203:
                code_list[i] = 'Non-Authoritative Information'
            elif code_list[i] == 204:
                code_list[i] = 'No Content'
            elif code_list[i] == 205:
                code_list[i] = 'Reset Content'
            elif code_list[i] == 206:
                code_list[i] = 'Partial Content'
            elif code_list[i] == 300:
                code_list[i] = 'Multiple Choices'
            elif code_list[i] == 301:
                code_list[i] = 'Moved Permanently'
            elif code_list[i] == 302:
                code_list[i] = 'Found'
            elif code_list[i] == 303:
                code_list[i] = 'See Other'
            elif code_list[i] == 304:
                code_list[i] = 'Not Modified'
            elif code_list[i] == 305:
                code_list[i] = 'Use Proxy'
            elif code_list[i] == 307:
                code_list[i] = 'Temporary Redirect'
            elif code_list[i] == 400:
                code_list[i] = 'Bad Request'
            elif code_list[i] == 401:
                code_list[i] = 'Unauthorized'
            elif code_list[i] == 403:
                code_list[i] = 'Forbidden'
            elif code_list[i] == 404:
                code_list[i] = 'Not Found'
            elif code_list[i] == 405:
                code_list[i] = 'Method Not Allowed'
            elif code_list[i] == 406:
                code_list[i] = 'Not Acceptable'
            elif code_list[i] == 407:
                code_list[i] = 'Proxy Authentication Required'
            elif code_list[i] == 408:
                code_list[i] = 'Request Time-out'
            elif code_list[i] == 409:
                code_list[i] = 'Conflict'
            elif code_list[i] == 410:
                code_list[i] = 'Gone'
            elif code_list[i] == 411:
                code_list[i] = 'Length Required'
            elif code_list[i] == 412:
                code_list[i] = 'Precondition Failed'
            elif code_list[i] == 413:
                code_list[i] = 'Request Entity Too Large'
            elif code_list[i] == 414:
                code_list[i] = 'Request-URI Too Large'
            elif code_list[i] == 415:
                code_list[i] = 'Unsupported Media Type'
            elif code_list[i] == 416:
                code_list[i] = 'Requested range not satisfiable'
            elif code_list[i] == 417:
                code_list[i] = 'Expectation Failed'
            elif code_list[i] == 500:
                code_list[i] = 'Internal Server Error'
            elif code_list[i] == 501:
                code_list[i] = 'Not Implemented'
            elif code_list[i] == 502:
                code_list[i] = 'Bad Gateway'
            elif code_list[i] == 503:
                code_list[i] = 'Service Unavailable'
            elif code_list[i] == 504:
                code_list[i] = 'Gateway Time-out'
            elif code_list[i] == 505:
                code_list[i] = 'HTTP Version not supported'
            else:
                code_list[i] = ''

        excel_dic = {}
        excel_dic['时间'] = time_list
        excel_dic["操作行为"] = verb_list
        excel_dic['状态码'] = code_list
        excel_dic["企业空间"] = work_list
        excel_dic["项目"] = prject_list
        excel_dic["原因"] = message_list
        excel_dic["资源类型与名称"] = resource_list
        excel_dic["子资源"] = subresource_list
        excel_dic["操作者"] = username_list
        excel_dic["源IP地址"] = ip_data_list
        return excel_dic


def creat_excel(excel_path):
    excel_time = f'{datetime.datetime.now().year}' + '-' + \
                 f'{datetime.datetime.now().month}' + '-' + \
                 f'{datetime.datetime.now().day}'
    excel_name = excel_path + f'/vxTEL_audit_{excel_time}.xlsx'
    wb = openpyxl.Workbook()
    wb.save(excel_name)
    return excel_name


def to_excel(excel_dic, excel_name, sheet_name):
    excel_pd = pd.DataFrame(excel_dic)
    with pd.ExcelWriter(excel_name, mode='a', engine='openpyxl') as writer_excel:
        excel_pd.to_excel(writer_excel, sheet_name=sheet_name, index=False)
    rm = openpyxl.load_workbook(excel_name)
    # sheet = rm['Sheet']
    # rm.remove(sheet)
    # rm.save(excel_name)
