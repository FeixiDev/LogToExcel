import json
import pandas as pd
import openpyxl
from datetime import datetime, timedelta


def read_json(file_path):
    time_list = []
    verb_list = []
    code_list = []
    work_list = []
    # prject_list = []
    # message_list = []
    name_list = []
    resource_list = []
    # subresource_list = []
    username_list = []
    ip_data_list = []
    reason_list = []
    status_list = []
    code_dict = {100: 'Continue', 101: 'Switching Protocols', 200: 'Success',
                 201: 'Created', 202: 'Accepted', 203: 'Non-Authoritative Information', 204: 'No Content',
                 205: 'Reset Content', 206: 'Partial Content',
                 300: 'Multiple Choices', 301: 'Moved Permanently', 302: 'Found', 303: 'See Other', 304: 'Not Modified',
                 305: 'Use Proxy', 307: 'Temporary Redirect', 400: 'Bad Request',
                 401: 'Unauthorized', 403: 'Forbidden', 404: 'Not Found', 405: 'Method Not Allowed',
                 406: 'Not Acceptable',
                 407: 'Proxy Authentication Required',
                 408: 'Request Time-out', 409: 'Conflict', 410: 'Gone', 411: 'Length Required',
                 412: 'Precondition Failed',
                 413: 'Request Entity Too Large', 414: 'Request-URI Too Large', 415: 'Unsupported Media Type',
                 416: 'Requested range not satisfiable', 417: 'Expectation Failed', 500: 'Internal Server Error',
                 501: 'Not Implemented', 502: 'Bad Gateway',
                 503: 'Service Unavailable', 504: 'Gateway Time-out', 505: 'HTTP Version not supported', '': ''}
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f.readlines():
            json_data = json.loads(line)
            time_list.append(json_data['_source']['RequestReceivedTimestamp'])
            verb_list.append(json_data['_source']['Verb'])
            work_list.append(json_data['_source']['Workspace'])
            # prject_list.append(json_data['_source']['ObjectRef']["Namespace"])
            # message_list.append(json_data['_source']['Message'])
            resource_list.append(json_data['_source']['ObjectRef']["Resource"])
            name_list.append(json_data['_source']['ObjectRef']["Name"])
            # subresource_list.append(json_data['_source']['ObjectRef']["Subresource"])
            username_list.append(json_data['_source']["User"]["Username"])
            ip_data_list.append(json_data['_source']['SourceIPs'])
            code_list.append(code_dict[json_data.get('_source', {}).get('ResponseStatus', ()).get('code', '')])
            name_resource_list = [str(a) + " " + b for a, b in zip(resource_list, name_list)]
            code_verb_list = [str(a) + " " + b for a, b in zip(verb_list, code_list)]
            if json_data.get('_source', {}).get('ResponseStatus', ()).get('reason', ''):
                reason_list.append(json_data['_source']['ResponseStatus']['reason'])
            else:
                reason_list.append(code_verb_list[0])
            if json_data.get('_source', {}).get('ResponseStatus', ()).get('status', ''):
                status_list.append(json_data['_source']['ResponseStatus']['status'])
            else:
                status_list.append('')
        for i in range(len(time_list)):
            time_list[i] = time_list[i][:19]
            time_date = datetime.strptime(time_list[i], "%Y-%m-%dT%H:%M:%S")
            end_time = time_date.strftime("%Y-%m-%d %H:%M:%S")
        # for i in range(len(code_list)):
        #     if code_list[i] == 100:
        #         code_list[i] = 'Continue'
        #     elif code_list[i] == 101:
        #         code_list[i] = 'Switching Protocols'
        #     elif code_list[i] == 200:
        #         code_list[i] = 'Success'
        #     elif code_list[i] == 201:
        #         code_list[i] = 'Created'
        #     elif code_list[i] == 202:
        #         code_list[i] = 'Accepted'
        #     elif code_list[i] == 203:
        #         code_list[i] = 'Non-Authoritative Information'
        #     elif code_list[i] == 204:
        #         code_list[i] = 'No Content'
        #     elif code_list[i] == 205:
        #         code_list[i] = 'Reset Content'
        #     elif code_list[i] == 206:
        #         code_list[i] = 'Partial Content'
        #     elif code_list[i] == 300:
        #         code_list[i] = 'Multiple Choices'
        #     elif code_list[i] == 301:
        #         code_list[i] = 'Moved Permanently'
        #     elif code_list[i] == 302:
        #         code_list[i] = 'Found'
        #     elif code_list[i] == 303:
        #         code_list[i] = 'See Other'
        #     elif code_list[i] == 304:
        #         code_list[i] = 'Not Modified'
        #     elif code_list[i] == 305:
        #         code_list[i] = 'Use Proxy'
        #     elif code_list[i] == 307:
        #         code_list[i] = 'Temporary Redirect'
        #     elif code_list[i] == 400:
        #         code_list[i] = 'Bad Request'
        #     elif code_list[i] == 401:
        #         code_list[i] = 'Unauthorized'
        #     elif code_list[i] == 403:
        #         code_list[i] = 'Forbidden'
        #     elif code_list[i] == 404:
        #         code_list[i] = 'Not Found'
        #     elif code_list[i] == 405:
        #         code_list[i] = 'Method Not Allowed'
        #     elif code_list[i] == 406:
        #         code_list[i] = 'Not Acceptable'
        #     elif code_list[i] == 407:
        #         code_list[i] = 'Proxy Authentication Required'
        #     elif code_list[i] == 408:
        #         code_list[i] = 'Request Time-out'
        #     elif code_list[i] == 409:
        #         code_list[i] = 'Conflict'
        #     elif code_list[i] == 410:
        #         code_list[i] = 'Gone'
        #     elif code_list[i] == 411:
        #         code_list[i] = 'Length Required'
        #     elif code_list[i] == 412:
        #         code_list[i] = 'Precondition Failed'
        #     elif code_list[i] == 413:
        #         code_list[i] = 'Request Entity Too Large'
        #     elif code_list[i] == 414:
        #         code_list[i] = 'Request-URI Too Large'
        #     elif code_list[i] == 415:
        #         code_list[i] = 'Unsupported Media Type'
        #     elif code_list[i] == 416:
        #         code_list[i] = 'Requested range not satisfiable'
        #     elif code_list[i] == 417:
        #         code_list[i] = 'Expectation Failed'
        #     elif code_list[i] == 500:
        #         code_list[i] = 'Internal Server Error'
        #     elif code_list[i] == 501:
        #         code_list[i] = 'Not Implemented'
        #     elif code_list[i] == 502:
        #         code_list[i] = 'Bad Gateway'
        #     elif code_list[i] == 503:
        #         code_list[i] = 'Service Unavailable'
        #     elif code_list[i] == 504:
        #         code_list[i] = 'Gateway Time-out'
        #     elif code_list[i] == 505:
        #         code_list[i] = 'HTTP Version not supported'
        #     else:
        #         code_list[i] = ''

        excel_dic = {}
        excel_dic['时间'] = end_time
        excel_dic["操作者"] = username_list
        # excel_dic["操作行为"] = verb_list
        # excel_dic['状态码'] = code_list
        excel_dic["主机名"] = work_list
        excel_dic["审计对象"] = name_resource_list
        excel_dic["事件说明"] = reason_list
        # excel_dic["项目"] = prject_list
        # excel_dic["原因"] = message_list
        # excel_dic["子资源"] = subresource_list
        excel_dic["事件级别"] = status_list
        excel_dic["访问发起端"] = ip_data_list
        return excel_dic


def creat_excel(excel_path):
    excel_time = f'{datetime.now().year}' + '-' + \
                 f'{datetime.now().month}' + '-' + \
                 f'{datetime.now().day}'
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
