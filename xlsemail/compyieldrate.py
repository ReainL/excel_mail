#!/usr/bin/env python3.4
# encoding: utf-8
"""
Created on 18-7-16
@title: '分公司综合评价'
@author: Xusl
"""
import datetime
import os
import logging
import logging.config
import pandas as pd

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import parseaddr, formataddr


EMAIL_HOST = 'mail.chyjr.com'
EMAIL_PORT = 465
# EMAIL_HOST_USER = ''
# EMAIL_HOST_PASSWORD = ''


def get_log_config():
    _config = {
        'version': 1,
        'formatters': {
            'generic': {
                'format': '%(asctime)s %(levelname)-5.5s [%(name)s:%(lineno)s][%(threadName)s] %(message)s',
            },
            'simple': {
                'format': '%(asctime)s %(levelname)-5.5s %(message)s',
            },
        },
        'handlers': {
            'console': {
                'class': 'logging.StreamHandler',
                'formatter': 'generic',
            },
            'file': {
                'class': 'logging.FileHandler',
                'filename': "payrollsend.log",
                'encoding': 'utf-8',
                'formatter': 'generic',

            },
        },
        'root': {
            'level': "INFO",
            'handlers': ['console', 'file', ],
        }
    }
    return _config


log_config = get_log_config()
logging.config.dictConfig(log_config)
logger = logging.getLogger(__file__)


def _format_float(obj):
    if isinstance(obj, float):
        return '{:.2f}'.format(obj)
    else:
        return str(obj)


def _format_percent(obj):
    if isinstance(obj, float) or isinstance(obj, int):
        t = obj*100
        return '{:.2f}'.format(t) + '%'
    else:
        return str(obj)


def _format_int(obj):
    if isinstance(obj, float):
        return '{:.0f}'.format(obj)
    else:
        return str(obj)


def _format(obj):
    if obj is None:
        return ''
    if isinstance(obj, datetime.datetime):
        # return obj.strftime('%Y-%m-%d %H:%M:%S')
        return obj.strftime('%Y-%m-%d')
    elif isinstance(obj, datetime.date):
        return obj.strftime('%Y-%m-%d')
    elif isinstance(obj, str):
        return str(obj).strip()
    else:
        # print(type(obj))
        # print(obj)
        return obj


def get_datetime_str(d_date=None, pattern='%Y-%m-%d'):
    """
    获取指定日期 字符格式
    :param d_date:
    :param pattern:
    :return:
    """
    if not d_date:
        d_date = datetime.datetime.now()
    return datetime.datetime.strftime(d_date, pattern)


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))


def _get_html_1(mail_begin, mail_end, data):
    """

    :param mail_begin: 邮件开头
    :param mail_end: 邮件结尾
    :param data:
    :return:
    """
    html_message = '''
        <meta>分公司各位领导好：</meta>
            <p>%s</p>
              %s
            <p></p>
            <meta>%s</meta>
            <p>望知悉！</p>
            <p>谢谢</p>
            ''' % (mail_begin, data, mail_end)
    return html_message


def send_stub(full_name, desc_full):
    """
    发送时间
    :param full_name:存储路径加名称
    :param desc_full:备份路径
    :return:
    """
    logging.info("邮件发送......")
    server = None
    df_desc = None
    tp = None
    try:
        df = pd.read_excel(full_name, sheetname='正文')
        df_content = pd.read_excel(full_name, sheetname='配置')
        subject = df_content.columns[1]
        mail_begin = df_content.iloc[:, 1].values[0]
        mail_end = df_content.iloc[:, 1].values[1]
        email_host_user = df_content.iloc[:, 1].values[2].replace(' ', '')
        email_host_password = df_content.iloc[:, 1].values[3].replace(' ', '')
        server = smtplib.SMTP_SSL('mail.chyjr.com', 465)
        server.ehlo()
        server.login(email_host_user, email_host_password)
        df_desc = df.copy()
        df_desc['result'] = ''
        # 定义删除列规则(删除后两列,即收件人、抄送人)
        del_col = df.columns[-2:]
        # 如果某列的值有%,则把数值转换为字符型
        for c in df.columns:
            if "%" in c:
                df[c] = df[c].apply(_format_percent)
        for index, row in df.iterrows():
            try:
                tp_1 = pd.DataFrame(columns=df.columns)
                # 加入内容
                tp = tp_1.append(df.iloc[index])
                msg = MIMEMultipart()
                # 取收件人列的值
                to_li = tp.iloc[:, -2].values[0]
                # 对收件人邮箱地址进行校验,替换不规则字符
                to_list = to_li.strip().replace('，', ',').replace('；', ',').replace(';', ',').replace(' ', '')
                to_lists = []
                for i_to in to_list.split(','):
                    # 对特殊字符检验
                    i_to = i_to.strip()
                    if i_to:
                        to_lists.append(_format_addr('<%s>' % i_to))
                msg['To'] = ','.join(to_lists)
                # 取抄送人列的值
                to_c = tp.iloc[:, -1].values[0]
                # 对抄送人邮箱地址进行校验,替换不规则字符
                to_cc = to_c.strip().replace('，', ',').replace('；', ',').replace(';', ',').replace(' ', '')
                to_ccs = []
                for i_cc in to_cc.split(','):
                    i_cc = i_cc.strip()
                    if i_cc:
                        to_ccs.append(_format_addr('<%s>' % i_cc))
                msg['Cc'] = ','.join(to_ccs)
                msg['Subject'] = Header(subject, 'utf-8').encode()
                # 删除发送人和抄送人
                # tp_new = tp.drop(del_col, axis=1, inplace=False).copy()
                tp.drop(del_col, axis=1, inplace=True)
                tp_html_start = tp.to_html(index=False)  # 将excle内容转换为html
                tp_html_new = tp_html_start.replace('text-align: right;"',
                                                    'text-align: center; color: #FFFFFF; background-color: #B22222"')
                tp_html = tp_html_new.replace('<tr', '<tr align="center"')
                html_message = _get_html_1(mail_begin, mail_end, tp_html)
                msg_text = MIMEText(html_message, 'html', 'utf-8')
                msg.attach(msg_text)
                server.sendmail(email_host_user, to_lists + to_ccs, msg.as_string())
                df_desc.iloc[index, -1] = '发送成功'
                logger.info('%s发送成功 ' % tp)
            except Exception as e:
                df_desc.iloc[index, -1] = '发送失败' + str(e)
                logger.error('%s发送失败 ' % tp)
    finally:
        df_desc.to_excel(desc_full, index=False)
        try:
            if server:
                server.quit()
        except Exception as e:
            logging.info(str(e))


def main():
    """
    入口
    :return:
    """
    func_name = "邮件发送"
    logger.info("start %s" % func_name)
    d_now = datetime.datetime.now()
    s_now = d_now.strftime("%Y%m%d%H%M%S")
    # 当前路径
    pwd_path = os.getcwd()
    # 文件存放路径
    cur_path = os.path.join(pwd_path, 'src')
    cur_dirs = os.listdir(cur_path)
    file_name = ''
    for cur_dir in cur_dirs:
        if cur_dir.find('.xlsx') != -1:
            file_name = cur_dir
    if not file_name:
        logging.info("没有Excel文件!")
        return
    logging.info(file_name)
    full_name = os.path.join(cur_path, file_name)
    # 将文件拷贝到desc目录
    desc_full = os.path.join(os.path.join(pwd_path, 'desc'), s_now + file_name)

    send_stub(full_name, desc_full)

    # 删除src目录内容
    os.remove(full_name)
    logger.info('... end %s' % func_name)


if __name__ == '__main__':
    main()
