import poplib
import email
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from bs4 import BeautifulSoup
import requests
from datetime import datetime
import os

poplib._MAXLINE=20480

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def guess_charset(msg):
    # 先从msg对象获取编码:
    charset = msg.get_charset()
    if charset is None:
        # 如果获取不到，再从Content-Type字段获取:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def get_email_headers(msg):
    # 邮件的From, To, Subject存在于根对象上:
    headers = {}
    for header in ['From', 'To', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['date'] = value
            if header == 'Subject':
                # 需要解码Subject字符串:
                subject = decode_str(value)
                headers['subject'] = subject
            else:
                # 需要解码Email地址:
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                value = u'%s <%s>' % (name, addr)
                if header == 'From':
                    from_address = value
                    headers['from'] = from_address
                else:
                    to_address = value
                    headers['to'] = to_address
    content_type = msg.get_content_type()
    # print('head content_type: ')
    # print(content_type)
    return headers

# indent用于缩进显示:
def get_email_cntent(message, base_save_path):
    j = 0
    content = ''
    attachment_files = []
    for part in message.walk():
        j = j + 1
        file_name = part.get_filename()
        contentType = part.get_content_type()
        # 保存附件
        if file_name: # Attachment
            # Decode filename
            h = email.header.Header(file_name)
            dh = email.header.decode_header(h)
            filename = dh[0][0]
            # print(dh)
            # print(type(filename))
            if dh[0][1]: # 如果包含编码的格式，则按照该格式解码
                # filename = unicode(filename, dh[0][1])
                # filename = filename.encode("utf-8")
                filename = filename.decode(encoding=dh[0][1])
            data = part.get_payload(decode=True)
            att_file = open(base_save_path + filename, 'wb')
            attachment_files.append(filename)
            att_file.write(data)
            att_file.close()
        elif contentType == 'text/plain' or contentType == 'text/html':
            # 保存正文
            data = part.get_payload(decode=True)
            charset = guess_charset(part)
            if charset:
                charset = charset.strip().split(';')[0]
                # print('charset:')
                # print(charset)
                data = data.decode(charset)
            content = data
    return content, attachment_files

def get_page(url):
    # print(url)
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.146 Safari/537.36'
    }
    res = requests.get(url, headers=headers, timeout=5)
    # print(res.text)
    data = res.text.encode('latin1',errors='ignore').decode('utf-8', errors='ignore')
    # soup = BeautifulSoup(data, 'lxml', fromEncoding="gbk")
    soup = BeautifulSoup(data, 'lxml')
    return soup

if __name__ == '__main__':

    #---------------参数------------------
    email_address = '190xxxxxxx@pku.edu.cn'
    email_password = 'xxxxxxxx'
    pop3_server = 'mail.pku.edu.cn'
    search_count = 50
    base_save_path = './pku_mail/'
    # keyword = '检索系统平台'
    # target_date = '2020-03-20'
    keyword = input('请输入关键字（默认：检索系统平台）：')
    target_date = input('请输入查询日期（格式：2020-MM-dd, 默认：今天）：')
    if keyword == '':
        keyword = '检索系统平台'
    if target_date == '':
        now = datetime.now()
        current_date = str(now.strftime('%Y-%m-%d'))
        target_date = current_date
    print(keyword)
    print(target_date)    
    #---------------------------------------

    if not os.path.exists(base_save_path):
        os.mkdir(base_save_path)
    server = poplib.POP3(pop3_server)
    # print(server.getwelcome())
    server.user(email_address)
    server.pass_(email_password)

    # messageCount, messageSize = server.stat()

    # print('messageCount:')
    # print(messageCount)
    # print('messageSize:')
    # print(messageSize)
    
    resp, mails, octets = server.list()
    # print('----------resp----------')
    # print(resp)
    # print('----------mails----------')
    # # print(mails)
    # print('----------octets----------')
    # print(octets)

    length = len(mails)
    print('总共有%d封邮件，搜索前%d封' %(length, search_count))
    
    # length = 10
    mail_count = 0
    downloaded = 0
    for i in range(length -1, length - search_count - 1,  - 1):
        resp, lines, octets = server.retr(i + 1)
        # print(lines)
        msg_content = b'\n'.join(lines)
        msg = Parser().parsestr(msg_content.decode("gb2312", 'ignore'))
        # print(msg)

        print('----------正在检索----------')
        
        msg_headers = get_email_headers(msg)
        # content, attachment_files = get_email_cntent(msg, base_save_path)
        # print('from_address:')
        # print(msg_headers)
        msg_date_time_GMT = msg_headers['date']
        # print(msg_date_time_GMT)
        GMT_FORMAT = '%a, %d %b %Y %H:%M:%S +0800 (CST)'
        msg_datetime = datetime.strptime(msg_date_time_GMT, GMT_FORMAT)
        msg_date = msg_datetime.date()
        # print(str(msg_date))
        
        # print('subject:')
        # print(type(msg_headers))
        msg_subject = msg_headers['subject']
        print(msg_subject)
        filename = msg_subject + '.pdf'
        if keyword in msg_subject and target_date == str(msg_date):
        # if False:
            print("找到一封符合条件邮件")
            mail_count = mail_count + 1
            content, attachment_files = get_email_cntent(msg, base_save_path)
            soup = BeautifulSoup(content, 'lxml')
            link = soup.find('a').get('href')
            file_page = get_page(link)
            tc_left = file_page.find('div', {'id':'tc_left'})
            tc_left_a = tc_left.find_all('a')
            # print('tc_left_a')
            # print(tc_left_a)
            # print('href')
            href = tc_left_a[-1].get('href')
            # print(href)
            if not href == '':
                print('获取PDF链接成功，正在下载...')
            else:
                print('获取PDF链接失败，跳过改文件')
            r = requests.get(href, timeout = 120)
            # print("len(r): " + str(len(r)))
            with open(base_save_path + filename, 'wb+') as f:
                f.write(r.content)
                f.close()
                print('文件下载成功')
                downloaded = downloaded + 1
            # break
        print('----------检索完毕----------\n')
    print('搜索完毕，共找到%d封符合条件邮件，成功下载了%d个PDF文件' %(mail_count, downloaded))
    input('按任意键退出程序：')
    server.quit()



