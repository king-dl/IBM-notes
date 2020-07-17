# encoding: gb2312
# notes 9.0
'''
@author: k
@contact: 524180147@qq.com
@file: *********
@time: 2019/12/25 16:43
@desc:
'''

from win32com.client import DispatchEx
from win32com.client import makepy

makepy.GenerateFromTypeLibSpec('Lotus Domino Objects')


class NotesMail(object):
    def __init__(self, server, file):
        print('init mail client')
        self.session = DispatchEx('Notes.NotesSession')
        self.db = self.session.GetDatabase(server, file)
        if not self.db.IsOpen:
            print('open mail db')
            try:
                self.db.OPENMAIL
            except Exception as e:
                print(str(e))
                print('could not open database: {}'.format(db_name))

    def send_mail(self,sendto,copyto, blindcopyto,subject, body_text, attach):
        doc = self.db.CREATEDOCUMENT
        doc.sendto = sendto
        if copyto is not None:
            doc.copyto = copyto
        if blindcopyto is not None:
            doc.blindcopyto = blindcopyto
        doc.Subject = subject
        # body
        body = doc.CreateRichTextItem("Body")
        body.AppendText(body_text)

        # attachment
        if attach is not None:
            attachment = doc.CreateRichTextItem("Attachment")
            for att in attach:
                print(att)
                attachment.EmbedObject(1454, "", att, "Attachment")
        doc.SaveMessageOnSend = True
        doc.Send(False)
        print('send success')


def main():
    #"""notes服务器信息：主机 邮件文件"""
    mail = NotesMail('*', '*.nsf')
    #"""发送"""
    sendto = ['k']
    #"""抄送"""
    copyto= ''
    #"""密送"""
    blindcopyto=''
    #"""主题"""
    subject = 'TEST'
    #"""邮件正文"""
    body_text = 'TEST'
    #"""附件"""
    attach = [r'D:\DevOps\send_mail.py']

    mail.send_mail(sendto,copyto, blindcopyto,subject,body_text,attach)

if __name__ == '__main__':
    main()