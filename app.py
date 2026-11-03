"""本地浏览器界面；文件仅在本机内存中处理，无外部API调用。"""
import argparse
import json
import secrets
import threading
import webbrowser
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from generator import generate, InputError

ROOT=Path(__file__).parent
TOKEN=secrets.token_urlsafe(32)
LIMIT=25*1024*1024

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # 不记录上传文件名或内容

    def send(self,code,body,kind='application/json; charset=utf-8',download=False):
        self.send_response(code)
        self.send_header('Content-Type',kind)
        self.send_header('Content-Length',str(len(body)))
        self.send_header('Cache-Control','no-store')
        self.send_header('X-Content-Type-Options','nosniff')
        if download:self.send_header('Content-Disposition','attachment; filename="audit-report.docx"')
        self.end_headers();self.wfile.write(body)

    def do_GET(self):
        if self.path!='/':return self.send(404,b'{}')
        html=(ROOT/'web/index.html').read_text(encoding='utf-8').replace('__TOKEN__',TOKEN)
        self.send(200,html.encode(), 'text/html; charset=utf-8')

    def do_POST(self):
        if self.path not in ('/analyze','/generate'):return self.send(404,b'{}')
        if not secrets.compare_digest(self.headers.get('X-Request-Token',''),TOKEN):
            return self.send(403,b'{"error":"Invalid request token"}')
        try:
            size=int(self.headers.get('Content-Length','0'))
            if not 0<size<=LIMIT:raise InputError('上传文件合计不得超过25MB。')
            content_type=self.headers.get('Content-Type','')
            if not content_type.startswith('multipart/form-data;'):raise InputError('请使用文件上传表单。')
            message=BytesParser(policy=default).parsebytes(('Content-Type: '+content_type+'\r\nMIME-Version: 1.0\r\n\r\n').encode()+self.rfile.read(size))
            files={};fields={}
            for part in message.iter_parts():
                name=part.get_param('name',header='content-disposition');filename=part.get_filename()
                data=part.get_payload(decode=True) or b''
                if filename:
                    expected={'excel':'.xlsx','template':'.docx'}.get(name)
                    if not expected or Path(filename).suffix.lower()!=expected:raise InputError('请选择 .xlsx 缺陷表和 .docx 模板。')
                    files[name]=data
                else:fields[name]=data.decode('utf-8').strip()
            if not files.get('excel'):raise InputError('请先选择缺陷汇总表。')
            if not files.get('template'):raise InputError('请选择本地 Word 报告模板（.docx）。')
            data,info=generate(files['excel'],files.get('template'),fields.get('mode','chapters'),fields.get('company',''),fields.get('original_period',''),fields.get('rolling_period',''))
            if self.path=='/analyze':self.send(200,json.dumps(info,ensure_ascii=False).encode())
            else:self.send(200,data,'application/vnd.openxmlformats-officedocument.wordprocessingml.document',True)
        except (InputError,ValueError) as e:
            self.send(400,json.dumps({'error':str(e)},ensure_ascii=False).encode())
        except Exception:
            self.send(500,json.dumps({'error':'文件处理失败，请确认文件未损坏，并选择与原报告结构一致的模板。'},ensure_ascii=False).encode())

def main():
    p=argparse.ArgumentParser(description='审计整改报告生成器（本地运行）')
    p.add_argument('--port',type=int,default=8765);p.add_argument('--no-browser',action='store_true')
    a=p.parse_args();server=ThreadingHTTPServer(('127.0.0.1',a.port),Handler)
    print(f'程序已启动，请在本机浏览器打开 http://127.0.0.1:{a.port} ；按 Ctrl+C 退出。',flush=True)
    if not a.no_browser:threading.Timer(0.8,lambda:webbrowser.open(f'http://127.0.0.1:{a.port}')).start()
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()

if __name__=='__main__':main()
