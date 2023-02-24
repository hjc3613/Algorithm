import tornado.web
import tornado.ioloop
import tornado.httpserver
import tornado.options
import requests
requests.post()

from tornado.options import define, options
from tornado.web import RequestHandler, url

class IndexHandler(tornado.web.RequestHandler):
    def get(self):
        self.write('hello itcast')

class SubjectHandler(RequestHandler):
    def initialize(self, subject):
        self.subject = subject

    def get(self, ):
        self.get_argument()
        self.write(self.subject)

if __name__ == '__main__':
    app = tornado.web.Application(
        [
            (r'/', IndexHandler),
            (r'/python', SubjectHandler, {'subject':'python'}),
            url(r'/cpp', SubjectHandler, {'subject':'cpp'}, name='cpp_url')
        ]
    )
    app.listen(8000)
    tornado.ioloop.IOLoop.current().start()