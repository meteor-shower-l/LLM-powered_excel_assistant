from flask import Flask

app = Flask(__name__)

@app.route('/mst')
def index():
    return '莫斯童666'

if __name__ == '__main__':
    app.run()


