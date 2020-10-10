from flask import Flask

from apscheduler.schedulers.background import BackgroundScheduler


app = Flask(__name__)

def test_job(name):
    try:
        print('I am working...'+name)
        a + 1
    except:
        print("Errr")

scheduler = BackgroundScheduler()



@app.route('/get_state', methods=["GET"])
def get_state():
    job = scheduler.add_job(test_job, 'interval', minutes=1, args=['sadf'])
    scheduler.start()
    return f"state"

if __name__ == '__main__':
    app.run()