import win32api
import subprocess
import socket
import base64
import os
import pyaudio
import numpy as np
from win32con import VK_MEDIA_PLAY_PAUSE, VK_MEDIA_NEXT_TRACK, VK_MEDIA_PREV_TRACK, KEYEVENTF_EXTENDEDKEY
from flask import Flask, json, request, jsonify
app = Flask(__name__)

device_info = {}


### Инициализация ###
# Выбор устройства воспроизведения #

p = pyaudio.PyAudio()
try:
    default_device = p.get_default_output_device_info()
except IOError:
    default_device = -1
default_name = default_device["name"]

while True:
    print ("Доступные WASAPI аудиоинтерфейсы:\n")
    for i in range(0, p.get_device_count()):
        info = p.get_device_info_by_index(i)
        if("WASAPI" in p.get_host_api_info_by_index(info["hostApi"])["name"] and info["maxInputChannels"] == 0):
            print(str(info["index"]) + ":" + info["name"])
            if(default_name in info["name"]):
                default_device_index = i
    print(str(default_device_index) + ":" + default_name)
    device_id = int(input("Выберите аудиоинтерфейс [" + str(default_device_index) + ":" + str(default_name) + "]: ") or default_device_index)
    if(device_id < p.get_device_count()):
        device_info = p.get_device_info_by_index(device_id)
        print(device_info)
        if("WASAPI" in p.get_host_api_info_by_index(device_info["hostApi"])["name"] and device_info["maxInputChannels"] == 0):
            break
        else:
            print("\n\nАудиоинтерфейс не поддерживает WASAPI. Попробуйте выбрать ещё раз.")
    else:
        print("\n\nПопробуйте выбрать ещё раз.")





### Веб часть ###

@app.route('/')
def index():
    # Шаблон страницы #
    html = """
    <!DOCTYPE html>
    <html class="player-ui">
    <head>
    <meta charset="UTF-8">
    <title>Media Keys Remote Control</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/normalize/5.0.0/normalize.min.css">
    <link rel='stylesheet prefetch' href='https://fonts.googleapis.com/css?family=Roboto:300,400,500,700'>
    <link rel='stylesheet prefetch' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.6.3/css/font-awesome.min.css'>
    <style type="text/css">
        html {
        font-size: 16px;
        }

        body {
        font-family: 'Roboto', Arial, Verdana, sans-serif;
        background: #e4f2fb;
        }

        a {
        text-decoration: none;
        }

        .player__container {
        margin-top: 2rem;
        margin-right: auto;
        margin-left: auto;
        max-width: 20rem;
        background: #fff;
        border-radius: 0.25rem;
        box-shadow: 0 10px 20px -5px rgba(0, 0, 0, 0.19), 0 6px 6px -10px rgba(0, 0, 0, 0.23);
        }

        .body__cover {
        position: relative;
        }

        .body__cover img {
        max-width: 100%;
        border-radius: 0.25rem;
        }

        .list {
        display: -webkit-box;
        display: -ms-flexbox;
        display: flex;
        margin: 0;
        padding: 0;
        list-style-type: none;
        }

        .body__buttons,
        .body__info,
        .player__footer {
        padding-right: 2rem;
        padding-left: 2rem;
        }

        .list--cover,
        .list--footer {
        -webkit-box-pack: justify;
            -ms-flex-pack: justify;
                justify-content: space-between;
        }

        .list--header .list__link,
        .list--footer .list__link {
        color: #888;
        }

        .list--cover {
        position: absolute;
        top: .5rem;
        width: 100%;
        }
        .list--cover li:first-of-type {
        margin-left: .75rem;
        }
        .list--cover li:last-of-type {
        margin-right: .75rem;
        }
        .list--cover a {
        font-size: 1.15rem;
        color: #fff;
        }

        .range {
        position: relative;
        top: -1.5rem;
        right: 0;
        left: 0;
        margin: auto;
        background: rgba(255, 255, 255, 0.95);
        width: 80%;
        height: 0.125rem;
        border-radius: 0.25rem;
        cursor: pointer;
        }
        .range:before, .range:after {
        content: "";
        position: absolute;
        cursor: pointer;
        }
        .range:before {
        width: 3rem;
        height: 100%;
        background: -webkit-linear-gradient(left, rgba(211, 3, 32, 0.5), rgba(211, 3, 32, 0.85));
        background: linear-gradient(to right, rgba(211, 3, 32, 0.5), rgba(211, 3, 32, 0.85));
        border-radius: 0.25rem;
        overflow: hidden;
        }
        .range:after {
        top: -0.375rem;
        left: 3rem;
        z-index: 3;
        width: 0.875rem;
        height: 0.875rem;
        background: #fff;
        border-radius: 50%;
        box-shadow: 0 0 3px rgba(0, 0, 0, 0.15), 0 2px 4px rgba(0, 0, 0, 0.15);
        -webkit-transition: all 0.25s cubic-bezier(0.4, 0, 1, 1);
        transition: all 0.25s cubic-bezier(0.4, 0, 1, 1);
        }
        .range:focus:after, .range:hover:after {
        background: rgba(211, 3, 32, 0.95);
        }

        .body__info {
        padding-top: 1.5rem;
        padding-bottom: 1.25rem;
        text-align: center;
        }

        .info__album,
        .info__song {
        margin-bottom: .5rem;
        }

        .info__artist,
        .info__album {
        font-size: .75rem;
        font-weight: 300;
        color: #666;
        }

        .info__song {
        font-size: 1.15rem;
        font-weight: 400;
        color: #d30320;
        }

        .body__buttons {
        padding-bottom: 2rem;
        }

        .body__buttons {
        padding-top: 1rem;
        }

        .list--buttons {
        -webkit-box-align: center;
            -ms-flex-align: center;
                align-items: center;
        -webkit-box-pack: center;
            -ms-flex-pack: center;
                justify-content: center;
        }

        .list--buttons li:nth-of-type(n+2) {
        margin-left: 1.25rem;
        }

        .list--buttons a {
        padding-top: .45rem;
        padding-right: .75rem;
        padding-bottom: .45rem;
        padding-left: .75rem;
        font-size: 1rem;
        border-radius: 50%;
        box-shadow: 0 3px 6px rgba(33, 33, 33, 0.1), 0 3px 12px rgba(33, 33, 33, 0.15);
        }
        .list--buttons a:focus, .list--buttons a:hover {
        color: rgba(171, 2, 26, 0.95);
        opacity: 1;
        box-shadow: 0 6px 9px rgba(33, 33, 33, 0.1), 0 6px 16px rgba(33, 33, 33, 0.15);
        }
        .list--buttons li:nth-of-type(2) a {
        padding-top: .82rem;
        padding-right: 1rem;
        padding-bottom: .82rem;
        padding-left: 1.19rem;
        margin-left: .5rem;
        font-size: 1.25rem;
        color: rgba(211, 3, 32, 0.95);
        }
        .list--buttons li:first-of-type a,
        .list--buttons li:last-of-type a {
        font-size: .95rem;
        color: #212121;
        opacity: .5;
        }
        .list--buttons li:first-of-type a:focus, .list--buttons li:first-of-type a:hover,
        .list--buttons li:last-of-type a:focus,
        .list--buttons li:last-of-type a:hover {
        color: #d30320;
        opacity: .75;
        }

        .list__link {
        -webkit-transition: all 0.25s cubic-bezier(0.4, 0, 1, 1);
        transition: all 0.25s cubic-bezier(0.4, 0, 1, 1);
        }
        .list__link:focus, .list__link:hover {
        color: #d30320;
        }

        .player__footer {
        padding-top: 1rem;
        padding-bottom: 2rem;
        }

        .list--footer a {
        opacity: .5;
        }
        .list--footer a:focus, .list--footer a:hover {
        opacity: .9;
        }
    </style>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>

    <script type=text/javascript>
            $(function() {
                $('a#prev').bind('click', function() {
                    $.getJSON('/prev',
                        function(data) {
                    });
                    return false;
                });
            });
            $(function() {
                $('a#next').bind('click', function() {
                    $.getJSON('/next',
                        function(data) {
                    });
                    return false;
                });
            });

            window.setInterval(function(){
                $.getJSON('/status',
                    function(data) {
                        if(data.audio == 'True'){
                            document.getElementById('playpause').className = "fa fa-pause"; 
                        } else {
                            document.getElementById('playpause').className = "fa fa-play";
                        }
                });
            }, 250);

            $(function() {
                $('a#play').bind('click', function() {
                    $.getJSON('/play',
                        function(data) {
                    });
                    return false;
                });
            });
            
    </script>
    </head>

    <body>
    <div class="wrapper">
    <div class="player__container">
        <div class="player__body">
        <div class="body__cover">
            <ul class="list list--cover">
            <li>
                <a class="list__link" href=""></a>
            </li>
            </ul>

            <img src="WindowsUserAvatarBase64" alt="User avatar" />
        </div>
        
        <div class="body__info">
            <div class="info__album">Playing on WindowsComputerName</div>

            <div class="info__song">WindowsUserName</div>

            <div class="info__artist">WindowsIPAdress</div>
        </div>
            <div class="body__buttons">
                <ul class="list list--buttons">
                <li><a href="#" id=prev><i class="fa fa-step-backward"></i></a></li>

                <li><a href="#" id=play><i class="fa fa-play" id="playpause"></i></a></li>
                
                <li><a href="#" id=next><i class="fa fa-step-forward"></i></a></li>
                </ul>
            </div>
        </div>
    </div>
    </div>
    
    
    </body>
    </html>
    """

    # Получение аватарки пользователя #
    p = subprocess.Popen("""wmic useraccount where name='%username%' get sid""", stdout=subprocess.PIPE, shell=True)
    (output, err) = p.communicate()
    p_status = p.wait()
    sid = 'S-' + str(output).split('S-')[1].split('  ')[0]
    files = os.listdir("C:/Users/Public/AccountPictures/" + sid)
    for i in files:
        if "448.jpg" in i:
            picture = "C:/Users/Public/AccountPictures/" + sid + "/" + i
            break
    with open(picture, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read())

    # Вставляем информацию в шаблон #
    html = html.replace("WindowsIPAdress", str(socket.gethostbyname(socket.gethostname())))
    html = html.replace("WindowsUserName", str(os.getlogin()))
    html = html.replace("WindowsComputerName", str(os.environ['COMPUTERNAME']))
    html = html.replace("WindowsUserAvatarBase64", "data:image/jpg;base64," + str(encoded_string).replace("b'","").replace("'",""))

    # Возвращаем клиенту ответ на запрос #
    return html


@app.route('/prev')
def prev():
    win32api.keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENDEDKEY, 0)
    return 'N'

@app.route('/next')
def next():
    win32api.keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENDEDKEY, 0)
    return 'N'

@app.route('/play')
def play():
    win32api.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENDEDKEY, 0)
    return 'N'

@app.route('/status')
def status():
    # Открываем loopback выбранного аудиоинтерфейса #
    stream = p.open(format = pyaudio.paInt16,
                channels = device_info["maxOutputChannels"],
                rate = int(device_info["defaultSampleRate"]),
                input = True,
                frames_per_buffer = 512,
                input_device_index = device_info["index"],
                as_loopback = True)

    # Возвращаем активность аудиоинтерфейса #
    if(stream and stream.is_active()):
        level = abs(np.fromstring(stream.read(2**16, exception_on_overflow = False),dtype=np.int16)[0])
        if(level > 1):
            return jsonify(audio='True')
        else:
            return jsonify(audio='False')
    else:
        return jsonify(audio='False')

# Запускаем flask #
if __name__ == "__main__":
    app.run(host='0.0.0.0')