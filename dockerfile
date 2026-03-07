FROM python:3.11-slim

WORKDIR /app

COPY . /app

ENV TZ=Asia/Taipei
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo  $TZ > /etc/tiomezone

COPY requirements.txt /app/
RUN python3 -m pip install -r requirements.txt
CMD ["python3", "plc2google.py"]