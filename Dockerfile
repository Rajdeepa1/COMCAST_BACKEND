FROM python:3

ENV PYTHONBUFFERED=1 

WORKDIR /app

ADD . /app

COPY ./requirements.txt /app/requirements.txt

RUN chmod -R 777 *

RUN echo "export PATH=/app/chromedriver:${PATH}" >> /root/.bashrc

RUN pip install -r requirements.txt

COPY . /app

