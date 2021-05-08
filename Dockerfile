FROM python:3.9-alpine

RUN apk --no-cache update

RUN apk add --no-cache g++ gcc libxslt-dev

RUN apk add --no-cache py3-pip

RUN apk add --no-cache libreoffice

WORKDIR /main

COPY ./requirements.txt ./

RUN pip3 install -r requirements.txt --no-cache-dir

COPY ./ ./

CMD python3 main.py
