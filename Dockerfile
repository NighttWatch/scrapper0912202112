# Base Image - First Stage
FROM python:3.8

COPY . /app

WORKDIR /app

#Run requrements to user field (/root/.local)
RUN pip3 install -r requirements.txt

#run
CMD [ "python","-u" ,"./scrapper.py" ]