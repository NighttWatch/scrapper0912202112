# Base Image - First Stage
FROM python:3.9

RUN apt-get update 
RUN apt-get install -y gconf-service libasound2 libatk1.0-0 libcairo2 libcups2 libfontconfig1 libgdk-pixbuf2.0-0 libgtk-3-0 libnspr4 libpango-1.0-0 libxss1 fonts-liberation libappindicator1 libnss3 lsb-release xdg-utils

#download and install chrome
RUN wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
RUN dpkg -i google-chrome-stable_current_amd64.deb; apt-get -fy install

WORKDIR /distributor

COPY requirements.txt requirements.txt
#Run requrements to user field (/root/.local)
RUN pip3 install -r requirements.txt

#local to copy folder to container
COPY . .
#run
CMD [ "python","-u" ,"./scrapper.py" ]